"""
Studio base
===========

Shared infrastructure for the notebook-based launchers: the embedding shim,
the background-job base class, the individual tab classes, and `_BaseStudio` —
a notebook window that lazily builds fault-isolated tabs. Two launchers subclass
it and pick which tabs to show: `data_extraction.studio.main`
(DataExtractionStudio, the non-EASA tools) and `data_extraction.studio.easa`
(EASAStudio, the EASA extraction/review tabs).

Design goals (intentional):
  * NON-INVASIVE. Each hosted tool keeps its own module and still runs on its
    own (via `python -m ...`). This module only *hosts* them.
  * Each tool stays a self-contained panel in its own tab; their internals are
    never merged together.

How the hosting works
---------------------
The existing UIs (`ExtractionLauncherUI`, `SectionReviewApp`,
`ChunkAIReviewApp`) were written to own a top-level
window: they call ``root.title(...)``, ``root.geometry(...)``,
``root.destroy()`` etc. A notebook tab is a ``Frame``, not a window, so we
hand each tool an ``_EmbeddedRoot`` -- a Frame that quietly no-ops the
window-only calls and treats ``destroy()`` as "do nothing" so a tool 'closing
its window' just leaves its tab in place.

The chunk-review window is normally opened by ``launch_review_app`` creating a
brand new ``tk.Tk()``. With a studio already running that would spawn a second
root + nested mainloop. We monkeypatch ``launch_review_app`` *on the imported
modules only* (files on disk are untouched) to open a modal ``Toplevel`` of the
studio instead, then chain a Section Review modal automatically when chunk
review completes.

There is also a standalone "Section Review" tab. The Section Review editor
lives permanently *inside* that tab: a picker bar stays pinned at the top,
and whenever a valid logged-chunks JSON is selected the embedded editor
below it is (re)loaded for that file. Invalid selections are validated
up-front and simply rejected with a dialog — they never tear down the tab
or the currently loaded review.

Tabs are built lazily on first view, each inside a try/except, so a tool whose
heavy dependencies (docling, markitdown, PyMuPDF, ...) are not installed simply
shows an error panel in its own tab while every other tab keeps working.
"""

import json
import logging
import os
import queue
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

_SENTINEL = object()


# --------------------------------------------------------------------------- #
# Embedding shim                                                              #
# --------------------------------------------------------------------------- #
class _EmbeddedRoot(tk.Frame):
    """A Frame that mimics the subset of Tk root methods the embedded tools
    call, so a window-owning UI class can be hosted inside a notebook tab
    without any modification to its source."""

    def title(self, *args, **kwargs):
        pass

    def geometry(self, *args, **kwargs):
        pass

    def resizable(self, *args, **kwargs):
        pass

    def minsize(self, *args, **kwargs):
        pass

    def maxsize(self, *args, **kwargs):
        pass

    def iconbitmap(self, *args, **kwargs):
        pass

    def protocol(self, *args, **kwargs):
        pass

    def destroy(self):
        # A hosted tool "closing its root" must not tear down its tab; the
        # studio owns the real lifecycle. The frame is still destroyed for real
        # when the parent window closes (Tk cascades at the C level).
        pass


def _make_embedded_host(parent):
    host = _EmbeddedRoot(parent)
    host.pack(fill="both", expand=True)
    return host


class _QueueStream:
    """Minimal stdout replacement that forwards writes onto a queue so a
    background worker's ``print`` output can be shown in the UI safely."""

    def __init__(self, q):
        self._q = q

    def write(self, s):
        if s:
            self._q.put(s)

    def flush(self):
        pass


# --------------------------------------------------------------------------- #
# Base class for the two "run a backend job" tabs (EASA, Markdown)            #
# --------------------------------------------------------------------------- #
class _JobTab:
    """Shared plumbing: a scrolling log + a background worker that captures the
    backend's ``print`` output without freezing the UI."""

    def __init__(self, frame):
        self.frame = frame
        self.run_btn = None  # subclass must assign
        self._q = None

    def _build_log(self):
        ttk.Label(self.frame, text="Log output:").pack(anchor="w", padx=10)
        self.txt = scrolledtext.ScrolledText(self.frame, height=16, wrap="word")
        self.txt.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _log(self, s):
        self.txt.insert("end", s)
        self.txt.see("end")

    def _run_threaded(self, job):
        self._q = queue.Queue()
        if self.run_btn is not None:
            self.run_btn.config(state="disabled")
        threading.Thread(target=self._worker, args=(job,), daemon=True).start()
        self.frame.after(100, self._poll)

    def _worker(self, job):
        # NOTE: stdout is redirected process-wide for the duration of the job.
        # These tools are meant to be run one at a time, so that is acceptable.
        old_stdout = sys.stdout
        sys.stdout = _QueueStream(self._q)
        try:
            job()
            print("\n[COMPLETED]\n")
        except Exception as exc:  # noqa: BLE001 - surfaced to the log panel
            print(f"\n[ERROR] {exc}\n{traceback.format_exc()}\n")
        finally:
            sys.stdout = old_stdout
            self._q.put(_SENTINEL)

    def _poll(self):
        try:
            while True:
                item = self._q.get_nowait()
                if item is _SENTINEL:
                    if self.run_btn is not None:
                        self.run_btn.config(state="normal")
                    return
                self._log(item)
        except queue.Empty:
            pass
        self.frame.after(100, self._poll)


# --------------------------------------------------------------------------- #
# EASA XML extraction tab                                                     #
# --------------------------------------------------------------------------- #
class _EasaTab(_JobTab):
    def __init__(self, frame):
        super().__init__(frame)

        form = ttk.LabelFrame(
            frame,
            text=" EASA XML ZIP  ->  Structured JSON / Master Excel / Images / Tables ",
            padding=10,
        )
        form.pack(fill="x", padx=10, pady=10)

        ttk.Label(form, text="Source (.zip file or a folder of .zip files):").grid(
            row=0, column=0, sticky="w"
        )
        self.ent_src = ttk.Entry(form, width=78)
        self.ent_src.grid(row=0, column=1, padx=5, pady=3)
        ttk.Button(form, text="File...", command=self._browse_file).grid(row=0, column=2)
        ttk.Button(form, text="Folder...", command=self._browse_folder).grid(
            row=0, column=3, padx=(2, 0)
        )

        ttk.Label(form, text="Workspace / storage directory:").grid(
            row=1, column=0, sticky="w"
        )
        self.ent_store = ttk.Entry(form, width=78)
        self.ent_store.grid(row=1, column=1, padx=5, pady=3)
        ttk.Button(form, text="Browse...", command=self._browse_store).grid(row=1, column=2)

        self.var_graph = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            form,
            text="Also build Cosmograph node/edge CSVs from each output JSON",
            variable=self.var_graph,
        ).grid(row=2, column=1, sticky="w", pady=(4, 0))

        self.run_btn = ttk.Button(form, text="Run EASA Extraction", command=self._run)
        self.run_btn.grid(row=3, column=1, sticky="w", pady=8)

        self._build_log()

    def _browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("ZIP archives", "*.zip")])
        if path:
            self.ent_src.delete(0, tk.END)
            self.ent_src.insert(0, path)

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Select a folder containing .zip files")
        if path:
            self.ent_src.delete(0, tk.END)
            self.ent_src.insert(0, path)

    def _browse_store(self):
        path = filedialog.askdirectory(title="Select a workspace directory")
        if path:
            self.ent_store.delete(0, tk.END)
            self.ent_store.insert(0, path)

    def _run(self):
        src = self.ent_src.get().strip()
        store_dir = self.ent_store.get().strip()
        if not src or not store_dir:
            messagebox.showerror("Missing input", "Please provide both a source and a storage directory.")
            return
        if not os.path.exists(src):
            messagebox.showerror("Not found", f"Source path does not exist:\n{src}")
            return

        def job():
            from data_extraction.easa import run_main
            run_main.main(src_path=src, storage_base=store_dir, build_cosmograph=self.var_graph.get())

        self._run_threaded(job)


# --------------------------------------------------------------------------- #
# PDF -> Markdown tab                                                         #
# --------------------------------------------------------------------------- #
class _MarkdownTab(_JobTab):
    def __init__(self, frame):
        super().__init__(frame)

        form = ttk.LabelFrame(
            frame, text=" PDF  ->  Markdown (via MarkItDown) ", padding=10
        )
        form.pack(fill="x", padx=10, pady=10)

        ttk.Label(form, text="Source (.pdf file or a folder of .pdf files):").grid(
            row=0, column=0, sticky="w"
        )
        self.ent_src = ttk.Entry(form, width=78)
        self.ent_src.grid(row=0, column=1, padx=5, pady=3)
        ttk.Button(form, text="File...", command=self._browse_file).grid(row=0, column=2)
        ttk.Button(form, text="Folder...", command=self._browse_folder).grid(
            row=0, column=3, padx=(2, 0)
        )

        ttk.Label(form, text="Output directory (.md files written here):").grid(
            row=1, column=0, sticky="w"
        )
        self.ent_out = ttk.Entry(form, width=78)
        self.ent_out.grid(row=1, column=1, padx=5, pady=3)
        ttk.Button(form, text="Browse...", command=self._browse_out).grid(row=1, column=2)

        self.run_btn = ttk.Button(form, text="Convert to Markdown", command=self._run)
        self.run_btn.grid(row=2, column=1, sticky="w", pady=8)

        self._build_log()

    def _browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("PDF documents", "*.pdf")])
        if path:
            self.ent_src.delete(0, tk.END)
            self.ent_src.insert(0, path)

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Select a folder containing .pdf files")
        if path:
            self.ent_src.delete(0, tk.END)
            self.ent_src.insert(0, path)

    def _browse_out(self):
        path = filedialog.askdirectory(title="Select output directory")
        if path:
            self.ent_out.delete(0, tk.END)
            self.ent_out.insert(0, path)

    def _run(self):
        src = self.ent_src.get().strip()
        out_dir = self.ent_out.get().strip()
        if not src or not out_dir:
            messagebox.showerror("Missing input", "Please provide a source and an output directory.")
            return
        if not os.path.exists(src):
            messagebox.showerror("Not found", f"Source path does not exist:\n{src}")
            return

        def job():
            from data_extraction.markdown import converter as md_mod
            os.makedirs(out_dir, exist_ok=True)

            if os.path.isdir(src):
                pdfs = sorted(str(p) for p in Path(src).rglob("*.pdf"))
            elif src.lower().endswith(".pdf"):
                pdfs = [src]
            else:
                pdfs = []

            if not pdfs:
                print("No .pdf files found at the source path.")
                return

            print(f"Found {len(pdfs)} PDF(s) to convert.")
            for pdf_path in pdfs:
                out_md = os.path.join(out_dir, Path(pdf_path).stem + ".md")
                md_mod.convert_pdf_to_markdown(pdf_path, out_md)

        self._run_threaded(job)


# --------------------------------------------------------------------------- #
# Standalone Section Review tab                                              #
# --------------------------------------------------------------------------- #
class _SectionReviewTab:
    """
    Hosts the Section Review editor permanently inside the tab.

    Layout:
      * A picker bar pinned at the top (entry + Browse + Load/Reload) that
        never goes away.
      * A body area below it that holds either a hint placeholder (nothing
        loaded yet) or an embedded `SectionReviewApp` for the selected file.

    Selecting a different valid logged-chunks JSON swaps the embedded editor
    in place. Every candidate file is validated *before* the current editor
    is touched, so a missing/corrupt/empty file only produces an error dialog
    and leaves the tab — and whatever review is currently open — fully intact.
    Because the editor runs on an `_EmbeddedRoot`, its own "close the window"
    paths (Export & Finish, Exit) are no-ops and the editor simply stays open
    in the tab, which is the intended behavior here.
    """

    def __init__(self, frame):
        self.frame = frame
        self._app = None          # current SectionReviewApp instance (or None)
        self._host = None         # its _EmbeddedRoot host frame (or None)
        self._loaded_path = None  # path currently loaded into the editor

        self._logger = logging.getLogger("SectionReviewTab")
        self._logger.setLevel(logging.INFO)
        if not self._logger.handlers:
            self._logger.addHandler(logging.StreamHandler())

        self._build_picker()

        # Body container: holds either the placeholder or the embedded app.
        self.body = ttk.Frame(self.frame)
        self.body.pack(fill="both", expand=True)
        self._show_placeholder()

    # -- persistent picker bar ---------------------------------------------- #
    def _build_picker(self):
        form = ttk.LabelFrame(
            self.frame,
            text=" Logged-Chunks JSON — Section Review opens below ",
            padding=10,
        )
        form.pack(fill="x", padx=10, pady=(10, 4))

        ttk.Label(
            form,
            text='Logged Chunks Output JSON (must contain "merged_headings"):',
        ).grid(row=0, column=0, sticky="w")

        self.ent_output = ttk.Entry(form, width=78)
        self.ent_output.grid(row=1, column=0, padx=(0, 5), pady=3, sticky="we")
        ttk.Button(form, text="Browse...", command=self._browse).grid(row=1, column=1)
        self.btn_load = ttk.Button(form, text="Load / Reload", command=self._load_clicked)
        self.btn_load.grid(row=1, column=2, padx=(4, 0))
        form.columnconfigure(0, weight=1)

        self.lbl_status = ttk.Label(form, text="No file loaded.", foreground="#666666")
        self.lbl_status.grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select logged chunks / sections JSON file",
            filetypes=[("JSON Files", "*.json")],
        )
        if path:
            self.ent_output.delete(0, tk.END)
            self.ent_output.insert(0, path)
            # Browsing to a file loads it immediately — no extra click needed.
            self._load(path)

    def _load_clicked(self):
        self._load(self.ent_output.get().strip())

    # -- body management ------------------------------------------------------ #
    def _show_placeholder(self):
        self._clear_body()
        ttk.Label(
            self.body,
            text=(
                "Select a logged-chunks output JSON above — Section Review will "
                "open right here and update whenever you pick a different file.\n\n"
                "Tip: this is the same output file chunk review writes to, e.g.\n"
                "  <storage>/<doc_name>/<date>/<hash>_docling_logged_chunks.json"
            ),
            foreground="#666666",
            padding=20,
            justify="left",
        ).pack(anchor="w")

    def _clear_body(self):
        for child in self.body.winfo_children():
            try:
                # Force a *real* destroy even for _EmbeddedRoot, whose own
                # destroy() is intentionally a no-op.
                tk.Frame.destroy(child)
            except Exception:  # noqa: BLE001 - best-effort teardown
                try:
                    child.destroy()
                except Exception:  # noqa: BLE001
                    pass
        self._app = None
        self._host = None

    # -- loading -------------------------------------------------------------- #
    def _load(self, path):
        """Validate `path` and (re)load the embedded Section Review editor.

        All failure paths return *before* the currently loaded editor is
        touched, so a bad selection can never blank or crash the tab.
        """
        if not path:
            messagebox.showerror("Error", "Please select an output JSON file.")
            return
        if not os.path.isfile(path):
            messagebox.showerror("File Not Found", f"Output file not found:\n{path}")
            return

        # Pre-validate the file ourselves so SectionReviewApp's internal
        # error handling (which destroys its root) is never triggered.
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except json.JSONDecodeError as exc:
            messagebox.showerror("JSON Error", f"Failed to parse output file:\n{exc}")
            return
        except OSError as exc:
            messagebox.showerror("Error", f"Failed to read output file:\n{exc}")
            return

        sections = data.get("merged_headings") if isinstance(data, dict) else None
        if not sections:
            messagebox.showwarning(
                "No Sections Found",
                "No merged sections were found in the selected file.\n"
                'It must be a chunk-review output JSON containing "merged_headings".',
            )
            return

        # Import lazily so a missing module only affects loading, not the tab.
        try:
            from data_extraction.chunking.section_review_ui import SectionReviewApp
        except Exception as exc:  # noqa: BLE001 - surfaced to the user
            messagebox.showerror(
                "Section Review Unavailable",
                f"Could not import Section Review:\n\n{type(exc).__name__}: {exc}",
            )
            return

        # Protect unsaved edits in the currently open review before swapping.
        if self._app is not None and getattr(self._app, "unsaved_changes", False):
            if not messagebox.askyesno(
                "Unsaved Changes",
                "The currently open review has unsaved changes.\n\n"
                "Discard them and load the selected file?",
            ):
                return

        # Swap: tear down the old editor (or placeholder) and embed a new one.
        self._clear_body()
        host = _make_embedded_host(self.body)
        try:
            app = SectionReviewApp(
                root=host, output_file_path=path, logger=self._logger
            )
        except Exception as exc:  # noqa: BLE001 - surface into the tab itself
            try:
                tk.Frame.destroy(host)
            except Exception:  # noqa: BLE001
                pass
            self._show_placeholder()
            self.lbl_status.config(text="Failed to load — see error dialog.")
            messagebox.showerror(
                "Section Review Error",
                "Section Review could not be loaded.\n\n"
                f"{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}",
            )
            return

        # Belt-and-suspenders: if the app bailed out of __init__ early (its
        # loader returned False), it never built its widgets. Fall back to
        # the placeholder instead of leaving an empty frame behind.
        if not hasattr(app, "status_label"):
            try:
                tk.Frame.destroy(host)
            except Exception:  # noqa: BLE001
                pass
            self._show_placeholder()
            self.lbl_status.config(text="Could not load sections from the selected file.")
            return

        self._app = app
        self._host = host
        self._loaded_path = path
        self.lbl_status.config(text=f"Loaded: {path}")
        self._logger.info("Section Review loaded in-tab for: %s", path)


# --------------------------------------------------------------------------- #
# Base studio: shared notebook window, lazy fault-isolated tabs, and every     #
# tab builder — reused by the main studio and the EASA studio.                 #
# --------------------------------------------------------------------------- #
class _BaseStudio:
    """A notebook window that lazily builds fault-isolated tabs.

    Subclasses set WINDOW_TITLE / HEADER / GEOMETRY / MINSIZE and TAB_SPECS —
    a list of (label, builder-method-name) pairs. Every tab builder lives on
    this base class, so multiple launchers (the main studio, the EASA studio)
    can compose their own subset of tabs without duplicating UI plumbing.
    """

    WINDOW_TITLE = "Studio"
    HEADER = ""
    GEOMETRY = "1320x880"
    MINSIZE = (1000, 640)
    TAB_SPECS = []  # [(label, builder_method_name), ...] — set by subclasses

    def __init__(self, root):
        self.root = root
        root.title(self.WINDOW_TITLE)
        root.geometry(self.GEOMETRY)
        root.minsize(*self.MINSIZE)

        # Apply the clam ttk theme for the whole studio right at startup.
        # Previously only the AI Review page set it (review_panel), so the
        # app looked different until that tab was first opened; tabs load
        # lazily, so the theme must be set here, not in a tab builder.
        try:
            style = ttk.Style(root)
            if "clam" in style.theme_names():
                style.theme_use("clam")
        except tk.TclError:
            pass

        self._review_launcher = self._make_review_launcher()

        if self.HEADER:
            ttk.Label(
                root,
                text=self.HEADER,
                font=("TkDefaultFont", 10, "bold"),
                padding=(10, 6),
            ).pack(fill="x")

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True, padx=6, pady=6)

        self._tabs = []
        for label, builder_name in self.TAB_SPECS:
            self._add_tab(label, getattr(self, builder_name))

        self.nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        # Build the initially-selected tab once the loop is running.
        self.root.after(50, self._on_tab_changed)

    # -- lazy tab machinery -------------------------------------------------- #
    def _add_tab(self, title, builder):
        frame = ttk.Frame(self.nb)
        self.nb.add(frame, text=title)
        self._tabs.append({"frame": frame, "builder": builder, "built": False})

    def _on_tab_changed(self, event=None):
        try:
            idx = self.nb.index(self.nb.select())
        except tk.TclError:
            return
        tab = self._tabs[idx]
        if tab["built"]:
            return
        tab["built"] = True
        try:
            tab["builder"](tab["frame"])
        except Exception as exc:  # noqa: BLE001 - render into the failing tab
            self._render_error(tab["frame"], exc)

    def _render_error(self, frame, err):
        for child in frame.winfo_children():
            try:
                tk.Frame.destroy(child)  # force real destroy even for _EmbeddedRoot
            except Exception:  # noqa: BLE001
                pass
        message = (
            "This tool could not be loaded.\n\n"
            "This usually means a required Python package or a source file is "
            "missing on this machine. The other tabs are unaffected.\n\n"
            f"{type(err).__name__}: {err}\n\n"
            f"{traceback.format_exc()}"
        )
        box = scrolledtext.ScrolledText(frame, wrap="word")
        box.insert("1.0", message)
        box.config(state="disabled")
        box.pack(fill="both", expand=True, padx=10, pady=10)

    # -- review-window launcher patch --------------------------------------- #
    def _make_review_launcher(self):
        """
        Creates a patched launch_review_app function that:
        1. Opens the bulk chunk-triage review in a Toplevel modal window
           (chunks are sorted automatically first — see
           ``data_extraction.chunking.chunk_triage``).
        2. When the review is accepted, auto-launches Section Review in a
           second Toplevel modal window.
        Used by the PDF Extraction tab (both its pipeline and its
        review-from-cache flow).
        """
        studio_root = self.root

        def _launch(chunks_data, logged_chunks, processed_indices, output_file,
                    logger, bulk=True):
            from data_extraction.chunking.chunk_triage_ui import ChunkTriageApp
            from data_extraction.chunking.chunk_review_ui import ChunkReviewApp
            from data_extraction.chunking.logic import _triage_llm
            from data_extraction.chunking.section_review_ui import SectionReviewApp

            def on_chunk_review_complete():
                logger.info("Chunk review completed. Launching section review...")
                section_win = tk.Toplevel(studio_root)
                section_win.transient(studio_root)
                SectionReviewApp(
                    root=section_win,
                    output_file_path=output_file,
                    logger=logger,
                )
                section_win.grab_set()
                studio_root.wait_window(section_win)

            chunk_win = tk.Toplevel(studio_root)
            chunk_win.transient(studio_root)
            if bulk:
                ChunkTriageApp(
                    root=chunk_win,
                    chunks_data=chunks_data,
                    output_file_name=output_file,
                    logger=logger,
                    on_complete_callback=on_chunk_review_complete,
                    llm=_triage_llm(logger),
                    prior_history=logged_chunks,   # resume pre-applies decisions
                )
            else:
                ChunkReviewApp(
                    root=chunk_win,
                    chunks_data=chunks_data,
                    logged_chunks=logged_chunks,
                    processed_indices=processed_indices,
                    output_file_name=output_file,
                    logger=logger,
                    on_complete_callback=on_chunk_review_complete,
                )
            # If there was nothing to review (e.g. resuming an already-completed
            # document, or an empty chunk set), the review app may finish inside
            # its own __init__ and destroy chunk_win — and its on_complete
            # callback has already run Section Review synchronously. Only
            # grab/wait when the window is still alive, otherwise grab_set()
            # would raise a TclError.
            if chunk_win.winfo_exists():
                chunk_win.grab_set()
                studio_root.wait_window(chunk_win)

        return _launch

    # -- tab builders -------------------------------------------------------- #
    def _build_extraction_tab(self, frame):
        from data_extraction.chunking import logic as chunk_logic
        chunk_logic.launch_review_app = self._review_launcher  # patch module hook
        host = _make_embedded_host(frame)
        chunk_logic.ExtractionLauncherUI(host)

    def _build_chunk_ai_tab(self, frame):
        from data_extraction.chunking.ai_review_ui import ChunkAIReviewApp
        host = _make_embedded_host(frame)
        ChunkAIReviewApp(host)

    def _build_section_review_tab(self, frame):
        _SectionReviewTab(frame)

    def _build_easa_tab(self, frame):
        _EasaTab(frame)

    def _build_easa_review_tab(self, frame):
        from data_extraction.easa.json_review_ui import EASAJsonReviewApp
        host = _make_embedded_host(frame)
        EASAJsonReviewApp(host)

    def _build_markdown_tab(self, frame):
        _MarkdownTab(frame)

    def _build_data_analysis_tab(self, frame):
        from data_extraction.studio.data_analysis_tab import DataAnalysisTab
        DataAnalysisTab(frame)