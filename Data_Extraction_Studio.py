"""
Data Extraction Studio
======================

A single-window launcher that hosts every existing extraction/curation tool in
this repository under one `ttk.Notebook`, one tab per tool.

Design goals (intentional):
  * NON-INVASIVE. None of the existing tool files are modified. Each tool keeps
    its own class / module and continues to run standalone via its own
    ``__main__`` block. This file only *hosts* them.
  * Each tool stays a self-contained panel in its own tab; their internals are
    never merged together.

How the hosting works
---------------------
The existing UIs (`ExtractionLauncherUI`, `CacheReviewLauncher`,
`SectionRewriterUI`) were written to own a top-level window: they call
``root.title(...)``, ``root.geometry(...)``, ``root.destroy()`` etc. A notebook
tab is a ``Frame``, not a window, so we hand each tool an ``_EmbeddedRoot`` -- a
Frame that quietly no-ops the window-only calls and treats ``destroy()`` as
"do nothing" so a tool 'closing its window' just leaves its tab in place.

The chunk-review window is normally opened by ``launch_review_app`` creating a
brand new ``tk.Tk()``. With a studio already running that would spawn a second
root + nested mainloop. We monkeypatch ``launch_review_app`` *on the imported
modules only* (files on disk are untouched) to open a modal ``Toplevel`` of the
studio instead.

Tabs are built lazily on first view, each inside a try/except, so a tool whose
heavy dependencies (docling, markitdown, PyMuPDF, ...) are not installed simply
shows an error panel in its own tab while every other tab keeps working.
"""

import importlib
import importlib.util
import os
import queue
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

BASE = Path(__file__).resolve().parent

# Make the tool packages importable without touching them.
for _sub in ("Manual_Chunking_With_UI", "EASA_Data_Extractors", "Markdown_extraction"):
    _p = str(BASE / _sub)
    if _p not in sys.path:
        sys.path.insert(0, _p)

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
        path = filedialog.askdirectory(title="Select workspace / storage directory")
        if path:
            self.ent_store.delete(0, tk.END)
            self.ent_store.insert(0, path)

    def _run(self):
        src = self.ent_src.get().strip()
        store = self.ent_store.get().strip()
        if not src or not store:
            messagebox.showerror("Missing input", "Please provide a source and a storage directory.")
            return
        if not os.path.exists(src):
            messagebox.showerror("Not found", f"Source path does not exist:\n{src}")
            return

        build_graph = self.var_graph.get()

        def job():
            from EASA_Parser import extract_easa_from_zip_v3, resolve_paths
            from EASA_Graph_builder import export_to_cosmograph_csv

            if os.path.isdir(src):
                targets = sorted(str(p) for p in Path(src).rglob("*.zip"))
            elif src.lower().endswith(".zip"):
                targets = [src]
            else:
                targets = []

            if not targets:
                print("No .zip files found at the source path.")
                return

            print(f"Found {len(targets)} archive(s) to process.")
            for zip_path in targets:
                print(f"\n=== Processing: {zip_path} ===")
                extract_easa_from_zip_v3(zip_path, store)
                if build_graph:
                    try:
                        paths = resolve_paths(store, zip_path)
                        export_to_cosmograph_csv(paths["output_json"])
                    except Exception as exc:  # noqa: BLE001
                        print(f"[graph] Skipped graph build for {zip_path}: {exc}")

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

    @staticmethod
    def _load_md_module():
        # The source filename contains a space, so it cannot be imported by name.
        path = BASE / "Markdown_extraction" / "Process_files_to _Markdown.py"
        spec = importlib.util.spec_from_file_location("process_files_to_markdown", str(path))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod

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
            md_mod = self._load_md_module()
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
# Main studio window                                                         #
# --------------------------------------------------------------------------- #
class DataExtractionStudio:
    def __init__(self, root):
        self.root = root
        root.title("Data Extraction Studio")
        root.geometry("1320x880")
        root.minsize(1000, 640)

        self._review_launcher = self._make_review_launcher()

        header = ttk.Label(
            root,
            text="Data Extraction Studio — every tool in one place. Pick a tab to begin.",
            font=("TkDefaultFont", 10, "bold"),
            padding=(10, 6),
        )
        header.pack(fill="x")

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True, padx=6, pady=6)

        self._tabs = []
        self._add_tab("PDF Extraction & Review", self._build_extraction_tab)
        self._add_tab("Cache Review Launcher", self._build_cache_tab)
        self._add_tab("Section Re-Writer", self._build_rewriter_tab)
        self._add_tab("EASA XML Extraction", self._build_easa_tab)
        self._add_tab("PDF -> Markdown", self._build_markdown_tab)

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
        studio_root = self.root

        def _launch(chunks_data, logged_chunks, processed_indices, output_file, logger):
            from chunk_review_ui import ChunkReviewApp

            win = tk.Toplevel(studio_root)
            win.transient(studio_root)
            ChunkReviewApp(
                win, chunks_data, logged_chunks, processed_indices, output_file, logger
            )
            win.grab_set()
            studio_root.wait_window(win)

        return _launch

    # -- tab builders -------------------------------------------------------- #
    def _build_extraction_tab(self, frame):
        chunk_logic = importlib.import_module("Chunk_review_logic")
        chunk_logic.launch_review_app = self._review_launcher  # patch import only
        host = _make_embedded_host(frame)
        chunk_logic.ExtractionLauncherUI(host)

    def _build_cache_tab(self, frame):
        cache_mod = importlib.import_module("cache_launcher")
        cache_mod.launch_review_app = self._review_launcher  # patch import only
        host = _make_embedded_host(frame)
        cache_mod.CacheReviewLauncher(host)

    def _build_rewriter_tab(self, frame):
        rewriter_mod = importlib.import_module("Raw_sec_rewriter")
        host = _make_embedded_host(frame)
        rewriter_mod.SectionRewriterUI(host)

    def _build_easa_tab(self, frame):
        _EasaTab(frame)

    def _build_markdown_tab(self, frame):
        _MarkdownTab(frame)


def main():
    root = tk.Tk()
    DataExtractionStudio(root)
    root.mainloop()


if __name__ == "__main__":
    main()
