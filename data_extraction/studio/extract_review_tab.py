"""
Unified "Extract & Review" tab.

Realises the approved mockup: one workspace-driven surface where PDF (Docling
chunks) and EASA (XML) extraction and review live together, instead of the
separate Extraction / Section-Review / AI-Review / EASA tabs.

It is an *orchestrator*, not a rewrite — it reuses the existing, unmodified
tools by hosting them inside a review pane, driven by three shared pieces:

  * a **source selector** (PDF · Docling chunks / EASA · XML),
  * a **workspace bar** (the storage folder both extraction and review share,
    backed by the SQLite store), and
  * a **document rail** listing what the workspace already holds (from the DB).

Picking "Extract new" hosts the source's extractor (workspace pre-filled);
picking a document in the rail hosts its review tool, with mode buttons to
switch between the review surfaces the source offers. Every hosted tool keeps
all of its own functionality; nothing here removes a feature.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from data_extraction.studio.base import _make_embedded_host


class ExtractReviewTab:
    #: source code -> (label, [(mode label, host-method name), ...])
    _MODES = {
        "pdf": [("Section review", "_host_section_review"),
                ("AI review", "_host_pdf_ai_review")],
        "easa": [("Browse + AI review", "_host_easa_review")],
    }
    _SOURCE_LABELS = {"pdf": "PDF · Docling chunks", "easa": "EASA · XML"}

    def __init__(self, frame, review_launcher=None):
        self.frame = frame
        self.review_launcher = review_launcher   # studio's patched launch_review_app
        self.source = "pdf"
        self.workspace = None          # facade.Workspace once opened
        self.workspace_dir = None
        self._row_meta = {}            # rail iid -> document dict
        self._selected = None          # selected document dict
        self._host_frame = None        # current embedded tool host frame
        self._build()

    # ------------------------------------------------------------------ UI -- #
    def _build(self):
        # header: source selector + workspace bar
        head = ttk.Frame(self.frame, padding=(10, 8))
        head.pack(fill="x")
        ttk.Label(head, text="Source:").pack(side="left")
        self.src_pdf = ttk.Button(head, text="PDF · Docling chunks",
                                  command=lambda: self._set_source("pdf"))
        self.src_pdf.pack(side="left", padx=(6, 2))
        self.src_easa = ttk.Button(head, text="EASA · XML",
                                   command=lambda: self._set_source("easa"))
        self.src_easa.pack(side="left")

        wsbar = ttk.Frame(self.frame, padding=(10, 0))
        wsbar.pack(fill="x")
        ttk.Label(wsbar, text="Workspace:").pack(side="left")
        self.ent_ws = ttk.Entry(wsbar)
        self.ent_ws.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(wsbar, text="Browse…", command=self._browse_ws).pack(side="left")
        ttk.Button(wsbar, text="Open", command=self._open_ws_clicked).pack(side="left", padx=(4, 0))
        self.ws_hint = ttk.Label(wsbar, foreground="#666666", text="")
        self.ws_hint.pack(side="left", padx=(8, 0))

        body = ttk.PanedWindow(self.frame, orient="horizontal")
        body.pack(fill="both", expand=True, padx=10, pady=(6, 8))

        # left rail
        rail = ttk.LabelFrame(body, text=" Documents ", padding=4)
        body.add(rail, weight=1)
        ttk.Button(rail, text="＋ Extract new…",
                   command=self._show_extractor).pack(fill="x", pady=(0, 4))
        sr = ttk.Frame(rail)
        sr.pack(fill="x")
        ttk.Label(sr, text="🔎").pack(side="left")
        self.filter = ttk.Entry(sr)
        self.filter.pack(side="left", fill="x", expand=True, padx=4)
        self.filter.bind("<KeyRelease>", lambda _e: self._refresh_rail())
        ttk.Button(sr, text="↻", width=3, command=self._refresh_rail).pack(side="left")
        self.rail = ttk.Treeview(rail, columns=("status",), show="tree headings",
                                 selectmode="browse", height=14)
        self.rail.heading("#0", text="Document")
        self.rail.heading("status", text="Status")
        self.rail.column("#0", width=180, anchor="w")
        self.rail.column("status", width=76, anchor="center", stretch=False)
        self.rail.pack(fill="both", expand=True, pady=(4, 0))
        self.rail.bind("<<TreeviewSelect>>", self._on_rail_select)

        # right: mode bar + host container
        right = ttk.Frame(body)
        body.add(right, weight=3)
        self.mode_bar = ttk.Frame(right)
        self.mode_bar.pack(fill="x")
        self.host_wrap = ttk.Frame(right)
        self.host_wrap.pack(fill="both", expand=True, pady=(6, 0))

        self._set_source("pdf")
        self._show_placeholder()

    # --------------------------------------------------------------- source -- #
    def _set_source(self, source):
        self.source = source
        for btn, code in ((self.src_pdf, "pdf"), (self.src_easa, "easa")):
            # a pressed look via state; ttk buttons don't toggle, so use text marker
            btn.configure(text=("● " if code == source else "○ ")
                          + self._SOURCE_LABELS[code])
        self._selected = None
        self._refresh_rail()
        self._build_mode_bar()
        self._show_placeholder()

    def _build_mode_bar(self):
        for w in self.mode_bar.winfo_children():
            w.destroy()
        ttk.Label(self.mode_bar, text="View:").pack(side="left")
        self._mode_btns = {}
        for label, method in self._MODES[self.source]:
            b = ttk.Button(self.mode_bar, text=label,
                           command=lambda m=method: self._host_mode(m))
            b.pack(side="left", padx=(6, 0))
            b.configure(state="disabled")   # enabled once a doc is selected
            self._mode_btns[method] = b

    # ------------------------------------------------------------ workspace -- #
    def _browse_ws(self):
        path = filedialog.askdirectory(title="Select the workspace / storage folder")
        if path:
            self.ent_ws.delete(0, tk.END)
            self.ent_ws.insert(0, path)
            self._open_workspace(path)

    def _open_ws_clicked(self):
        self._open_workspace(self.ent_ws.get().strip())

    def _open_workspace(self, path):
        if not path or not os.path.isdir(path):
            messagebox.showinfo("Pick a folder", "Choose an existing workspace / "
                                "storage folder.")
            return
        try:
            from data_extraction.db import facade
            self.workspace = facade.open_workspace(path)
        except Exception as exc:  # noqa: BLE001
            self.workspace = None
            self.ws_hint.config(text=f"database unavailable: {exc}")
        self.workspace_dir = path
        self._refresh_rail()

    def _refresh_rail(self):
        self.rail.delete(*self.rail.get_children())
        self._row_meta.clear()
        if self.workspace is None:
            self.ws_hint.config(text="no workspace open — extract or open one")
            return
        needle = (self.filter.get() or "").strip().lower()
        docs = [d for d in self.workspace.store.list_documents(doc_type=self.source)
                if needle in (d.get("doc_name") or "").lower()]
        for d in docs:
            iid = self.rail.insert("", "end", text=d.get("doc_name") or d["doc_uuid"],
                                   values=(d.get("status") or "",))
            self._row_meta[iid] = d
        self.ws_hint.config(text=f"{len(docs)} {self.source.upper()} document(s)")

    # --------------------------------------------------------------- hosting -- #
    def _clear_host(self):
        for child in self.host_wrap.winfo_children():
            try:
                tk.Frame.destroy(child)   # force real destroy even for _EmbeddedRoot
            except Exception:  # noqa: BLE001
                try:
                    child.destroy()
                except Exception:  # noqa: BLE001
                    pass
        self._host_frame = None

    def _show_placeholder(self):
        self._clear_host()
        msg = ("Open a workspace and pick a document on the left to review it, "
               "or press “＋ Extract new…” to add one.\n\n"
               "The workspace is the storage folder your extractions and reviews "
               "share; everything in it is listed on the left.")
        ttk.Label(self.host_wrap, text=msg, foreground="#666666",
                  justify="left", padding=16, wraplength=560).pack(anchor="nw")

    def _host(self, builder):
        """Swap a fresh embedded tool into the review pane. ``builder(host)``
        constructs the tool on the given host frame. Errors render in-pane so
        one broken tool never blanks the tab."""
        self._clear_host()
        host = _make_embedded_host(self.host_wrap)
        try:
            builder(host)
            self._host_frame = host
        except Exception as exc:  # noqa: BLE001
            import traceback
            try:
                tk.Frame.destroy(host)
            except Exception:  # noqa: BLE001
                pass
            box = tk.Text(self.host_wrap, wrap="word")
            box.insert("1.0", f"This tool could not be loaded.\n\n"
                       f"{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}")
            box.config(state="disabled")
            box.pack(fill="both", expand=True)

    def _show_extractor(self):
        """Host the source's extractor with the workspace pre-filled."""
        self._selected = None
        self.rail.selection_remove(self.rail.selection())
        for b in getattr(self, "_mode_btns", {}).values():
            b.configure(state="disabled")
        if self.source == "pdf":
            self._host(self._build_pdf_extractor)
        else:
            self._host(self._build_easa_extractor)

    def _build_pdf_extractor(self, host):
        from data_extraction.chunking import logic as chunk_logic
        if self.review_launcher is not None:
            chunk_logic.launch_review_app = self.review_launcher
        app = chunk_logic.ExtractionLauncherUI(host)
        # Pre-fill the storage destination with the workspace and refresh the
        # rail whenever a review is saved there.
        if self.workspace_dir:
            try:
                app.entry_store.delete(0, tk.END)
                app.entry_store.insert(0, self.workspace_dir)
            except Exception:  # noqa: BLE001
                pass
        orig = app._open_section_review
        def wrapped(saved_path=None):
            orig(saved_path)
            self._after_extract()
        app._open_section_review = wrapped

    def _build_easa_extractor(self, host):
        from data_extraction.easa.extraction_ui import EASAExtractionApp
        app = EASAExtractionApp(host)
        if self.workspace_dir:
            try:
                app.ent_store.delete(0, tk.END)
                app.ent_store.insert(0, self.workspace_dir)
            except Exception:  # noqa: BLE001
                pass

    def _after_extract(self):
        """Called after an extraction/review saved into the workspace: reopen
        the store (new rows) and refresh the rail."""
        if self.workspace_dir:
            self._open_workspace(self.workspace_dir)

    # --------------------------------------------------------------- review -- #
    def _on_rail_select(self, _event=None):
        sel = self.rail.selection()
        if not sel:
            return
        self._selected = self._row_meta.get(sel[0])
        for b in getattr(self, "_mode_btns", {}).values():
            b.configure(state="normal")
        # default to the source's first review mode
        self._host_mode(self._MODES[self.source][0][1])

    def _host_mode(self, method):
        for m, b in getattr(self, "_mode_btns", {}).items():
            try:
                b.state(["pressed"] if m == method else ["!pressed"])
            except tk.TclError:
                pass
        getattr(self, method)()

    def _review_path(self):
        doc = self._selected or {}
        return doc.get("review_path")

    def _host_section_review(self):
        path = self._review_path()
        if not path or not os.path.exists(path):
            self._missing_file(path)
            return
        import logging
        self._host(lambda host: self._import_section_review()(
            root=host, output_file_path=path,
            logger=logging.getLogger("ExtractReviewTab")))

    def _host_pdf_ai_review(self):
        path = self._review_path()
        if not path or not os.path.exists(path):
            self._missing_file(path)
            return
        from data_extraction.chunking.ai_review_ui import ChunkAIReviewApp
        self._host(lambda host: ChunkAIReviewApp(host, json_path=path))

    def _host_easa_review(self):
        path = self._review_path()
        if not path or not os.path.exists(path):
            self._missing_file(path)
            return
        import logging
        from data_extraction.easa.json_review_ui import EASAJsonReviewApp
        self._host(lambda host: EASAJsonReviewApp(
            host, json_path=path, logger=logging.getLogger("ExtractReviewTab")))

    @staticmethod
    def _import_section_review():
        from data_extraction.chunking.section_review_ui import SectionReviewApp
        return SectionReviewApp

    def _missing_file(self, path):
        self._clear_host()
        msg = ("The stored review file for this document was not found"
               + (f":\n{path}" if path else ".")
               + "\n\nRe-extract or re-review the document to regenerate it.")
        ttk.Label(self.host_wrap, text=msg, foreground="#666666",
                  justify="left", padding=16, wraplength=560).pack(anchor="nw")
