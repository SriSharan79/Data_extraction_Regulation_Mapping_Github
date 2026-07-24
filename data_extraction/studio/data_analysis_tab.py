"""
Data & Analysis tab.

A read-across view of everything the extraction and AI-review flows have
persisted to SQLite (see :mod:`data_extraction.db`), plus a place to run and
store cross-document AI analysis.

It is additive: it never touches the extraction/review tabs, it only *reads*
what they wrote (and appends analysis results). Hosted by both studios via
``_BaseStudio._build_data_analysis_tab``.

Layout
------
* A workspace bar: open one workspace's database, or scan the global registry
  to list every document across all workspaces.
* Left: the document list (name · type · status · sections).
* Right: a detail notebook — Sections / AI Reviews / Entities for the selected
  document — and, below it, a Cross-document Analysis panel (scope + prompt +
  gated LLM picker → answer, stored to ``analysis_results``).
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


class DataAnalysisTab:
    def __init__(self, frame):
        self.frame = frame
        self.store = None          # ExtractionStore of the loaded workspace
        self.workspace_dir = None
        self._llm_avail = {}       # {label: [models]} from a gated probe
        self._row_meta = {}        # tree iid -> (db_path, doc_uuid)
        self._build()

    # ------------------------------------------------------------------ UI -- #
    def _build(self):
        bar = ttk.LabelFrame(self.frame, text=" Workspace database ", padding=8)
        bar.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Label(bar, text="Workspace folder:").grid(row=0, column=0, sticky="w")
        self.ent_ws = ttk.Entry(bar, width=70)
        self.ent_ws.grid(row=0, column=1, padx=5, sticky="we")
        ttk.Button(bar, text="Browse…", command=self._browse_ws).grid(row=0, column=2)
        ttk.Button(bar, text="Open", command=self._open_ws_clicked).grid(row=0, column=3, padx=(4, 0))
        ttk.Button(bar, text="Scan all workspaces",
                   command=self._scan_all).grid(row=0, column=4, padx=(8, 0))
        self.lbl_ws = ttk.Label(bar, foreground="#666666", text="No database loaded.")
        self.lbl_ws.grid(row=1, column=0, columnspan=5, sticky="w", pady=(4, 0))
        bar.columnconfigure(1, weight=1)

        body = ttk.PanedWindow(self.frame, orient="horizontal")
        body.pack(fill="both", expand=True, padx=10, pady=(0, 4))

        # left: document list
        left = ttk.LabelFrame(body, text=" Documents ", padding=4)
        body.add(left, weight=1)
        cols = ("type", "status", "secs")
        self.docs = ttk.Treeview(left, columns=cols, show="tree headings",
                                 selectmode="browse", height=12)
        self.docs.heading("#0", text="Document")
        self.docs.heading("type", text="Type")
        self.docs.heading("status", text="Status")
        self.docs.heading("secs", text="Sections")
        self.docs.column("#0", width=200, anchor="w")
        self.docs.column("type", width=60, anchor="center", stretch=False)
        self.docs.column("status", width=80, anchor="center", stretch=False)
        self.docs.column("secs", width=70, anchor="e", stretch=False)
        ds = ttk.Scrollbar(left, orient="vertical", command=self.docs.yview)
        self.docs.configure(yscrollcommand=ds.set)
        ds.pack(side="right", fill="y")
        self.docs.pack(side="left", fill="both", expand=True)
        self.docs.bind("<<TreeviewSelect>>", self._on_doc_select)

        # right: detail notebook + analysis panel
        right = ttk.Frame(body)
        body.add(right, weight=2)

        self.detail_nb = ttk.Notebook(right)
        self.detail_nb.pack(fill="both", expand=True)
        self.tab_sections = self._make_tree_tab(
            ("page", "decision", "text"),
            ("Heading", "Page", "Decision", "Text"))
        self.detail_nb.add(self.tab_sections["frame"], text="Sections")
        self.tab_reviews = self._make_tree_tab(
            ("service", "model", "fallback", "values"),
            ("Section", "Service", "Model", "Fallback", "Columns"))
        self.detail_nb.add(self.tab_reviews["frame"], text="AI Reviews")
        self.tab_entities = self._make_tree_tab(
            ("reference", "process", "chain"),
            ("Section", "Reference", "Process", "Chain"))
        self.detail_nb.add(self.tab_entities["frame"], text="Entities")

        self._build_analysis_panel(right)

    def _make_tree_tab(self, cols, headings):
        frame = ttk.Frame(self.detail_nb)
        tv = ttk.Treeview(frame, columns=cols, show="tree headings", height=8)
        tv.heading("#0", text=headings[0])
        tv.column("#0", width=200, anchor="w")
        for c, h in zip(cols, headings[1:]):
            tv.heading(c, text=h)
            tv.column(c, width=110, anchor="w")
        sc = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sc.set)
        sc.pack(side="right", fill="y")
        tv.pack(side="left", fill="both", expand=True)
        return {"frame": frame, "tv": tv}

    def _build_analysis_panel(self, parent):
        panel = ttk.LabelFrame(parent, text=" Cross-document AI analysis ", padding=8)
        panel.pack(fill="x", pady=(6, 0))

        row = ttk.Frame(panel)
        row.pack(fill="x")
        ttk.Label(row, text="Scope:").pack(side="left")
        self.scope = ttk.Combobox(row, state="readonly", width=22,
                                  values=["All documents", "PDF documents only",
                                          "EASA documents only"])
        self.scope.set("All documents")
        self.scope.pack(side="left", padx=(4, 12))
        ttk.Label(row, text="LLM:").pack(side="left")
        self.llm_service = ttk.Combobox(row, state="readonly", width=11, values=[])
        self.llm_service.pack(side="left", padx=4)
        self.llm_service.bind("<<ComboboxSelected>>", self._llm_service_picked)
        self.llm_model = ttk.Combobox(row, state="readonly", width=30, values=[])
        self.llm_model.pack(side="left", padx=4)
        ttk.Button(row, text="↻", width=3,
                   command=lambda: self._probe_llm(force=True)).pack(side="left")
        self.llm_hint = ttk.Label(row, foreground="#666666", text="checking services…")
        self.llm_hint.pack(side="left", padx=(8, 0))

        ttk.Label(panel, text="Question / instruction:").pack(anchor="w", pady=(6, 0))
        self.prompt = scrolledtext.ScrolledText(panel, height=3, wrap="word")
        self.prompt.pack(fill="x")

        arow = ttk.Frame(panel)
        arow.pack(fill="x", pady=(4, 0))
        self.btn_run = ttk.Button(arow, text="▶ Run analysis", command=self._run_analysis)
        self.btn_run.pack(side="left")
        ttk.Button(arow, text="Past analyses…",
                   command=self._show_past_analyses).pack(side="left", padx=6)
        self.analysis_status = ttk.Label(arow, foreground="#2c3e50", text="")
        self.analysis_status.pack(side="left", padx=(8, 0))

        ttk.Label(panel, text="Answer:").pack(anchor="w", pady=(6, 0))
        self.answer = scrolledtext.ScrolledText(panel, height=6, wrap="word",
                                                state="disabled")
        self.answer.pack(fill="both", expand=True)

        try:
            self.frame.after(300, self._probe_llm)
        except tk.TclError:
            pass

    # ----------------------------------------------------------- workspace -- #
    def _browse_ws(self):
        path = filedialog.askdirectory(title="Select a workspace folder")
        if path:
            self.ent_ws.delete(0, tk.END)
            self.ent_ws.insert(0, path)
            self._open_workspace(path)

    def _open_ws_clicked(self):
        self._open_workspace(self.ent_ws.get().strip())

    def _open_workspace(self, path):
        from data_extraction.db.store import ExtractionStore, WORKSPACE_DB_NAME
        if not path:
            return
        db_path = os.path.join(path, WORKSPACE_DB_NAME)
        if not os.path.exists(db_path):
            messagebox.showinfo(
                "No database here",
                f"No {WORKSPACE_DB_NAME} in:\n{path}\n\nRun an extraction or a "
                "review with this folder as the storage destination first, or "
                "use 'Scan all workspaces'.")
            return
        self.store = ExtractionStore(db_path)
        self.workspace_dir = path
        self.lbl_ws.config(text=f"Loaded: {db_path}")
        self._refresh_docs_from_store()

    def _refresh_docs_from_store(self):
        self.docs.delete(*self.docs.get_children())
        self._row_meta.clear()
        db_path = os.path.join(self.workspace_dir,
                               os.path.basename(self.store.db_path))
        for doc in self.store.list_documents():
            iid = self.docs.insert(
                "", "end", text=doc.get("doc_name") or doc["doc_uuid"],
                values=(doc.get("doc_type") or "", doc.get("status") or "",
                        doc.get("num_sections") or 0))
            self._row_meta[iid] = (self.store.db_path, doc["doc_uuid"])

    def _scan_all(self):
        from data_extraction.db.registry import WorkspaceRegistry
        try:
            rows = WorkspaceRegistry().list_all_documents()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Registry error", str(exc))
            return
        self.docs.delete(*self.docs.get_children())
        self._row_meta.clear()
        self.store = None
        self.lbl_ws.config(text=f"Scanned registry — {len(rows)} document(s) "
                                "across all workspaces.")
        for d in rows:
            label = f"{d.get('doc_name')}  ·  {os.path.basename(d.get('workspace_path') or '')}"
            iid = self.docs.insert("", "end", text=label,
                                   values=(d.get("doc_type") or "",
                                           d.get("status") or "", ""))
            self._row_meta[iid] = (d.get("workspace_db"), d.get("doc_uuid"))

    # -------------------------------------------------------------- detail -- #
    def _on_doc_select(self, _event=None):
        sel = self.docs.selection()
        if not sel:
            return
        db_path, doc_uuid = self._row_meta.get(sel[0], (None, None))
        if not db_path or not doc_uuid or not os.path.exists(db_path):
            return
        from data_extraction.db.store import ExtractionStore
        store = ExtractionStore(db_path)
        self._fill_sections(store, doc_uuid)
        self._fill_reviews(store, doc_uuid)
        self._fill_entities(store, doc_uuid)

    def _fill_sections(self, store, doc_uuid):
        tv = self.tab_sections["tv"]
        tv.delete(*tv.get_children())
        for sec in store.get_sections(doc_uuid):
            text = (sec.get("text") or "").replace("\n", " ")
            tv.insert("", "end", text=sec.get("heading") or "(untitled)",
                      values=(sec.get("page") or "", sec.get("decision") or "",
                              text[:160]))

    def _fill_reviews(self, store, doc_uuid):
        tv = self.tab_reviews["tv"]
        tv.delete(*tv.get_children())
        for r in store.get_ai_reviews(doc_uuid):
            vals = r.get("column_values") or {}
            summary = "; ".join(f"{k}={v}" for k, v in vals.items()
                                if k != "Section")[:160]
            tv.insert("", "end", text=r.get("section_title") or "",
                      values=(r.get("service_used") or "", r.get("model_used") or "",
                              r.get("fallback") or "", summary))

    def _fill_entities(self, store, doc_uuid):
        tv = self.tab_entities["tv"]
        tv.delete(*tv.get_children())
        for e in store.get_entities(doc_uuid):
            tv.insert("", "end", text=e.get("section_title") or "",
                      values=(e.get("reference") or "", e.get("process") or "",
                              e.get("chain") or ""))

    # ----------------------------------------------------------- LLM picker -- #
    def _probe_llm(self, force=False):
        """Populate the LLM picker from the stored model lists; ``force=True``
        (the ↻ button) re-fetches them live first."""
        try:
            self.llm_hint.config(text="refreshing services…" if force
                                 else "checking services…")
        except tk.TclError:
            pass

        def work():
            try:
                from data_extraction.ai_utils import llm_utils as _lu
                avail = _lu.probe_available_services(force_refresh=force)
            except Exception:  # noqa: BLE001
                avail = []
            try:
                self.frame.after(0, lambda a=avail: self._apply_llm(a))
            except (RuntimeError, tk.TclError):
                pass

        threading.Thread(target=work, daemon=True).start()

    def _apply_llm(self, avail):
        self._llm_avail = {label: models for label, models in (avail or [])}
        if not self._llm_avail:
            self.llm_service.configure(values=[])
            self.llm_service.set("")
            self.llm_model.configure(values=[])
            self.llm_model.set("")
            self.llm_hint.config(text="no usable LLM service — add API keys")
            self.btn_run.config(state="disabled")
            return
        labels = list(self._llm_avail)          # BlaBla first
        self.llm_service.configure(values=labels)
        self.llm_service.set(labels[0])
        self.llm_hint.config(text="")
        self.btn_run.config(state="normal")
        self._llm_service_picked()

    def _llm_service_picked(self, _event=None):
        models = self._llm_avail.get(self.llm_service.get().strip()) or []
        self.llm_model.configure(values=models)
        if models:
            self.llm_model.set(models[0])

    def _service_code(self):
        return {"DLR Ollama": "o", "Chat AI": "c"}.get(self.llm_service.get().strip(), "b")

    # -------------------------------------------------------------- analysis -- #
    def _run_analysis(self):
        if self.store is None:
            messagebox.showinfo(
                "Open a workspace",
                "Cross-document analysis runs over one workspace's database. "
                "Open a workspace folder above first (Scan all workspaces is "
                "for browsing only).")
            return
        question = self.prompt.get("1.0", "end").strip()
        if not question:
            messagebox.showinfo("No question", "Enter a question or instruction "
                                "for the analysis.")
            return
        doc_type = {"PDF documents only": "pdf",
                    "EASA documents only": "easa"}.get(self.scope.get())
        code = self._service_code()
        model = self.llm_model.get().strip() or None
        self.btn_run.config(state="disabled")
        self.analysis_status.config(text="Running analysis…")

        def work():
            from data_extraction.db import analysis as _an
            llm = _an.default_llm(code, model)
            result = _an.run_cross_document_analysis(
                self.store, question, llm, name=question[:60], doc_type=doc_type)
            try:
                self.frame.after(0, lambda r=result: self._analysis_done(r))
            except (RuntimeError, tk.TclError):
                pass

        threading.Thread(target=work, daemon=True).start()

    def _analysis_done(self, result):
        self.btn_run.config(state="normal")
        self.answer.config(state="normal")
        self.answer.delete("1.0", "end")
        if not result:
            self.analysis_status.config(text="No answer (no LLM / no stored data).")
            self.answer.insert("1.0", "The analysis produced no result. Check that "
                               "the workspace has stored documents and a usable LLM "
                               "service is selected.")
        else:
            scope = result.get("scope", {})
            self.analysis_status.config(
                text=f"Saved — {scope.get('total_sections', 0)} sections from "
                     f"{len(scope.get('documents', []))} document(s) · "
                     f"{result.get('service_used') or '?'}")
            self.answer.insert("1.0", result.get("result_text") or "")
        self.answer.config(state="disabled")

    def _show_past_analyses(self):
        if self.store is None:
            messagebox.showinfo("Open a workspace", "Open a workspace to see its "
                                "stored analyses.")
            return
        rows = self.store.list_analyses()
        win = tk.Toplevel(self.frame.winfo_toplevel())
        win.title("Stored analyses")
        win.geometry("720x460")
        tv = ttk.Treeview(win, columns=("when", "service"), show="tree headings")
        tv.heading("#0", text="Name")
        tv.heading("when", text="When")
        tv.heading("service", text="Service")
        tv.column("#0", width=360, anchor="w")
        tv.pack(fill="both", expand=True, side="top")
        box = scrolledtext.ScrolledText(win, height=8, wrap="word", state="disabled")
        box.pack(fill="both", expand=True)
        by_iid = {}
        for r in rows:
            iid = tv.insert("", "end", text=r.get("name") or "(analysis)",
                            values=(r.get("created_at") or "",
                                    r.get("service_used") or ""))
            by_iid[iid] = r.get("result_text") or ""

        def show(_e=None):
            sel = tv.selection()
            if not sel:
                return
            box.config(state="normal")
            box.delete("1.0", "end")
            box.insert("1.0", by_iid.get(sel[0], ""))
            box.config(state="disabled")
        tv.bind("<<TreeviewSelect>>", show)
        if not rows:
            box.config(state="normal")
            box.insert("1.0", "No analyses stored in this workspace yet.")
            box.config(state="disabled")
