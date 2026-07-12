"""
Reusable AI-review workbench (mixin) shared by the review UIs.

`AIReviewMixin` provides the complete AI Review page: a shared sections
queue, the free-form review mode (presets, answer-format directive,
pre-run storage dialog with per-result auto-save, export) and the
column-analysis mode (column definitions with unique-element checkboxes,
live prompt preview, row-by-row saving into an accumulating Excel
workbook with per-run snapshot sheets and Uniq sheets via
scripts/excel_file_utils, table export), the evaluation tab (run the
selected checks — substring grounding, Jaccard, ROUGE, BLEU, Levenshtein,
similarity ratio, WER — on a stored analysis workbook against the section
texts that generated it, into the same or a new workbook; with the
auto-evaluate box ticked each section row is evaluated right after it is
saved, via data_extraction.evaluation.column_evaluator),
plus the LLM service/model picker and the API-key manager button.

Host contract — a class mixing this in must provide:
  * ``self.root``          — the Tk root/host widget (for after()/dialogs)
  * ``self.status_var``    — a tk.StringVar used for status messages
  * ``_ai_current_section()``  -> (title, text) or None
  * ``_ai_checked_sections()`` -> list of (title, text) in document order
and call ``_init_ai_state()`` in its ``__init__`` and
``_build_ai_page(parent)`` to build the page into a container widget.
Extracted from data_extraction/easa/json_review_ui.py so the EASA studio
and the chunking studio share one implementation.
"""

import json
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from .llm_utils import get_selected_model, list_available_models, set_selected_model

# Preset review instructions (label -> prompt sent with the section text).
# "Custom…" leaves the box for the user to write their own.
_AI_PRESETS = {
    "Summarize this section": "Summarize this regulatory section concisely.",
    "List key obligations / requirements":
        "List the key obligations or requirements in this section as short bullet points.",
    "Explain in plain language": "Explain this section in plain, non-legal language.",
    "Extract defined terms": "Extract every defined term in this section and its definition.",
    "Find references to other rules":
        "List any references this section makes to other regulations, articles or sections.",
    "Custom…": "",
}

# UI service label -> llm_call service code
_AI_SERVICES = {"Blablador": "b", "DLR Ollama": "o", "Local model": "l"}

# Answer-format label -> directive appended to the LLM prompt ("" = free-form)
_AI_OUTPUT_FORMATS = {
    "Plain text": "",
    "Markdown": "Format your entire answer as Markdown.",
    "JSON": ("Return your entire answer as a single valid JSON object — "
             "no code fences, no commentary outside the JSON."),
    "CSV table": ("Return your entire answer as CSV with a header row — "
                  "no code fences, no commentary outside the CSV."),
}


# Evaluation tab options: UI label -> column_evaluator metric name(s).
_EVAL_METRIC_OPTIONS = [
    ("Substring check (item grounded in reference)", ("grounding",)),
    ("Jaccard similarity", ("jaccard",)),
    ("ROUGE-1/2/L", ("rouge1", "rouge2", "rougeL")),
    ("BLEU", ("bleu",)),
    ("Levenshtein distance", ("levenshtein_distance",)),
    ("Similarity ratio", ("similarity_ratio",)),
    ("Word error rate (WER)", ("word_error_rate",)),
]

_EXCEL_UTILS = None


def _excel_utils():
    """Load scripts/excel_file_utils.py (repo root) once, by file path."""
    global _EXCEL_UTILS
    if _EXCEL_UTILS is None:
        import importlib.util
        p = Path(__file__).resolve().parents[2] / "scripts" / "excel_file_utils.py"
        spec = importlib.util.spec_from_file_location("excel_file_utils", str(p))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        _EXCEL_UTILS = mod
    return _EXCEL_UTILS


class AIReviewMixin:
    """Complete AI Review page; see module docstring for the host contract."""

    def _init_ai_state(self):
        self._ai_batch = []           # queued (title, text) sections for looped AI runs
        self._ai_results = []         # accumulated free-form AI results for export
        self._ai_busy = False
        self._ai_columns = [          # (name, what the LLM should extract)
            ("Summary", "A one-sentence summary of the section."),
            ("References", "Other rules, articles or sections this section refers to."),
        ]
        self._col_results = []        # column-analysis rows (dicts) for export
        self._col_table_cols = []     # column names currently shown in the table
        self._col_uniq_checked = {name for name, _ in self._ai_columns}
        self._col_run_uniq_cols = []  # checked columns snapshotted at run start
        self._ai_autosave_path = None # where the current free-form run auto-saves
        self._ai_run_start = 0        # index of the first result of the current run
        self._ai_save_error_shown = False
        self._col_store_path = None   # workbook the column analyses accumulate into
        self._col_rows_saved = 0      # rows of the current run already on disk
        self._col_run_sheet = None    # snapshot sheet of the current run (.xlsx)
        self._col_save_error_shown = False
        self._col_run_refs = {}       # {section title: text} of the current run
        self._col_run_auto_eval = False  # evaluate each row as it is saved
        self._col_run_eval_metrics = []  # evaluations chosen for this run
        self._col_run_eval_uniq = False  # also evaluate unique values per row
        self._col_eval_entries = []   # per-section evaluations of this run
        self._eval_busy = False       # an evaluation is running

    # -- host hooks --------------------------------------------------------- #
    def _ai_current_section(self):
        """Return (title, text) for the host's current section, or None."""
        raise NotImplementedError

    def _ai_checked_sections(self):
        """Return the host's checked sections as [(title, text), ...]."""
        raise NotImplementedError

    def _default_export_name(self, kind, ext):
        stem = Path(getattr(self, "json_path", None) or "review").stem
        return f"{stem}_{kind}.{ext}"

    def _build_ai_page(self, parent):
        ai = ttk.Frame(parent, padding=6)
        ai.pack(fill="both", expand=True)

        # Integrated Modern Service Engine Frame
        llm_frame = ttk.Frame(ai)
        llm_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(llm_frame, text="LLM Processing Service Engine:").pack(side="left", padx=5)
        self.llm_choice_an = ttk.Combobox(llm_frame, values=["O", "B"], width=5, state="readonly")
        self.llm_choice_an.set("O")
        self.llm_choice_an.pack(side="left", padx=5)
        
        ttk.Button(llm_frame, text="Choose Model...",
                   command=lambda: self._choose_model_action(self.llm_choice_an.get())
                   ).pack(side="left", padx=5)
        
        ttk.Button(llm_frame, text="API keys…", command=self._ai_manage_keys).pack(side="right", padx=5)

        body = ttk.PanedWindow(ai, orient="horizontal")
        body.pack(fill="both", expand=True, pady=(6, 0))

        # Left: the shared sections queue both run modes consume
        batch_wrap = ttk.LabelFrame(body, text=" Sections queue ", padding=4)
        body.add(batch_wrap, weight=1)
        self.ai_batch_list = tk.Listbox(batch_wrap, height=6, selectmode="extended")
        self.ai_batch_list.pack(fill="both", expand=True)
        qbtns = ttk.Frame(batch_wrap)
        qbtns.pack(fill="x", pady=(4, 0))
        ttk.Button(qbtns, text="＋ Current node", command=self._ai_add_to_batch).pack(side="left")
        ttk.Button(qbtns, text="＋ Checked", command=self._ai_add_checked).pack(side="left", padx=4)
        ttk.Button(qbtns, text="Remove", command=self._ai_remove_selected).pack(side="left")
        ttk.Button(qbtns, text="Clear", command=self._ai_clear_batch).pack(side="left", padx=4)

        # Right: free-form review | column analysis
        self.ai_nb = ttk.Notebook(body)
        body.add(self.ai_nb, weight=3)
        self._build_ai_freeform_tab()
        self._build_ai_columns_tab()
        self._build_ai_eval_tab()

    def _build_ai_freeform_tab(self):
        ff = ttk.Frame(self.ai_nb, padding=6)
        self.ai_nb.add(ff, text="Free-form review")

        cfg = ttk.Frame(ff)
        cfg.pack(fill="x")
        ttk.Label(cfg, text="Preset:").pack(side="left")
        self.ai_preset = ttk.Combobox(cfg, state="readonly", width=28,
                                      values=list(_AI_PRESETS.keys()))
        self.ai_preset.pack(side="left", padx=4)
        self.ai_preset.bind("<<ComboboxSelected>>", self._ai_on_preset)
        ttk.Label(cfg, text="Answer as:").pack(side="left", padx=(10, 0))
        self.ai_format = ttk.Combobox(cfg, state="readonly", width=10,
                                      values=list(_AI_OUTPUT_FORMATS.keys()))
        self.ai_format.set("Plain text")
        self.ai_format.pack(side="left", padx=4)

        ttk.Label(ff, text="Instruction (the section's text is appended automatically):").pack(
            anchor="w", pady=(6, 0))
        self.ai_prompt = scrolledtext.ScrolledText(ff, height=3, wrap="word")
        self.ai_prompt.pack(fill="x")
        self.ai_prompt.insert("1.0", _AI_PRESETS["Summarize this section"])
        self.ai_preset.set("Summarize this section")

        btns = ttk.Frame(ff)
        btns.pack(fill="x", pady=6)
        self.ai_run_btn = ttk.Button(btns, text="Run on current node", command=self._ai_run_current)
        self.ai_run_btn.pack(side="left")
        self.ai_batch_btn = ttk.Button(btns, text="Run queued (0)", command=self._ai_run_batch)
        self.ai_batch_btn.pack(side="left", padx=4)
        ttk.Button(btns, text="Export results…", command=self._ai_export).pack(side="right")

        res_wrap = ttk.LabelFrame(ff, text=" Results ", padding=4)
        res_wrap.pack(fill="both", expand=True)
        self.ai_results = scrolledtext.ScrolledText(res_wrap, wrap="word", state="disabled")
        self.ai_results.pack(fill="both", expand=True)

    def _build_ai_columns_tab(self):
        ca = ttk.Frame(self.ai_nb, padding=6)
        self.ai_nb.add(ca, text="Column analysis")

        top = ttk.PanedWindow(ca, orient="horizontal")
        top.pack(fill="x")

        # Column definitions: every change here rebuilds the prompt preview.
        cols_wrap = ttk.LabelFrame(top, text=" Columns to extract (✓ = collect unique elements) ",
                                   padding=4)
        top.add(cols_wrap, weight=1)
        self.col_defs = ttk.Treeview(cols_wrap, columns=("uniq", "what"),
                                     show="tree headings", height=5)
        self.col_defs.heading("#0", text="Column")
        self.col_defs.heading("uniq", text="✓", command=self._col_toggle_uniq_all)
        self.col_defs.heading("what", text="What should the LLM extract?")
        self.col_defs.column("#0", width=140, anchor="w")
        self.col_defs.column("uniq", width=34, minwidth=28, anchor="center", stretch=False)
        self.col_defs.column("what", width=350, anchor="w")
        self.col_defs.pack(fill="x")
        self.col_defs.bind("<<TreeviewSelect>>", self._col_on_select)
        self.col_defs.bind("<Button-1>", self._col_defs_click)

        edit = ttk.Frame(cols_wrap)
        edit.pack(fill="x", pady=(4, 0))
        ttk.Label(edit, text="Name:").pack(side="left")
        self.col_name = ttk.Entry(edit, width=16)
        self.col_name.pack(side="left", padx=4)
        ttk.Label(edit, text="Extract:").pack(side="left")
        self.col_what = ttk.Entry(edit)
        self.col_what.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(edit, text="Add / Update", command=self._col_add_update).pack(side="left")
        ttk.Button(edit, text="Remove", command=self._col_remove).pack(side="left", padx=4)
        ttk.Button(edit, text="Clear all", command=self._col_clear).pack(side="left")

        prev_wrap = ttk.LabelFrame(top, text=" Prompt preview (sent per section) ", padding=4)
        top.add(prev_wrap, weight=1)
        self.col_preview = scrolledtext.ScrolledText(prev_wrap, height=8, wrap="word",
                                                     state="disabled")
        self.col_preview.pack(fill="both", expand=True)

        run_row = ttk.Frame(ca)
        run_row.pack(fill="x", pady=6)
        self.col_run_btn = ttk.Button(run_row, text="Analyze queued (0)", command=self._col_run)
        self.col_run_btn.pack(side="left")
        # When on, Analyze queued first asks which evaluations to run
        # automatically once the batch finishes (see the Evaluation tab).
        self.eval_auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(run_row, variable=self.eval_auto_var,
                        text="Auto-evaluate after the analysis").pack(side="left", padx=10)
        ttk.Button(run_row, text="Export table…", command=self._col_export).pack(side="right")

        table_wrap = ttk.LabelFrame(ca, text=" Analysis table (one row per section) ", padding=4)
        table_wrap.pack(fill="both", expand=True)
        self.col_table = ttk.Treeview(table_wrap, show="headings")
        ct_y = ttk.Scrollbar(table_wrap, orient="vertical", command=self.col_table.yview)
        ct_x = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.col_table.xview)
        self.col_table.configure(yscrollcommand=ct_y.set, xscrollcommand=ct_x.set)
        ct_y.pack(side="right", fill="y")
        ct_x.pack(side="bottom", fill="x")
        self.col_table.pack(side="left", fill="both", expand=True)
        self.col_table.bind("<Double-Button-1>", self._col_show_row)

        self._col_refresh_defs()

    def _ai_manage_keys(self):
        try:
            from ..ai_utils.key_manager_ui import open_api_key_dialog
        except Exception as exc:  # pragma: no cover - ai_utils should always ship
            messagebox.showerror("API Keys", f"Key manager unavailable:\n{exc}")
            return
        open_api_key_dialog(self.root)

    def _ai_on_preset(self, event=None):
        preset = self.ai_preset.get()
        text = _AI_PRESETS.get(preset, "")
        if preset == "Custom…":
            return  # leave whatever the user typed
        self.ai_prompt.delete("1.0", tk.END)
        self.ai_prompt.insert("1.0", text)

    def _ai_append(self, s):
        self.ai_results.configure(state="normal")
        self.ai_results.insert("end", s)
        self.ai_results.see("end")
        self.ai_results.configure(state="disabled")

    def _ai_run_current(self):
        section = self._ai_current_section()
        if section is None:
            messagebox.showinfo("No section", "Select a section first.")
            return
        self._ai_run([section])

    def _ai_add_to_batch(self):
        section = self._ai_current_section()
        if section is None:
            messagebox.showinfo("No section", "Select a section first.")
            return
        self._ai_batch.append(section)
        self.ai_batch_list.insert("end", section[0])
        self._ai_update_queue_counts()

    def _ai_add_checked(self):
        sections = self._ai_checked_sections()
        if not sections:
            messagebox.showinfo(
                "Nothing checked",
                "Tick the ✓ checkboxes first (or use Select all).")
            return
        queued = {(t, txt) for t, txt in self._ai_batch}
        added = 0
        for job in sections:
            if job in queued:
                continue
            self._ai_batch.append(job)
            self.ai_batch_list.insert("end", job[0])
            added += 1
        self._ai_update_queue_counts()
        self.status_var.set(f"Added {added} checked section(s) to the AI batch"
                            + (f" ({len(sections) - added} already queued)." if added < len(sections) else "."))

    def _ai_remove_selected(self):
        for i in reversed(self.ai_batch_list.curselection()):
            del self._ai_batch[i]
            self.ai_batch_list.delete(i)
        self._ai_update_queue_counts()

    def _ai_clear_batch(self):
        self._ai_batch = []
        self.ai_batch_list.delete(0, tk.END)
        self._ai_update_queue_counts()

    def _ai_update_queue_counts(self):
        n = len(self._ai_batch)
        self.ai_batch_btn.configure(text=f"Run queued ({n})")
        self.col_run_btn.configure(text=f"Analyze queued ({n})")

    def _ai_run_batch(self):
        if not self._ai_batch:
            messagebox.showinfo("Empty batch", "Add one or more sections to the batch first.")
            return
        fmt = self._ai_ask_format()
        if fmt is None:  # user cancelled
            return
        self.ai_format.set(fmt)
        self._ai_run(list(self._ai_batch))

    def _ai_ask_format(self):
        """Modal dialog: how should the LLM format the results? None = cancel."""
        win = tk.Toplevel(self.root)
        win.title("Result format")
        win.transient(self.root.winfo_toplevel())
        win.resizable(False, False)
        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="In which format should the LLM return the batch results?\n"
                            "(The instruction is sent to the model with every section.)").pack(anchor="w")
        var = tk.StringVar(value=self.ai_format.get() or "Plain text")
        for fmt in _AI_OUTPUT_FORMATS:
            ttk.Radiobutton(frm, text=fmt, value=fmt, variable=var).pack(anchor="w", pady=1)
        chosen = []
        btns = ttk.Frame(frm)
        btns.pack(anchor="e", pady=(8, 0))
        ttk.Button(btns, text="Run", command=lambda: (chosen.append(var.get()), win.destroy())).pack(
            side="left", padx=4)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left")
        win.grab_set()
        win.wait_window()
        return chosen[0] if chosen else None

    def _ai_run(self, jobs):
        if self._ai_busy:
            messagebox.showinfo("Busy", "An AI run is already in progress.")
            return
        instruction = self.ai_prompt.get("1.0", tk.END).strip()
        if not instruction:
            messagebox.showinfo("No instruction", "Enter or pick a review instruction first.")
            return

        # Always pick the storage file first; its extension is the export format.
        path = filedialog.asksaveasfilename(
            title="Store this run's results as… (file type = format)",
            defaultextension=".md",
            filetypes=[("Markdown", "*.md"), ("Text", "*.txt"), ("CSV", "*.csv"),
                       ("Excel", "*.xlsx"), ("JSON", "*.json")],
            initialfile=self._default_export_name("ai_review", "md"))
        if not path:
            self.status_var.set("AI run cancelled — no storage file chosen.")
            return
        self._ai_autosave_path = path
        self._ai_run_start = len(self._ai_results)
        self._ai_save_error_shown = False

        # Dynamically map the modern engine panel to active session choices
        provider_code = self.llm_choice_an.get().upper()
        service = "o" if provider_code == "O" else "b"
        service_label = "DLR Ollama" if provider_code == "O" else "BlaBla"
        model = get_selected_model(service_label)
        
        fmt = self.ai_format.get() or "Plain text"
        directive = _AI_OUTPUT_FORMATS.get(fmt, "")

        self._ai_busy = True
        self._ai_set_busy_buttons("disabled")
        self._ai_q = queue.Queue()
        self._ai_append(f"\n=== Running on {len(jobs)} section(s) via {service_label}"
                        f" ({model or 'default'}) — answer as {fmt} ===\n")
        threading.Thread(target=self._ai_worker,
                         args=(jobs, instruction, service, service_label, model,
                               fmt, directive),
                         daemon=True).start()
        self.root.after(150, self._ai_poll)

    def _ai_worker(self, jobs, instruction, service, service_label, model,
                   fmt, directive):
        try:
            from data_extraction.ai_utils.llm_utils import llm_call
        except Exception as exc:  # noqa: BLE001
            self._ai_q.put(("error", f"Could not load LLM utilities: {exc}"))
            self._ai_q.put(None)
            return
        full_instruction = f"{instruction}\n{directive}" if directive else instruction
        for title, text in jobs:
            prompt = f"{full_instruction}\n\n---\nSECTION: {title}\n\n{text}"
            try:
                resp = llm_call(prompt, None, service, model)
            except Exception as exc:  # noqa: BLE001
                resp = f"[ERROR] {exc}"
            self._ai_q.put(("result", {
                "title": title, "instruction": instruction, "response": resp,
                "service": service_label, "model": model or "", "format": fmt,
            }))
        self._ai_q.put(None)

    def _ai_poll(self):
        try:
            while True:
                item = self._ai_q.get_nowait()
                if item is None:
                    self._ai_busy = False
                    self._ai_set_busy_buttons("normal")
                    done = f"AI review complete — {len(self._ai_results)} result(s) total"
                    saved = len(self._ai_results) - self._ai_run_start
                    if self._ai_autosave_path and saved:
                        done += f", {saved} saved to {os.path.basename(self._ai_autosave_path)}"
                    self.status_var.set(done + ".")
                    return
                kind, payload = item
                if kind == "error":
                    self._ai_append(f"[ERROR] {payload}\n")
                else:
                    self._ai_results.append(payload)
                    self._ai_append(f"\n## {payload['title']}\n{payload['response']}\n")
                    self._ai_autosave()  # keep the file current after every result
        except queue.Empty:
            pass
        self.root.after(150, self._ai_poll)

    def _ai_export(self):
        if not self._ai_results:
            messagebox.showinfo("No results", "Run an AI review first, then export.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[("Markdown", "*.md"), ("Text", "*.txt"), ("CSV", "*.csv"),
                       ("Excel", "*.xlsx"), ("JSON", "*.json")],
            initialfile=self._default_export_name("ai_review", "md"))
        if not path:
            return
        try:
            self._write_ai_results(path, self._ai_results)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Export failed", str(exc))
            return
        self.status_var.set(f"Exported {len(self._ai_results)} AI result(s) → {path}")
        messagebox.showinfo("Export complete", f"Wrote {len(self._ai_results)} result(s) to:\n{path}")

    @staticmethod
    def _write_ai_results(path, results):
        """Write free-form results; the extension picks the format. Raises on failure."""
        ext = os.path.splitext(path)[1].lower()
        columns = ["title", "service", "model", "format", "instruction", "response"]
        if ext == ".json":
            with open(path, "w", encoding="utf-8") as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
        elif ext == ".csv":
            import csv
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=columns)
                w.writeheader()
                w.writerows(results)
        elif ext == ".xlsx":
            from openpyxl import Workbook
            from openpyxl.styles import Alignment, Font
            wb = Workbook()
            ws = wb.active
            ws.title = "AI Review"
            ws.append([c.capitalize() for c in columns])
            for cell in ws[1]:
                cell.font = Font(bold=True)
            for r in results:
                ws.append([str(r.get(c, "")) for c in columns])
            widths = {"A": 40, "B": 12, "C": 18, "D": 12, "E": 40, "F": 90}
            for col, width in widths.items():
                ws.column_dimensions[col].width = width
            wrap = Alignment(wrap_text=True, vertical="top")
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = wrap
            wb.save(path)
        else:  # markdown / text
            blocks = []
            for r in results:
                blocks.append(f"## {r['title']}\n"
                              f"*{r['service']} {r['model']}* — {r['instruction']}\n\n"
                              f"{r['response']}\n")
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(blocks))

    def _ai_autosave(self):
        """Rewrite the storage file with the current run's results so far —
        called after every LLM reply, so the file grows result by result."""
        path = self._ai_autosave_path
        results = self._ai_results[self._ai_run_start:]
        if not path or not results:
            return
        try:
            self._write_ai_results(path, results)
        except Exception as exc:  # noqa: BLE001
            if not self._ai_save_error_shown:
                self._ai_save_error_shown = True
                messagebox.showerror("Auto-save failed", f"{path}\n\n{exc}")
            self.status_var.set(f"Auto-save failed: {exc}")
        
    def _choose_model_action(self, provider_code):
        """
        Fetch the live list of available models for the selected provider
        ('O' = DLR Ollama, 'B' = Blablador), let the user pick one, and store
        it as the session model used by all subsequent LLM calls.
        """
        service = "DLR Ollama" if str(provider_code).upper() == "O" else "BlaBla"

        print(f"Fetching available {service} models...")
        try:
            models = list_available_models(service)
        except Exception as e:
            messagebox.showerror("Model list failed", f"Could not fetch models for {service}:\n{e}")
            return

        if not models:
            messagebox.showwarning("No models", f"No models returned for {service}.\nKeeping current: {get_selected_model(service)}")
            return

        current = get_selected_model(service)

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Select {service} Model")
        dialog.geometry("560x400")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text=f"Available {service} models (current: {current}):",
                  font=("Arial", 10, "bold")).pack(padx=10, pady=8, anchor="w")

        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.pack(side="left", fill="both", expand=True)

        for m in models:
            listbox.insert(tk.END, m)
        if current in models:
            idx = models.index(current)
            listbox.selection_set(idx)
            listbox.see(idx)

        def _on_confirm():
            sel = listbox.curselection()
            if not sel:
                messagebox.showinfo("No selection", "Please select a model, or close the dialog to keep the current one.")
                return
            set_selected_model(service, models[sel[0]])
            dialog.destroy()

        ttk.Button(dialog, text="Use Selected Model", command=_on_confirm).pack(pady=10)

    def _ai_set_busy_buttons(self, state):
        self.ai_run_btn.configure(state=state)
        self.ai_batch_btn.configure(state=state)
        self.col_run_btn.configure(state=state)
        self.eval_run_btn.configure(state=state)

    # ------------------------------------------------- AI column analysis -- #
    def _col_refresh_defs(self):
        """Redraw the column-definition list and the live prompt preview."""
        self.col_defs.delete(*self.col_defs.get_children())
        for name, what in self._ai_columns:
            glyph = "☑" if name in self._col_uniq_checked else "☐"
            self.col_defs.insert("", "end", text=name, values=(glyph, what))
        self.col_preview.configure(state="normal")
        self.col_preview.delete("1.0", tk.END)
        if self._ai_columns:
            self.col_preview.insert("1.0", self._col_prompt_text()
                                    + "\n\n---\nSECTION: <title>\n\n<section text>")
        else:
            self.col_preview.insert("1.0", "Add at least one column to build the prompt.")
        self.col_preview.configure(state="disabled")

    def _col_prompt_text(self):
        lines = ["Analyze the regulatory section below and extract the following data.",
                 "Return ONLY one valid JSON object with exactly these keys "
                 "(no code fences, no commentary):"]
        for name, what in self._ai_columns:
            lines.append(f'- "{name}": {what or "extract this value"}')
        lines.append('Use "" for any value the section does not contain.')
        lines.append('Use ";" for multiple values the section does contain.')
        return "\n".join(lines)

    def _col_on_select(self, event=None):
        sel = self.col_defs.selection()
        if not sel:
            return
        name = self.col_defs.item(sel[0], "text")
        values = self.col_defs.item(sel[0], "values") or ("", "")
        what = values[1] if len(values) > 1 else ""
        self.col_name.delete(0, tk.END)
        self.col_name.insert(0, name)
        self.col_what.delete(0, tk.END)
        self.col_what.insert(0, what)

    def _col_defs_click(self, event):
        """Toggle unique-element collection when the ✓ cell is clicked."""
        if self.col_defs.identify("region", event.x, event.y) != "cell":
            return None
        if self.col_defs.identify_column(event.x) != "#1":
            return None
        iid = self.col_defs.identify_row(event.y)
        if not iid:
            return None
        name = self.col_defs.item(iid, "text")
        if name in self._col_uniq_checked:
            self._col_uniq_checked.discard(name)
        else:
            self._col_uniq_checked.add(name)
        self.col_defs.set(iid, "uniq", "☑" if name in self._col_uniq_checked else "☐")
        return "break"

    def _col_toggle_uniq_all(self):
        """✓ heading click: check every column, or clear if all are checked."""
        names = {name for name, _ in self._ai_columns}
        if names and names <= self._col_uniq_checked:
            self._col_uniq_checked -= names
        else:
            self._col_uniq_checked |= names
        self._col_refresh_defs()

    def _col_add_update(self):
        name = self.col_name.get().strip()
        what = self.col_what.get().strip()
        if not name:
            messagebox.showinfo("Column name missing", "Give the column a name first.")
            return
        for i, (existing, _) in enumerate(self._ai_columns):
            if existing == name:
                self._ai_columns[i] = (name, what)
                break
        else:
            self._ai_columns.append((name, what))
            self._col_uniq_checked.add(name)   # new columns collect uniques by default
        self._col_refresh_defs()

    def _col_remove(self):
        names = {self.col_defs.item(iid, "text") for iid in self.col_defs.selection()}
        if not names:
            return
        self._ai_columns = [(n, w) for n, w in self._ai_columns if n not in names]
        self._col_uniq_checked -= names
        self._col_refresh_defs()

    def _col_clear(self):
        self._ai_columns = []
        self._col_uniq_checked.clear()
        self._col_refresh_defs()

    def _col_run(self):
        if self._ai_busy:
            messagebox.showinfo("Busy", "An AI run is already in progress.")
            return
        if not self._ai_columns:
            messagebox.showinfo("No columns", "Define at least one column to extract.")
            return
        if not self._ai_batch:
            messagebox.showinfo("Empty queue",
                                "Add sections to the queue first (current node or checked nodes).")
            return
        jobs = list(self._ai_batch)
        columns = [name for name, _ in self._ai_columns]
        prompt_head = self._col_prompt_text()

        # Always pick the storage file first. Re-choosing the same .xlsx appends
        # this run as a new snapshot sheet instead of overwriting.
        opts = {}
        if self._col_store_path:
            opts["initialdir"] = os.path.dirname(self._col_store_path)
            opts["initialfile"] = os.path.basename(self._col_store_path)
        else:
            opts["initialfile"] = self._default_export_name("ai_columns", "xlsx")
        store = filedialog.asksaveasfilename(
            title="Store analysis results in… (existing .xlsx gets this run as a new sheet)",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx"), ("CSV", "*.csv"), ("JSON", "*.json")],
            confirmoverwrite=False, **opts)
        if not store:
            self.status_var.set("Column analysis cancelled — no storage file chosen.")
            return
        self._col_store_path = store
        # With the auto-evaluate box ticked, ask which evaluations should run
        # on each section result as soon as it is saved ("Skip evaluation"
        # opts out for this batch).
        auto_eval = False
        if self.eval_auto_var.get() and store.lower().endswith(".xlsx"):
            auto_eval = self._col_ask_auto_eval()
        self._col_run_auto_eval = auto_eval
        # Snapshot the choices so mid-run changes don't affect this batch.
        self._col_run_eval_metrics = self._eval_selected_metrics() if auto_eval else []
        self._col_run_eval_uniq = auto_eval and self.eval_uniq_var.get()
        self._col_eval_entries = []     # per-section evaluations of this run
        self._col_rows_saved = 0        # file/sheet is created with the 1st row
        self._col_run_sheet = None
        self._col_save_error_shown = False
        # Only ✓-checked columns get unique-element collection this run.
        self._col_run_uniq_cols = [c for c in columns if c in self._col_uniq_checked]
        # Reference texts of this run's sections, for the evaluation afterwards.
        self._col_run_refs = dict(jobs)

        # Dynamically map modern engine panel values to active session configurations
        provider_code = self.llm_choice_an.get().upper()
        service = "o" if provider_code == "O" else "b"
        service_label = "DLR Ollama" if provider_code == "O" else "BlaBla"
        model = get_selected_model(service_label)

        # New run: rebuild the table for the current column set.
        self._col_results = []
        self._col_table_cols = ["Section"] + columns
        self.col_table.delete(*self.col_table.get_children())
        self.col_table.configure(columns=self._col_table_cols)
        for c in self._col_table_cols:
            self.col_table.heading(c, text=c)
            self.col_table.column(c, width=140 if c == "Section" else 180, anchor="w")

        self._ai_busy = True
        self._ai_set_busy_buttons("disabled")
        self._col_q = queue.Queue()
        self.status_var.set(f"Analyzing {len(jobs)} section(s) into {len(columns)} column(s)…")
        threading.Thread(target=self._col_worker,
                         args=(jobs, columns, prompt_head, service, model),
                         daemon=True).start()
        self.root.after(150, self._col_poll)

    def _col_worker(self, jobs, columns, prompt_head, service, model):
        try:
            from data_extraction.ai_utils.llm_utils import llm_call
        except Exception as exc:  # noqa: BLE001
            self._col_q.put(("error", f"Could not load LLM utilities: {exc}"))
            self._col_q.put(None)
            return
        for title, text in jobs:
            prompt = f"{prompt_head}\n\n---\nSECTION: {title}\n\n{text}"
            try:
                resp = llm_call(prompt, None, service, model)
            except Exception as exc:  # noqa: BLE001
                resp = None
                parsed = {columns[0]: f"[ERROR] {exc}"}
            else:
                parsed = self._parse_llm_json(resp)
                if parsed is None:
                    parsed = {columns[0]: f"[unparsed] {resp}"}
            row = {"Section": title}
            for c in columns:
                v = parsed.get(c, "")
                if isinstance(v, (list, tuple)):
                    v = "; ".join(str(x) for x in v)
                elif isinstance(v, dict):
                    v = json.dumps(v, ensure_ascii=False)
                row[c] = "" if v is None else str(v)
            self._col_q.put(("row", row))
        self._col_q.put(None)

    @staticmethod
    def _parse_llm_json(resp):
        """Best-effort: pull the first {...} object out of an LLM reply."""
        if not resp or not isinstance(resp, str):
            return None
        start, end = resp.find("{"), resp.rfind("}")
        if start == -1 or end <= start:
            return None
        try:
            parsed = json.loads(resp[start:end + 1])
        except (ValueError, TypeError):
            return None
        return parsed if isinstance(parsed, dict) else None

    def _col_poll(self):
        try:
            while True:
                item = self._col_q.get_nowait()
                if item is None:
                    self._ai_busy = False
                    self._ai_set_busy_buttons("normal")
                    done = f"Column analysis complete — {len(self._col_results)} row(s)"
                    if self._col_store_path and self._col_rows_saved:
                        done += f", saved to {os.path.basename(self._col_store_path)}"
                        if self._col_run_sheet:
                            done += f" (sheet '{self._col_run_sheet}')"
                    self.status_var.set(done + ".")
                    # Rows were evaluated as they were saved; finish up.
                    self._col_eval_finish()
                    return
                kind, payload = item
                if kind == "error":
                    messagebox.showerror("Column analysis", payload)
                else:
                    self._col_results.append(payload)
                    self.col_table.insert("", "end", values=[
                        payload.get(c, "") for c in self._col_table_cols])
                    self._col_save_row(payload)
        except queue.Empty:
            pass
        self.root.after(150, self._col_poll)

    def _col_show_row(self, event=None):
        sel = self.col_table.selection()
        if not sel or not self._col_table_cols:
            return
        values = self.col_table.item(sel[0], "values")
        win = tk.Toplevel(self.root)
        win.title("Row details")
        win.transient(self.root.winfo_toplevel())
        txt = scrolledtext.ScrolledText(win, wrap="word", width=90, height=24)
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        for col, val in zip(self._col_table_cols, values):
            txt.insert(tk.END, f"■ {col}\n{val}\n\n")
        txt.configure(state="disabled")

    @staticmethod
    def _col_write_store(path, cols, rows):
        """Write one analysis snapshot. An existing .xlsx gets a new sheet per
        run (accumulating workbook); CSV/JSON hold the latest batch. Returns
        the sheet name used for .xlsx, else None. Raises on failure."""
        ext = os.path.splitext(path)[1].lower()
        if ext == ".json":
            with open(path, "w", encoding="utf-8") as f:
                json.dump(rows, f, indent=2, ensure_ascii=False)
            return None
        if ext == ".csv":
            import csv
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=cols)
                w.writeheader()
                w.writerows(rows)
            return None
        # Excel: one snapshot sheet per analysis run
        from datetime import datetime

        from openpyxl import Workbook, load_workbook
        from openpyxl.styles import Alignment, Font
        from openpyxl.utils import get_column_letter
        if os.path.exists(path):
            wb = load_workbook(path)
            ws = wb.create_sheet()
        else:
            wb = Workbook()
            ws = wb.active
        stamp = datetime.now().strftime("%Y-%m-%d %H.%M.%S")
        run_no = sum(1 for s in wb.sheetnames if s.startswith("Run ")) + 1
        ws.title = f"Run {run_no} {stamp}"[:31]
        ws.append(cols)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for r in rows:
            ws.append([str(r.get(c, "")) for c in cols])
        ws.column_dimensions["A"].width = 40
        for i in range(2, len(cols) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 50
        wrap = Alignment(wrap_text=True, vertical="top")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = wrap
        wb.save(path)
        return ws.title

    def _col_save_row(self, row):
        """Write one finished section to the storage file immediately: the file
        (and, for .xlsx, the run's snapshot sheet) is created with the first
        row, then updated and saved again for every further row."""
        path = self._col_store_path
        if not path:
            return
        cols = self._col_table_cols
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".json":
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(self._col_results, f, indent=2, ensure_ascii=False)
            elif ext == ".csv":
                import csv
                mode = "w" if self._col_rows_saved == 0 else "a"
                with open(path, mode, newline="", encoding="utf-8") as f:
                    w = csv.DictWriter(f, fieldnames=cols)
                    if self._col_rows_saved == 0:
                        w.writeheader()
                    w.writerow(row)
            else:  # .xlsx
                from datetime import datetime

                from openpyxl import Workbook, load_workbook
                from openpyxl.styles import Alignment, Font
                from openpyxl.utils import get_column_letter
                if self._col_rows_saved == 0:
                    if os.path.exists(path):
                        wb = load_workbook(path)
                        ws = wb.create_sheet()
                    else:
                        wb = Workbook()
                        ws = wb.active
                    stamp = datetime.now().strftime("%Y-%m-%d %H.%M.%S")
                    run_no = sum(1 for s in wb.sheetnames if s.startswith("Run ")) + 1
                    ws.title = f"Run {run_no} {stamp}"[:31]
                    self._col_run_sheet = ws.title
                    ws.append(cols)
                    for cell in ws[1]:
                        cell.font = Font(bold=True)
                    ws.column_dimensions["A"].width = 40
                    for i in range(2, len(cols) + 1):
                        ws.column_dimensions[get_column_letter(i)].width = 50
                else:
                    wb = load_workbook(path)
                    ws = wb[self._col_run_sheet]
                ws.append([str(row.get(c, "")) for c in cols])
                wrap = Alignment(wrap_text=True, vertical="top")
                for cell in ws[ws.max_row]:
                    cell.alignment = wrap
                wb.save(path)
                self._col_update_uniques(path)
                self._col_eval_saved_row(row)
            self._col_rows_saved += 1
            self.status_var.set(f"Saved row {self._col_rows_saved} "
                                f"({row.get('Section', '')}) → {os.path.basename(path)}")
        except Exception as exc:  # noqa: BLE001
            if not self._col_save_error_shown:
                self._col_save_error_shown = True
                messagebox.showerror(
                    "Row save failed",
                    f"{path}\n\n{exc}\n\nThe analysis continues; further save "
                    f"errors only go to the status bar.")
            self.status_var.set(f"Row save failed: {exc}")

    def _col_update_uniques(self, path):
        """Refresh the run's unique-elements sheet (element + count) via
        scripts/excel_file_utils after each saved row — but only for the
        columns whose ✓ checkbox was ticked when the run started."""
        data_cols = [c for c in self._col_run_uniq_cols if c != "Section"]
        if not data_cols or not self._col_run_sheet:
            return
        uniq_sheet = f"Uniq {self._col_run_sheet}"[:31]
        _excel_utils().save_unique_elements_to_new_sheet(
            path, data_cols, new_sheet_name=uniq_sheet,
            source_sheet=self._col_run_sheet)

    # ----------------------------------------------------- Evaluation tab -- #
    def _build_ai_eval_tab(self):
        """Evaluation page: pick a stored analysis workbook + run sheet, the
        reference sections, the evaluations to compute and where to write
        them (data_extraction.evaluation.column_evaluator)."""
        ev = ttk.Frame(self.ai_nb, padding=6)
        self.ai_nb.add(ev, text="Evaluation")

        wb_wrap = ttk.LabelFrame(ev, text=" Analysis workbook (.xlsx with 'Run N …' sheets) ",
                                 padding=4)
        wb_wrap.pack(fill="x")
        row = ttk.Frame(wb_wrap)
        row.pack(fill="x")
        ttk.Label(row, text="Workbook:").pack(side="left")
        self.eval_wb = ttk.Entry(row)
        self.eval_wb.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row, text="Browse…", command=self._eval_browse_wb).pack(side="left")
        row2 = ttk.Frame(wb_wrap)
        row2.pack(fill="x", pady=(4, 0))
        ttk.Label(row2, text="Run sheet:").pack(side="left")
        self.eval_sheet = ttk.Combobox(row2, state="readonly", width=34, values=[])
        self.eval_sheet.pack(side="left", padx=4)
        ttk.Button(row2, text="Refresh sheets",
                   command=self._eval_refresh_sheets).pack(side="left")

        ref_wrap = ttk.LabelFrame(ev, text=" Reference data (the section content the "
                                           "columns were generated from) ", padding=4)
        ref_wrap.pack(fill="x", pady=(6, 0))
        self.eval_ref_choice = tk.StringVar(value="run")
        ttk.Radiobutton(ref_wrap, text="Sections of the last column-analysis run",
                        variable=self.eval_ref_choice, value="run").pack(anchor="w")
        ttk.Radiobutton(ref_wrap, text="Currently queued sections",
                        variable=self.eval_ref_choice, value="queue").pack(anchor="w")
        jrow = ttk.Frame(ref_wrap)
        jrow.pack(fill="x")
        ttk.Radiobutton(jrow, text="Sections JSON file:",
                        variable=self.eval_ref_choice, value="json").pack(side="left")
        self.eval_ref_json = ttk.Entry(jrow)
        self.eval_ref_json.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(jrow, text="Browse…",
                   command=self._eval_browse_ref_json).pack(side="left")

        met_wrap = ttk.LabelFrame(ev, text=" Evaluations to run (each cell / item vs. "
                                           "its reference text) ", padding=4)
        met_wrap.pack(fill="x", pady=(6, 0))
        self.eval_metric_vars = {}
        grid = ttk.Frame(met_wrap)
        grid.pack(fill="x")
        for i, (label, _names) in enumerate(_EVAL_METRIC_OPTIONS):
            var = tk.BooleanVar(value=True)
            self.eval_metric_vars[label] = var
            ttk.Checkbutton(grid, text=label, variable=var).grid(
                row=i % 4, column=i // 4, sticky="w", padx=(0, 18))
        self.eval_uniq_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(met_wrap, variable=self.eval_uniq_var,
                        text="Also evaluate the unique generated values "
                             "(via save_unique_elements_to_new_sheet)").pack(
            anchor="w", pady=(4, 0))

        out_wrap = ttk.LabelFrame(ev, text=" Write results to ", padding=4)
        out_wrap.pack(fill="x", pady=(6, 0))
        self.eval_out_choice = tk.StringVar(value="same")
        ttk.Radiobutton(out_wrap, text="New sheets in the same workbook "
                                       "(Eval / Uniq Eval next to the run)",
                        variable=self.eval_out_choice, value="same").pack(anchor="w")
        ttk.Radiobutton(out_wrap, text="A new evaluation workbook (asks where; also "
                                       "gets a copy of the run sheet)",
                        variable=self.eval_out_choice, value="new").pack(anchor="w")

        brow = ttk.Frame(ev)
        brow.pack(fill="x", pady=6)
        self.eval_run_btn = ttk.Button(brow, text="Run evaluation",
                                       command=self._ai_eval_run)
        self.eval_run_btn.pack(side="left")
        ttk.Label(brow, foreground="#666666",
                  text="Auto-evaluation is toggled next to “Analyze queued” "
                       "on the Column analysis tab.").pack(side="left", padx=12)

    def _eval_browse_wb(self):
        path = filedialog.askopenfilename(
            title="Select a column-analysis workbook",
            filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            self.eval_wb.delete(0, tk.END)
            self.eval_wb.insert(0, path)
            self._eval_refresh_sheets()

    def _eval_refresh_sheets(self):
        """List the workbook's 'Run N …' snapshot sheets; preselect the last."""
        path = self.eval_wb.get().strip()
        runs = []
        if path and os.path.exists(path):
            try:
                from openpyxl import load_workbook
                wb = load_workbook(path, read_only=True)
                runs = [s for s in wb.sheetnames if s.startswith("Run ")]
                wb.close()
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror("Workbook", f"Could not read sheets:\n{exc}")
        values = (["All runs"] + runs) if runs else []
        self.eval_sheet.configure(values=values)
        self.eval_sheet.set(runs[-1] if runs else "")
        if path and not runs:
            self.status_var.set("No 'Run …' sheets found in the selected workbook.")

    def _eval_browse_ref_json(self):
        path = filedialog.askopenfilename(
            title="Select the sections JSON the analysis was made from",
            filetypes=[("JSON files", "*.json")])
        if path:
            self.eval_ref_json.delete(0, tk.END)
            self.eval_ref_json.insert(0, path)
            self.eval_ref_choice.set("json")

    def _eval_selected_metrics(self):
        metrics = []
        for label, names in _EVAL_METRIC_OPTIONS:
            if self.eval_metric_vars[label].get():
                metrics.extend(names)
        return metrics

    def _eval_references(self, auto=False):
        """Resolve the reference map for the chosen source, or None on error."""
        choice = self.eval_ref_choice.get()
        if choice == "queue":
            refs = dict(self._ai_batch)
            src = "queued sections"
        elif choice == "json":
            path = self.eval_ref_json.get().strip()
            if not path or not os.path.exists(path):
                if not auto:
                    messagebox.showinfo("Reference JSON",
                                        "Pick the sections JSON file first.")
                return None
            try:
                from data_extraction.evaluation.column_evaluator import references_from_json
                refs = references_from_json(path)
            except Exception as exc:  # noqa: BLE001
                if not auto:
                    messagebox.showerror("Reference JSON",
                                         f"Could not load sections:\n{exc}")
                return None
            src = os.path.basename(path)
        else:
            refs = dict(self._col_run_refs)
            src = "last analysis run"
        if not refs:
            if not auto:
                messagebox.showinfo(
                    "No reference sections",
                    f"No sections available from the {src}. Run a column "
                    "analysis, queue sections, or pick a sections JSON file.")
            return None
        return refs

    def _ai_eval_run(self, auto=False):
        """Run the selected evaluations on the chosen workbook/run sheet(s)."""
        if self._eval_busy or (self._ai_busy and not auto):
            if not auto:
                messagebox.showinfo("Busy", "Another run is already in progress.")
            return
        path = self.eval_wb.get().strip()
        if not path or os.path.splitext(path)[1].lower() != ".xlsx" \
                or not os.path.exists(path):
            if not auto:
                messagebox.showinfo("No workbook",
                                    "Pick a stored column-analysis .xlsx first "
                                    "(run an analysis, or Browse… to an "
                                    "existing workbook).")
            return
        sheet = self.eval_sheet.get().strip()
        run_sheets = None if (not sheet or sheet == "All runs") else [sheet]
        metrics = self._eval_selected_metrics()
        if not metrics:
            if not auto:
                messagebox.showinfo("No evaluations",
                                    "Tick at least one evaluation to run.")
            return
        references = self._eval_references(auto)
        if references is None:
            return
        out_path = None
        if self.eval_out_choice.get() == "new" and not auto:
            out_path = filedialog.asksaveasfilename(
                title="Write the evaluation workbook to…",
                defaultextension=".xlsx",
                filetypes=[("Excel workbook", "*.xlsx")],
                initialfile=Path(path).stem + "_evaluation.xlsx",
                confirmoverwrite=False)
            if not out_path:
                self.status_var.set("Evaluation cancelled — no output file chosen.")
                return
        uniq = "all" if self.eval_uniq_var.get() else None

        self._eval_busy = True
        self.eval_run_btn.configure(state="disabled")
        self.status_var.set(f"Evaluating {sheet or 'all runs'} against "
                            f"{len(references)} reference section(s)…")

        def work():
            try:
                from data_extraction.evaluation.column_evaluator import evaluate_workbook
                results = evaluate_workbook(path, references, metrics=metrics,
                                            run_sheets=run_sheets,
                                            uniq_columns=uniq, out_path=out_path)
            except Exception as exc:  # noqa: BLE001
                self.root.after(0, lambda e=exc: self._ai_eval_done(None, e, auto))
            else:
                self.root.after(0, lambda r=results: self._ai_eval_done(r, None, auto))

        threading.Thread(target=work, daemon=True).start()

    def _ai_eval_done(self, results, exc, auto):
        self._eval_busy = False
        self.eval_run_btn.configure(state="normal")
        if exc is not None:
            self.status_var.set(f"Evaluation failed: {exc}")
            if not auto:
                messagebox.showerror("Evaluation failed", str(exc))
            return
        if not results:
            self.status_var.set("Evaluation found no 'Run …' sheets to evaluate.")
            if not auto:
                messagebox.showinfo("Evaluation", "The workbook has no "
                                                  "'Run …' snapshot sheets.")
            return
        lines = []
        for sheet, s in results.items():
            line = f"{sheet}: {s['rows']} row(s)"
            if s.get("coverage") is not None:
                line += (f", grounded {s['true']}/{s['true'] + s['false']}"
                         f" ({s['coverage']}%)")
            line += f" → '{s['eval_sheet']}'"
            if s.get("uniq_eval_sheet"):
                line += f", uniques → '{s['uniq_eval_sheet']}'"
            lines.append(line)
        first = next(iter(results.values()))
        where = os.path.basename(first["out_path"])
        self.status_var.set(f"Evaluation complete → {where}: " + " | ".join(lines))
        if not auto:
            messagebox.showinfo("Evaluation complete",
                                f"Written to {first['out_path']}\n\n" + "\n".join(lines))

    def _col_ask_auto_eval(self):
        """Modal picker shown when 'Analyze queued' starts with the
        auto-evaluate box ticked: choose which of the possible evaluations
        run automatically on this batch once the analysis finishes. The
        checkboxes are shared with the Evaluation tab. Returns True to
        evaluate afterwards, False for 'Skip evaluation'."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Auto-evaluation of this analysis")
        dlg.transient(self.root.winfo_toplevel())
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, justify="left",
                  text="Choose the evaluations to run automatically once the "
                       "batch finishes\n(each cell / item is compared with the "
                       "section text it was generated from):").pack(anchor="w")
        for label, _names in _EVAL_METRIC_OPTIONS:
            ttk.Checkbutton(frm, text=label,
                            variable=self.eval_metric_vars[label]).pack(
                anchor="w", padx=10)
        ttk.Checkbutton(frm, variable=self.eval_uniq_var,
                        text="Also evaluate the unique generated values "
                             "(element + count, grounded in the references)").pack(
            anchor="w", pady=(8, 0))
        choice = {"ok": False}

        def _ok():
            if not self._eval_selected_metrics():
                messagebox.showinfo(
                    "No evaluations",
                    "Tick at least one evaluation — or press Skip evaluation.",
                    parent=dlg)
                return
            choice["ok"] = True
            dlg.destroy()

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(12, 0))
        ttk.Button(btns, text="OK — evaluate after the analysis",
                   command=_ok).pack(side="left")
        ttk.Button(btns, text="Skip evaluation",
                   command=dlg.destroy).pack(side="right")
        dlg.grab_set()
        self.root.wait_window(dlg)
        return choice["ok"]

    def _col_eval_saved_row(self, row):
        """Evaluate one just-saved section result (the chosen evaluations vs.
        its reference text) and refresh the 'Eval <run sheet>' — and, when
        chosen, the 'Uniq Eval <run sheet>' — before the next section is
        analyzed. Failures only reach the status bar; the analysis goes on."""
        if not (self._col_run_auto_eval and self._col_run_eval_metrics
                and self._col_run_sheet):
            return
        try:
            from data_extraction.evaluation.column_evaluator import (
                evaluate_row, evaluate_uniques, write_eval_sheet)
            data_cols = [c for c in self._col_table_cols if c != "Section"]
            entry = evaluate_row(row, self._col_run_refs, data_cols,
                                 self._col_run_eval_metrics)
            self._col_eval_entries.append(entry)
            write_eval_sheet(self._col_store_path, self._col_run_sheet,
                             self._col_eval_entries)
            if self._col_run_eval_uniq:
                uniq_cols = [c for c in self._col_run_uniq_cols if c != "Section"]
                if uniq_cols:
                    evaluate_uniques(self._col_store_path, self._col_run_sheet,
                                     self._col_run_refs, uniq_cols)
        except Exception as exc:  # noqa: BLE001 - never interrupt the analysis
            self.status_var.set(f"Row evaluation failed: {exc}")

    def _col_eval_finish(self):
        """After the batch: prefill the Evaluation tab with this run and, when
        rows were evaluated along the way, show the overall result."""
        path, sheet = self._col_store_path, self._col_run_sheet
        if not (path and sheet and os.path.splitext(path)[1].lower() == ".xlsx"
                and os.path.exists(path)):
            return
        self.eval_wb.delete(0, tk.END)
        self.eval_wb.insert(0, path)
        self._eval_refresh_sheets()
        self.eval_sheet.set(sheet)
        self.eval_ref_choice.set("run")
        if not (self._col_run_auto_eval and self._col_eval_entries):
            return
        parts = [f"{self.status_var.get()} Evaluated row by row → "
                 f"'Eval {sheet}'"[:200]]
        if "grounding" in self._col_run_eval_metrics:
            t = sum(e.get("Row True") or 0 for e in self._col_eval_entries)
            f = sum(e.get("Row False") or 0 for e in self._col_eval_entries)
            if t + f:
                parts.append(f"grounded {t}/{t + f} "
                             f"({round(100.0 * t / (t + f), 1)}%)")
        self.status_var.set(", ".join(parts) + ".")

    def _col_export(self):
        if not self._col_results:
            messagebox.showinfo("No table", "Run a column analysis first, then export.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("JSON", "*.json")],
            initialfile=self._default_export_name("ai_columns", "xlsx"),
            confirmoverwrite=False)
        if not path:
            return
        try:
            sheet = self._col_write_store(path, self._col_table_cols, self._col_results)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Export failed", str(exc))
            return
        where = path + (f" (sheet '{sheet}')" if sheet else "")
        self.status_var.set(f"Exported {len(self._col_results)} analysis row(s) → {where}")
        messagebox.showinfo("Export complete",
                            f"Wrote {len(self._col_results)} row(s) to:\n{where}")

