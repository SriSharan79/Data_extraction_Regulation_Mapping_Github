"""
data_extraction.chunking.chunk_triage_ui
========================================

Bulk chunk review built on :mod:`data_extraction.chunking.chunk_triage`.

The original review walked the document one chunk at a time and demanded a
click for every single one. Here the triage engine proposes a decision for
each chunk first (using the document's Table of Contents, or an LLM pass over
the extracted headings) and this screen shows them all at once, so the
reviewer confirms in bulk and only really inspects what is uncertain.

Nothing is written until the reviewer accepts: the proposals are shown with
their reason, can be re-decided individually or in bulk, and *Accept & save*
produces the same ``merged_headings`` / ``raw_session_history`` payload the
sequential tool always wrote — so Section Review and AI Review are unchanged.
"""

from __future__ import annotations

import json
import os
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from .chunk_triage import (analyze_chunks, build_output_payload,
                           normalize_heading)

_ACTION_LABEL = {"log": "Keep", "skip": "Skip", "review": "Review"}
_ACTION_TAGS = {"log": ("#1f7a1f", ""), "skip": ("#777777", ""),
                "review": ("#b00020", "")}
_FILTERS = ("Needs review", "All", "Keep", "Skip")


class ChunkTriageApp:
    """Bulk review of Docling chunks with automatic triage proposals."""

    def __init__(self, root, chunks_data, output_file_name, logger,
                 on_complete_callback=None, llm=None):
        self.root = root
        self.chunks_data = list(chunks_data or [])
        self.output_file_name = output_file_name
        self.logger = logger
        self.on_complete_callback = on_complete_callback
        self._llm = llm

        self.proposals = []          # engine output, edited in place
        self.result = None           # analyze_chunks() result
        self._iid_by_pos = {}
        self._busy = False

        if isinstance(root, (tk.Tk, tk.Toplevel)):
            root.title("Chunk Triage & Bulk Review")
            root.geometry("1180x760")
            root.minsize(940, 600)

        self._build_ui()
        self._run_triage()

    # ----------------------------------------------------------------- UI -- #
    def _build_ui(self):
        top = ttk.Frame(self.root, padding=(10, 8))
        top.pack(fill="x")
        self.summary_var = tk.StringVar(value="Analyzing chunks…")
        ttk.Label(top, textvariable=self.summary_var,
                  font=("Arial", 11, "bold")).pack(side="left")
        self.source_var = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.source_var,
                  foreground="#555555").pack(side="left", padx=12)

        bar = ttk.Frame(self.root, padding=(10, 0))
        bar.pack(fill="x")
        ttk.Label(bar, text="Show:").pack(side="left")
        self.filter_var = tk.StringVar(value="Needs review")
        self.filter_box = ttk.Combobox(bar, state="readonly", width=14,
                                       values=list(_FILTERS),
                                       textvariable=self.filter_var)
        self.filter_box.pack(side="left", padx=4)
        self.filter_box.bind("<<ComboboxSelected>>", lambda _e: self._refresh())
        ttk.Label(bar, text="Search:").pack(side="left", padx=(12, 0))
        self.search = ttk.Entry(bar, width=26)
        self.search.pack(side="left", padx=4)
        self.search.bind("<Return>", lambda _e: self._refresh())
        ttk.Button(bar, text="Find", command=self._refresh).pack(side="left")
        ttk.Button(bar, text="Clear",
                   command=self._clear_search).pack(side="left", padx=4)

        # ---- the table of proposals
        mid = ttk.Frame(self.root, padding=(10, 6))
        mid.pack(fill="both", expand=True)
        cols = ("action", "heading", "pages", "labels", "reason", "text")
        self.tree = ttk.Treeview(mid, columns=cols, show="tree headings",
                                 selectmode="extended")
        self.tree.heading("#0", text="Chunk")
        self.tree.heading("action", text="Decision")
        self.tree.heading("heading", text="Section heading")
        self.tree.heading("pages", text="Page")
        self.tree.heading("labels", text="Type")
        self.tree.heading("reason", text="Why")
        self.tree.heading("text", text="Text preview")
        self.tree.column("#0", width=70, anchor="w", stretch=False)
        self.tree.column("action", width=70, anchor="w", stretch=False)
        self.tree.column("heading", width=210, anchor="w")
        self.tree.column("pages", width=50, anchor="w", stretch=False)
        self.tree.column("labels", width=90, anchor="w", stretch=False)
        self.tree.column("reason", width=250, anchor="w")
        self.tree.column("text", width=320, anchor="w")
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")
        for action, (colour, _bg) in _ACTION_TAGS.items():
            self.tree.tag_configure(action, foreground=colour)
        self.tree.tag_configure("edited", font=("Arial", 10, "bold"))
        self.tree.bind("<Double-1>", lambda _e: self._edit_selected())

        # ---- bulk actions
        act = ttk.LabelFrame(self.root, text=" Selected chunks ", padding=6)
        act.pack(fill="x", padx=10, pady=(0, 4))
        ttk.Button(act, text="Keep",
                   command=lambda: self._set_action("log")).pack(side="left")
        ttk.Button(act, text="Skip",
                   command=lambda: self._set_action("skip")).pack(side="left",
                                                                  padx=4)
        ttk.Button(act, text="Set heading…",
                   command=self._set_heading).pack(side="left", padx=4)
        ttk.Button(act, text="Use heading above",
                   command=self._use_previous_heading).pack(side="left", padx=4)
        ttk.Button(act, text="Edit chunk…",
                   command=self._edit_selected).pack(side="left", padx=4)
        ttk.Label(act, foreground="#666666",
                  text="  (double-click a row to edit it; nothing is written "
                       "until you accept)").pack(side="left", padx=8)

        bottom = ttk.Frame(self.root, padding=(10, 6))
        bottom.pack(fill="x")
        self.accept_btn = ttk.Button(bottom, text="Accept & save",
                                     command=self._accept)
        self.accept_btn.pack(side="left")
        ttk.Button(bottom, text="Preview merged sections…",
                   command=self._preview_merged).pack(side="left", padx=6)
        ttk.Button(bottom, text="Re-run triage",
                   command=self._run_triage).pack(side="left", padx=6)
        ttk.Button(bottom, text="Close",
                   command=self.root.destroy).pack(side="right")
        self.status_var = tk.StringVar(value="")
        ttk.Label(bottom, textvariable=self.status_var,
                  foreground="#555555").pack(side="left", padx=12)

    # ------------------------------------------------------------ triage -- #
    def _run_triage(self):
        """Run the engine off the UI thread (an LLM pass may be involved)."""
        if self._busy:
            return
        self._busy = True
        self.summary_var.set("Analyzing chunks…")
        self.accept_btn.configure(state="disabled")

        def work():
            try:
                result = analyze_chunks(self.chunks_data, llm=self._llm,
                                        log=self._log)
            except Exception as exc:  # noqa: BLE001
                self.root.after(0, lambda e=exc: self._triage_done(None, e))
            else:
                self.root.after(0, lambda r=result: self._triage_done(r, None))

        threading.Thread(target=work, daemon=True).start()

    def _triage_done(self, result, exc):
        self._busy = False
        self.accept_btn.configure(state="normal")
        if exc is not None:
            self._log(f"[ERROR] triage failed: {exc}")
            messagebox.showerror("Triage failed", str(exc))
            self.summary_var.set("Triage failed.")
            return
        self.result = result
        self.proposals = result["proposals"]
        s = result["summary"]
        self.summary_var.set(
            f"{s['total']} chunks — {s['log']} to keep, {s['skip']} to skip, "
            f"{s['review']} need review ({s['auto_handled_pct']}% decided "
            "automatically)")
        ref = result["refinement"]
        where = {"toc": "Table of Contents", "llm": "LLM heading check",
                 "rules": "rule-based"}.get(ref["source"], ref["source"])
        self.source_var.set(f"sections from: {where} ({ref['note']})")
        # Land on whatever needs attention; if nothing does, show everything.
        self.filter_var.set("Needs review" if s["review"] else "All")
        self._refresh()

    def _log(self, msg):
        if self.logger:
            try:
                self.logger.info(msg)
            except Exception:  # noqa: BLE001 - logging must never break review
                pass

    # ------------------------------------------------------------- table -- #
    def _visible(self):
        choice = self.filter_var.get()
        needle = self.search.get().strip().casefold()
        out = []
        for p in self.proposals:
            if choice == "Needs review" and p["action"] != "review":
                continue
            if choice == "Keep" and p["action"] != "log":
                continue
            if choice == "Skip" and p["action"] != "skip":
                continue
            if needle and needle not in (
                    f"{p['heading']} {p['reason']} {p['text']}".casefold()):
                continue
            out.append(p)
        return out

    def _clear_search(self):
        self.search.delete(0, tk.END)
        self._refresh()

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        self._iid_by_pos = {}
        for p in self._visible():
            preview = " ".join(str(p["text"]).split())[:160]
            pages = ", ".join(str(x) for x in p["page_num"][:3])
            labels = ", ".join(sorted(set(p["labels"])))[:28]
            tags = [p["action"]]
            if p.get("edited"):
                tags.append("edited")
            iid = self.tree.insert(
                "", "end", text=str(p["chunk_index"]),
                values=(_ACTION_LABEL.get(p["action"], p["action"]),
                        p["heading"] or "—", pages, labels, p["reason"],
                        preview),
                tags=tuple(tags))
            self._iid_by_pos[iid] = p
        shown = len(self._iid_by_pos)
        self.status_var.set(f"{shown} row(s) shown of {len(self.proposals)}")

    def _selected(self):
        return [self._iid_by_pos[i] for i in self.tree.selection()
                if i in self._iid_by_pos]

    # ----------------------------------------------------------- editing -- #
    def _set_action(self, action):
        picked = self._selected()
        if not picked:
            messagebox.showinfo("Nothing selected",
                                "Select one or more rows first.")
            return
        for p in picked:
            p["action"] = action
            p["edited"] = True
            p["reason"] = f"set to '{_ACTION_LABEL[action]}' by reviewer"
        self._log(f"{len(picked)} chunk(s) set to {action} by the reviewer.")
        self._refresh()

    def _set_heading(self):
        picked = self._selected()
        if not picked:
            messagebox.showinfo("Nothing selected",
                                "Select one or more rows first.")
            return
        known = []
        for p in self.proposals:
            if p["heading"] and p["heading"] not in known:
                known.append(p["heading"])
        value = self._ask_heading(picked[0]["heading"], known)
        if value is None:
            return
        for p in picked:
            p["heading"] = value
            p["action"] = "log"
            p["edited"] = True
            p["reason"] = "heading set by reviewer"
        self._log(f"{len(picked)} chunk(s) assigned heading '{value}'.")
        self._refresh()

    def _use_previous_heading(self):
        """Give the selected chunks the heading of the nearest preceding kept
        chunk — the bulk version of the old 'Use Prev Heading' button, and it
        genuinely looks backwards in document order."""
        picked = self._selected()
        if not picked:
            messagebox.showinfo("Nothing selected",
                                "Select one or more rows first.")
            return
        by_pos = {p["position"]: p for p in self.proposals}
        missing = 0
        for p in picked:
            heading = ""
            for pos in range(p["position"] - 1, -1, -1):
                prev = by_pos.get(pos)
                if prev and prev["action"] == "log" and prev["heading"]:
                    heading = prev["heading"]
                    break
            if not heading:
                missing += 1
                continue
            p["heading"] = heading
            p["action"] = "log"
            p["edited"] = True
            p["reason"] = f"heading taken from the chunk above ('{heading}')"
        self._refresh()
        if missing:
            messagebox.showinfo(
                "No heading above",
                f"{missing} of the selected chunks have no kept chunk with a "
                "heading before them, so they were left unchanged.")

    def _ask_heading(self, initial, known):
        """Small modal: type a heading or pick one already used."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Section heading")
        dlg.transient(self.root.winfo_toplevel())
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Heading for the selected chunk(s):").pack(anchor="w")
        box = ttk.Combobox(frm, values=known, width=56)
        box.set(initial or "")
        box.pack(fill="x", pady=8)
        box.focus_set()
        out = {"value": None}

        def ok():
            out["value"] = box.get().strip()
            dlg.destroy()

        row = ttk.Frame(frm)
        row.pack(fill="x")
        ttk.Button(row, text="Apply", command=ok).pack(side="left")
        ttk.Button(row, text="Cancel", command=dlg.destroy).pack(side="left",
                                                                 padx=6)
        dlg.bind("<Return>", lambda _e: ok())
        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        dlg.grab_set()
        self.root.wait_window(dlg)
        return out["value"]

    def _edit_selected(self):
        picked = self._selected()
        if not picked:
            return
        p = picked[0]
        dlg = tk.Toplevel(self.root)
        dlg.title(f"Chunk {p['chunk_index']}")
        dlg.transient(self.root.winfo_toplevel())
        dlg.geometry("760x560")
        frm = ttk.Frame(dlg, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, foreground="#555555",
                  text=f"Page {p['page_num']} · {', '.join(p['labels']) or 'n/a'}"
                       f" · {p['reason']}").pack(anchor="w")
        ttk.Label(frm, text="Heading:").pack(anchor="w", pady=(8, 0))
        head = ttk.Entry(frm)
        head.insert(0, p["heading"])
        head.pack(fill="x")
        ttk.Label(frm, text="Text:").pack(anchor="w", pady=(8, 0))
        text = tk.Text(frm, wrap="word", height=20)
        text.insert("1.0", p["text"])
        text.pack(fill="both", expand=True)

        def save(action):
            p["heading"] = head.get().strip()
            new_text = text.get("1.0", "end").rstrip("\n")
            if new_text != p["text"]:
                p["text"] = new_text
                p["status"] = "logged (edited)"
            p["action"] = action
            p["edited"] = True
            p["reason"] = "edited by reviewer"
            dlg.destroy()
            self._refresh()

        row = ttk.Frame(frm)
        row.pack(fill="x", pady=(8, 0))
        ttk.Button(row, text="Save & keep",
                   command=lambda: save("log")).pack(side="left")
        ttk.Button(row, text="Save & skip",
                   command=lambda: save("skip")).pack(side="left", padx=6)
        ttk.Button(row, text="Cancel", command=dlg.destroy).pack(side="left")
        dlg.grab_set()
        self.root.wait_window(dlg)

    # ----------------------------------------------------------- outputs -- #
    def _preview_merged(self):
        payload = build_output_payload(self.proposals)
        sections = payload["merged_headings"]
        dlg = tk.Toplevel(self.root)
        dlg.title(f"Merged sections ({len(sections)})")
        dlg.transient(self.root.winfo_toplevel())
        dlg.geometry("820x600")
        frm = ttk.Frame(dlg, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, font=("Arial", 10, "bold"),
                  text=f"{len(sections)} section(s) will be written to "
                       f"{os.path.basename(self.output_file_name)}").pack(
            anchor="w")
        box = tk.Text(frm, wrap="word")
        box.pack(fill="both", expand=True, pady=6)
        for sec in sections:
            heading = sec["heading"][0] if sec["heading"] else "(no heading)"
            body = " ".join(str(sec["merged_text"]).split())[:400]
            box.insert("end", f"■ {heading}\n   chunks "
                              f"{sec['chunk_indices']}, pages "
                              f"{sec['page_nums']}\n   {body}…\n\n")
        box.configure(state="disabled")
        ttk.Button(frm, text="Close", command=dlg.destroy).pack(anchor="e")

    def _accept(self):
        """Write the review output — once, not per chunk."""
        payload = build_output_payload(self.proposals)
        sections = payload["merged_headings"]
        kept = sum(1 for p in self.proposals if p["action"] != "skip")
        still = sum(1 for p in self.proposals if p["action"] == "review")
        if still and not messagebox.askyesno(
                "Unreviewed chunks",
                f"{still} chunk(s) are still marked 'Review'. They will be "
                "kept under the heading shown (or with none).\n\nSave anyway?"):
            return
        try:
            with open(self.output_file_name, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=4, ensure_ascii=False)
        except Exception as exc:  # noqa: BLE001
            self._log(f"[ERROR] could not write {self.output_file_name}: {exc}")
            messagebox.showerror("Save failed", str(exc))
            return
        self._log(f"Triage accepted: {kept} chunk(s) kept in "
                  f"{len(sections)} section(s) → {self.output_file_name}")
        messagebox.showinfo(
            "Review saved",
            f"{kept} chunk(s) kept in {len(sections)} section(s).\n\n"
            f"Written to {self.output_file_name}\n\nProceeding to Section "
            "Review…")
        self.root.destroy()
        if self.on_complete_callback:
            self.on_complete_callback()


def unreviewed_headings(proposals):
    """Distinct headings still marked 'review' — handy for tests/logging."""
    seen, out = set(), []
    for p in proposals:
        if p["action"] != "review":
            continue
        key = normalize_heading(p["heading"])
        if key not in seen:
            seen.add(key)
            out.append(p["heading"])
    return out
