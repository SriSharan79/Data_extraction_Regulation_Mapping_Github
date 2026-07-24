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

Besides *Keep* and *Skip* a chunk can be decided **Merge ↑**: its text is
folded into the chunk above it, which is how a section Docling split across
two chunks is put back together. Merging asks every time whether the chunk's
heading should be written as a line before its text — a real sub-heading
wants that, a paragraph broken mid-sentence must not get one. The decision is
recorded, not applied destructively, so it stays visible in the table, is
undone by setting the chunk back to *Keep*, and survives closing and
resuming the session. It is available from the toolbar (bulk, on the
selection) and from the *Edit chunk* window (*Save & merge into above*).
"""

from __future__ import annotations

import json
import os
import threading
import time
import tkinter as tk
from tkinter import messagebox, ttk

from .chunk_triage import (TOC_SOURCE_LABEL, analyze_chunks,
                           build_output_payload, find_merge_target,
                           headings_match, normalize_heading, triage_summary)

_ACTION_LABEL = {"log": "Keep", "skip": "Skip", "review": "Review",
                 "merge": "Merge ↑"}
_ACTION_TAGS = {"log": ("#1f7a1f", ""), "skip": ("#777777", ""),
                "review": ("#b00020", ""), "merge": ("#0b6fa4", "")}
_FILTERS = ("Needs review", "All", "Keep", "Skip", "Merged")

# Fields a merge decision adds to a proposal — dropped again the moment the
# chunk is given any other decision, so a stale merge can never be written.
_MERGE_FIELDS = ("merge_add_heading", "merge_heading_text", "merged_into")


class ChunkTriageApp:
    """Bulk review of Docling chunks with automatic triage proposals."""

    def __init__(self, root, chunks_data, output_file_name, logger,
                 on_complete_callback=None, llm=None, prior_history=None,
                 prior_path=None, version_on_change=False,
                 destroy_on_accept=True, pdf_path=None, tables=None):
        self.root = root
        self.chunks_data = list(chunks_data or [])
        # With the source PDF at hand the TOC is read from the PDF itself
        # (embedded outline / printed TOC text) and headings are verified
        # against it. When the PDF offers none, the cascade carries on with
        # the chunks, then with the tables extracted from the contents pages
        # (``tables`` — the extractor's records from the cache), and only
        # then asks the LLM.
        self.pdf_path = pdf_path
        self.tables = list(tables or [])
        self.output_file_name = output_file_name
        self.logger = logger
        self.on_complete_callback = on_complete_callback
        self._llm = llm
        # Decisions saved by a previous session (the output file's
        # raw_session_history, either tool's shape). They are laid over the
        # fresh triage proposals on resume, so earlier accepts/skips/edits
        # survive closing the window.
        self._prior_history = list(prior_history or [])
        # Where that history came from, plus its fingerprint: with
        # ``version_on_change`` an accept that changed nothing keeps the
        # existing review file, while real changes are written to a NEW file
        # (timestamped when the target already exists) instead of
        # overwriting the earlier review.
        self._prior_path = prior_path
        self._prior_fp = (self._history_fingerprint(self._prior_history)
                          if self._prior_history else None)
        self.version_on_change = version_on_change
        # False when the review lives inside a tab (the extraction tab):
        # accepting saves and stays, rather than destroying the host.
        self.destroy_on_accept = destroy_on_accept
        self.saved_path = None       # where the last accept actually wrote

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

    @staticmethod
    def _history_fingerprint(entries):
        """Canonical form of the *meaningful* review state — per chunk: kept,
        skipped or merged into the chunk above (and whether its heading goes
        into the merged text), its heading and its text. Used to decide
        whether the reviewer actually changed anything compared to the prior
        review (bookkeeping fields like triage_reason are deliberately
        excluded). Reviews written before merging existed simply have the
        merge fields absent, which reads back as 'not merged'."""
        out = []
        for entry in entries or []:
            if not isinstance(entry, dict):
                continue
            heading = entry.get("heading")
            if isinstance(heading, list):
                heading = ", ".join(str(h) for h in heading if str(h).strip())
            status = str(entry.get("status") or "")
            out.append((str(entry.get("chunk_index")),
                        status == "skipped",
                        status.startswith("merged"),
                        bool(entry.get("merge_add_heading")),
                        str(entry.get("merge_heading_text") or ""),
                        str(heading or ""),
                        str(entry.get("chunk_text") or "")))
        return tuple(sorted(out))

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
        cols = ("action", "toc", "heading", "pages", "labels", "reason", "text")
        self.tree = ttk.Treeview(mid, columns=cols, show="tree headings",
                                 selectmode="extended")
        self.tree.heading("#0", text="Chunk")
        self.tree.heading("action", text="Decision")
        self.tree.heading("toc", text="TOC")
        self.tree.heading("heading", text="Section heading")
        self.tree.heading("pages", text="Page")
        self.tree.heading("labels", text="Type")
        self.tree.heading("reason", text="Why")
        self.tree.heading("text", text="Text preview")
        self.tree.column("#0", width=70, anchor="w", stretch=False)
        self.tree.column("action", width=82, anchor="w", stretch=False)
        self.tree.column("toc", width=45, anchor="center", stretch=False)
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
        ttk.Button(act, text="Merge into above ↑",
                   command=self._merge_into_above).pack(side="left", padx=4)
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
                                        log=self._log,
                                        pdf_path=self.pdf_path,
                                        tables=self.tables)
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
        restored = self._apply_prior_history()
        s = triage_summary(self.proposals)
        self.summary_var.set(
            f"{s['total']} chunks — {s['log']} to keep, {s['skip']} to skip, "
            f"{s['review']} need review ({s['auto_handled_pct']}% decided "
            "automatically)"
            + (f", {s['merge']} merged upwards" if s.get("merge") else "")
            + (f" — {restored} restored from the previous session"
               if restored else ""))
        ref = result["refinement"]
        where = {"toc": "Table of Contents", "llm": "LLM heading check",
                 "rules": "rule-based"}.get(ref["source"], ref["source"])
        if ref["source"] == "toc":
            # Say WHICH of the five TOC sources the cascade settled on.
            src = (result.get("toc") or {}).get("source")
            where = ("Table of Contents — "
                     + TOC_SOURCE_LABEL.get(src, src or "unknown source"))
        extra = ""
        if (result.get("toc") or {}).get("entries"):
            matched = sum(1 for p in self.proposals
                          if p.get("toc_match") is True)
            off = sum(1 for p in self.proposals
                      if p.get("toc_match") is False)
            extra = f" — TOC-verified: {matched} ✓ / {off} ✗"
        self.source_var.set(f"sections from: {where} ({ref['note']}){extra}")
        # Land on whatever needs attention; if nothing does, show everything.
        self.filter_var.set("Needs review" if s["review"] else "All")
        self._refresh()

    def _apply_prior_history(self):
        """Lay a previous session's saved decisions over the fresh triage
        proposals (matched by chunk index): action, heading and any edited
        text are restored, so resuming continues where the reviewer stopped
        instead of re-deciding everything. Returns how many were restored."""
        if not self._prior_history:
            return 0
        by_index = {p["chunk_index"]: p for p in self.proposals}
        restored = 0
        for entry in self._prior_history:
            if not isinstance(entry, dict):
                continue
            p = by_index.get(entry.get("chunk_index"))
            if p is None:
                continue
            heading = entry.get("heading")
            if isinstance(heading, list):
                heading = ", ".join(str(h) for h in heading if str(h).strip())
            status = str(entry.get("status") or "")
            self._clear_merge(p)
            if status == "skipped":
                p["action"] = "skip"
            elif status.startswith("merged"):
                # The history keeps each chunk's own text, so re-applying the
                # merge on accept reproduces exactly the previous output.
                p["action"] = "merge"
                p["merge_add_heading"] = bool(entry.get("merge_add_heading"))
                if entry.get("merge_heading_text"):
                    p["merge_heading_text"] = entry["merge_heading_text"]
                if entry.get("merged_into") is not None:
                    p["merged_into"] = entry["merged_into"]
            else:
                p["action"] = "log"
            p["heading"] = str(heading or "")
            if entry.get("chunk_text") is not None:
                p["text"] = entry["chunk_text"]
            if status and not status.startswith("merged"):
                p["status"] = status      # keeps 'logged (edited)' etc.
            p["edited"] = True
            p["reason"] = ("merged into the chunk above — restored from the "
                           "previous session" if p["action"] == "merge"
                           else "restored from the previous session")
            restored += 1
        if restored:
            self._log(f"Resume: {restored} decision(s) restored from the "
                      "previous session.")
        return restored

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
            if choice == "Merged" and p["action"] != "merge":
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
            toc_cell = {True: "✓", False: "✗"}.get(p.get("toc_match"), "")
            iid = self.tree.insert(
                "", "end", text=str(p["chunk_index"]),
                values=(_ACTION_LABEL.get(p["action"], p["action"]),
                        toc_cell, p["heading"] or "—", pages, labels,
                        p["reason"], preview),
                tags=tuple(tags))
            self._iid_by_pos[iid] = p
        shown = len(self._iid_by_pos)
        merged = sum(1 for p in self.proposals if p["action"] == "merge")
        self.status_var.set(
            f"{shown} row(s) shown of {len(self.proposals)}"
            + (f" · {merged} merged into the chunk above" if merged else ""))

    def _selected(self):
        return [self._iid_by_pos[i] for i in self.tree.selection()
                if i in self._iid_by_pos]

    # ----------------------------------------------------------- editing -- #
    @staticmethod
    def _clear_merge(proposal):
        """Drop any merge state from a proposal — called whenever it is given
        a different decision, so 'Keep' after 'Merge ↑' really is a keep."""
        for field in _MERGE_FIELDS:
            proposal.pop(field, None)
        if str(proposal.get("status") or "").startswith("merged"):
            proposal.pop("status", None)

    def _positions(self):
        """Map from proposal identity to its position in document order —
        what 'the chunk above' is resolved against."""
        return {id(p): i for i, p in enumerate(self.proposals)}

    def _set_action(self, action):
        picked = self._selected()
        if not picked:
            messagebox.showinfo("Nothing selected",
                                "Select one or more rows first.")
            return
        for p in picked:
            p["action"] = action
            self._clear_merge(p)
            p["edited"] = True
            p["toc_match"] = None      # reviewer decision, not TOC-verified
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
            self._clear_merge(p)
            p["edited"] = True
            p["toc_match"] = None
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
            self._clear_merge(p)
            p["edited"] = True
            p["toc_match"] = None
            p["reason"] = f"heading taken from the chunk above ('{heading}')"
        self._refresh()
        if missing:
            messagebox.showinfo(
                "No heading above",
                f"{missing} of the selected chunks have no kept chunk with a "
                "heading before them, so they were left unchanged.")

    # ------------------------------------------------------------ merging -- #
    def _merge_into_above(self):
        """Fold the selected chunks into the chunk above them.

        'Above' is the nearest preceding chunk that is kept in its own right:
        skipped chunks and chunks that are themselves merged are passed over,
        so selecting a whole run of fragments merges all of them into the one
        section that starts the run. The reviewer is asked once whether the
        heading should be written into the merged text — a section split
        across two chunks usually wants it, a paragraph split mid-sentence
        never does. Nothing is written until *Accept & save*, and setting the
        chunk back to Keep undoes the merge."""
        picked = self._selected()
        if not picked:
            messagebox.showinfo("Nothing selected",
                                "Select one or more rows first.")
            return
        pos = self._positions()
        picked = sorted(picked, key=lambda p: pos.get(id(p), 0))

        first_target = None
        for p in picked:
            first_target = find_merge_target(self.proposals, pos[id(p)])
            if first_target is not None:
                break
        if first_target is None:
            messagebox.showinfo(
                "Nothing above",
                "There is no kept chunk above the selected chunk(s), so there "
                "is nothing to merge into. Nothing was changed.")
            return

        options = self._ask_merge_options(picked, first_target)
        if options is None:
            return

        merged = orphaned = 0
        for p in picked:
            # Resolved chunk by chunk, after the previous ones were marked:
            # a run of consecutive merges therefore all lands in the same
            # chunk instead of chaining into each other.
            target = find_merge_target(self.proposals, pos[id(p)])
            if target is None:
                orphaned += 1
                continue
            self._apply_merge(p, target, options)
            merged += 1
        self._log(f"{merged} chunk(s) merged into the chunk above "
                  f"({'with' if options['add_heading'] else 'without'} the "
                  "heading in the text).")
        self._refresh()
        if orphaned:
            messagebox.showinfo(
                "Partly applied",
                f"{orphaned} of the selected chunks have no kept chunk above "
                "them, so they were left unchanged.")

    def _apply_merge(self, proposal, target, options):
        """Mark one chunk as merged into ``target`` with the chosen heading
        handling. Purely a decision — the text is folded when the review is
        built, which is what keeps it reversible."""
        add = bool(options.get("add_heading"))
        override = (options.get("heading_text") or "").strip()
        proposal["action"] = "merge"
        proposal["edited"] = True
        proposal["toc_match"] = None
        proposal["merge_add_heading"] = add
        if add and override:
            proposal["merge_heading_text"] = override
        else:
            proposal.pop("merge_heading_text", None)
        proposal["merged_into"] = target.get("chunk_index")
        shown = override or str(proposal.get("heading") or "").strip()
        proposal["reason"] = (
            f"merged into chunk {target.get('chunk_index')}"
            + (f" — heading '{shown}' added to the text" if add and shown
               else " — text only"))
        if str(proposal.get("status") or "").startswith("merged"):
            proposal.pop("status", None)

    def _ask_merge_options(self, picked, target, parent=None):
        """Ask whether the merged text gets the chunk's heading in front of
        it. Returns ``{"add_heading": bool, "heading_text": str|None}`` or
        None when the reviewer cancels."""
        single = len(picked) == 1
        own_heading = (str(picked[0].get("heading") or "").strip()
                       if single else "")
        target_heading = str((target or {}).get("heading") or "").strip()
        # Proposed answer: add the heading when this chunk carries one of its
        # own that is not already the heading of the chunk it goes into
        # (that would just repeat it). Bulk merges default to text only.
        default_add = bool(single and own_heading and not (
            target_heading and headings_match(own_heading, target_heading)))

        host = parent or self.root
        dlg = tk.Toplevel(host)
        dlg.title("Merge into the chunk above")
        dlg.transient(host.winfo_toplevel())
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)

        if single:
            headline = (f"Chunk {picked[0].get('chunk_index')} will be merged "
                        f"into chunk {(target or {}).get('chunk_index')}"
                        + (f" — {target_heading}" if target_heading else ""))
        else:
            headline = (f"{len(picked)} chunks will be merged into the kept "
                        "chunk above each of them.")
        ttk.Label(frm, text=headline, wraplength=470, justify="left",
                  font=("Arial", 10, "bold")).pack(anchor="w")
        ttk.Label(frm, wraplength=470, justify="left", foreground="#555555",
                  text="Their text is appended to that chunk, so they stop "
                       "being separate chunks in the saved review.").pack(
            anchor="w", pady=(2, 12))

        ttk.Label(frm, text="Add the heading to the merged text?",
                  font=("Arial", 10, "bold")).pack(anchor="w")
        add_var = tk.BooleanVar(value=default_add)
        ttk.Radiobutton(
            frm, variable=add_var, value=True,
            text="Yes — write the heading as a line before the text").pack(
            anchor="w", pady=(2, 0))
        ttk.Radiobutton(
            frm, variable=add_var, value=False,
            text="No — append the text only").pack(anchor="w")

        entry = None
        if single:
            ttk.Label(frm, text="Heading to write:").pack(anchor="w",
                                                          pady=(10, 0))
            entry = ttk.Entry(frm, width=58)
            entry.insert(0, own_heading)
            entry.pack(fill="x")
            ttk.Label(frm, foreground="#555555", wraplength=470,
                      justify="left",
                      text="(only used when 'Yes' is selected)").pack(
                anchor="w")
        else:
            ttk.Label(frm, foreground="#555555", wraplength=470,
                      justify="left",
                      text="Each chunk's own heading is used; chunks without "
                           "one contribute their text only.").pack(
                anchor="w", pady=(10, 0))

        out = {"value": None}

        def ok():
            out["value"] = {
                "add_heading": bool(add_var.get()),
                "heading_text": (entry.get().strip()
                                 if entry is not None else None),
            }
            dlg.destroy()

        row = ttk.Frame(frm)
        row.pack(fill="x", pady=(14, 0))
        ttk.Button(row, text="Merge", command=ok).pack(side="left")
        ttk.Button(row, text="Cancel",
                   command=dlg.destroy).pack(side="left", padx=6)
        dlg.bind("<Return>", lambda _e: ok())
        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        dlg.grab_set()
        host.wait_window(dlg)
        return out["value"]

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
        state = _ACTION_LABEL.get(p["action"], p["action"])
        if p["action"] == "merge":
            state += (f" into chunk {p.get('merged_into')}"
                      + (" · heading written into the text"
                         if p.get("merge_add_heading") else " · text only"))
        ttk.Label(frm, foreground="#555555",
                  text=f"Current decision: {state}").pack(anchor="w")
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
            if action == "merge":
                # Same decision as the toolbar button, with the edits above
                # already applied — so the heading typed here is the one
                # offered for the merged text.
                target = find_merge_target(self.proposals,
                                           self._positions()[id(p)])
                if target is None:
                    messagebox.showinfo(
                        "Nothing above",
                        "There is no kept chunk above this one, so there is "
                        "nothing to merge into.", parent=dlg)
                    return                      # keep the editor open
                options = self._ask_merge_options([p], target, parent=dlg)
                if options is None:
                    return                      # cancelled — keep editing
                self._apply_merge(p, target, options)
            else:
                p["action"] = action
                self._clear_merge(p)
                p["edited"] = True
                p["toc_match"] = None
                p["reason"] = "edited by reviewer"
            dlg.destroy()
            self._refresh()

        row = ttk.Frame(frm)
        row.pack(fill="x", pady=(8, 0))
        ttk.Button(row, text="Save & keep",
                   command=lambda: save("log")).pack(side="left")
        ttk.Button(row, text="Save & merge into above ↑",
                   command=lambda: save("merge")).pack(side="left", padx=6)
        ttk.Button(row, text="Save & skip",
                   command=lambda: save("skip")).pack(side="left", padx=(0, 6))
        ttk.Button(row, text="Cancel", command=dlg.destroy).pack(side="left")
        dlg.deiconify()           # Ensure the window is shown
        dlg.update_idletasks()    # Force Tkinter to draw the window before grabbing
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
        """Write the review output — once, not per chunk.

        With ``version_on_change``: when a prior review was pre-applied and
        the reviewer changed nothing, the existing file is kept as-is; when
        anything changed, the review is written as a NEW file (a timestamped
        name when the target already exists) so the earlier review survives.
        """
        payload = build_output_payload(self.proposals)
        sections = payload["merged_headings"]
        kept = sum(1 for p in self.proposals
                   if p["action"] not in ("skip", "merge"))
        merged = sum(1 for p in self.proposals if p["action"] == "merge")
        still = sum(1 for p in self.proposals if p["action"] == "review")
        if still and not messagebox.askyesno(
                "Unreviewed chunks",
                f"{still} chunk(s) are still marked 'Review'. They will be "
                "kept under the heading shown (or with none).\n\nSave anyway?"):
            return

        fp_new = self._history_fingerprint(payload["raw_session_history"])
        unchanged = (self.version_on_change and self._prior_fp is not None
                     and fp_new == self._prior_fp and self._prior_path
                     and os.path.exists(self._prior_path))
        if unchanged:
            self.saved_path = self._prior_path
            self._log("Triage accepted with no changes — existing review "
                      f"kept: {self._prior_path}")
            messagebox.showinfo(
                "No changes",
                f"Nothing was changed compared to the existing review —\n"
                f"{self._prior_path}\n\nis kept as it is. Proceeding to "
                "Section Review…")
        else:
            target = self.output_file_name
            if self.version_on_change and os.path.exists(target):
                stem, ext = os.path.splitext(target)
                target = f"{stem} {time.strftime('%H.%M.%S')}{ext}"
                while os.path.exists(target):      # same-second re-save
                    stem2 = target[: -len(ext)]
                    target = f"{stem2}b{ext}"
            try:
                os.makedirs(os.path.dirname(target) or ".", exist_ok=True)
                with open(target, "w", encoding="utf-8") as f:
                    json.dump(payload, f, indent=4, ensure_ascii=False)
            except Exception as exc:  # noqa: BLE001
                self._log(f"[ERROR] could not write {target}: {exc}")
                messagebox.showerror("Save failed", str(exc))
                return
            self.saved_path = target
            as_new = ("" if not self._prior_path
                      or os.path.abspath(self._prior_path)
                      == os.path.abspath(target)
                      else " (saved as a NEW file — the previous review is "
                           "untouched)")
            merged_note = (f" ({merged} merged into the chunk above)"
                           if merged else "")
            self._log(f"Triage accepted: {kept} chunk(s) kept in "
                      f"{len(sections)} section(s){merged_note} → "
                      f"{target}{as_new}")
            messagebox.showinfo(
                "Review saved",
                f"{kept} chunk(s) kept in {len(sections)} "
                f"section(s){merged_note}.\n\n"
                f"Written to {target}{as_new}\n\nProceeding to Section "
                "Review…")
            # the just-saved state is the new baseline for further accepts
            self._prior_fp = fp_new
            self._prior_path = target
        if self.destroy_on_accept:
            self.root.destroy()
        else:
            self.status_var.set(f"Saved → {os.path.basename(self.saved_path)}")
        if self.on_complete_callback:
            try:
                self.on_complete_callback(self.saved_path)
            except TypeError:
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