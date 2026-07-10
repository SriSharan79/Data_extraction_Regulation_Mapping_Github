"""
AI review UI for the docling chunking outputs.

Loads either a chunks cache (``*_docling_chunks_cache.json`` — the file the
extraction pipeline writes, ``{"chunks": [...]}`` or a bare chunk list) or a
processed review output (``Processed_chunks.json`` with ``merged_headings``)
and offers the same full AI Review workbench as the EASA studio
(`data_extraction.ai_utils.review_panel.AIReviewMixin`): a shared sections
queue, free-form review with presets/answer formats and per-result auto-save,
and column analysis with unique-element checkboxes, live prompt preview and
row-by-row saving into an accumulating Excel workbook.

Sections are listed on the left with ✓ checkboxes (click the ✓ cell or press
Space; the ✓ heading and the Select all / Clear checks buttons act on all
visible rows; a search box filters the list).

Run standalone:
    python -m data_extraction.chunking.ai_review_ui [path/to/chunks.json]

Also embeddable: pass a container Frame as ``root`` (the Data Extraction
Studio hosts it this way).
"""

import json
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from data_extraction.ai_utils.review_panel import AIReviewMixin


def _chunk_title(chunk, idx):
    headings = chunk.get("heading") or chunk.get("meta", {}).get("headings") or []
    if isinstance(headings, str):
        headings = [headings]
    label = " → ".join(str(h) for h in headings) if headings else "Untitled"
    number = chunk.get("chunk_index", idx + 1)
    return f"[{number}] {label}"


def _section_title(section, idx):
    heading = section.get("heading", [])
    if isinstance(heading, list):
        label = " → ".join(str(h) for h in heading) if heading else "Untitled"
    else:
        label = str(heading) if heading else "Untitled"
    return f"[{idx + 1}] {label}"


def sections_from_payload(data):
    """Normalize a chunks cache or a processed review output to
    [(title, text), ...]. Returns [] when the payload has neither shape."""
    if isinstance(data, dict) and data.get("merged_headings"):
        return [(_section_title(s, i), s.get("merged_text", "") or "")
                for i, s in enumerate(data["merged_headings"])]
    chunks = data.get("chunks") if isinstance(data, dict) else data
    if isinstance(chunks, list) and chunks:
        return [(_chunk_title(c, i), c.get("chunk_text") or c.get("text") or "")
                for i, c in enumerate(chunks) if isinstance(c, dict)]
    return []


class ChunkAIReviewApp(AIReviewMixin):
    """Sections list (chunks / merged sections) + the shared AI Review page."""

    def __init__(self, root, json_path: str = None, logger=None):
        self.root = root
        self.logger = logger
        if isinstance(root, (tk.Tk, tk.Toplevel)):
            root.title("Chunk AI Review")
            root.geometry("1280x820")
            root.minsize(960, 600)

        self.json_path = None
        self._sections = []           # [(title, text), ...] in document order
        self._checked = set()         # indices into _sections, survive filtering
        self.idx_by_iid = {}          # tree row -> section index
        self._init_ai_state()         # AI Review state (AIReviewMixin)

        self._build_ui()

        if json_path:
            self._load(json_path)

    # ------------------------------------------------------------------ UI -- #
    def _build_ui(self):
        bar = ttk.Frame(self.root, padding=(8, 6))
        bar.pack(fill="x")
        ttk.Label(bar, text="Chunks / sections JSON:").pack(side="left")
        self.ent_path = ttk.Entry(bar, width=60)
        self.ent_path.pack(side="left", padx=5)
        ttk.Button(bar, text="Browse…", command=self._browse).pack(side="left")
        ttk.Button(bar, text="Load", command=self._load_from_entry).pack(side="left", padx=(4, 12))

        ttk.Label(bar, text="Search:").pack(side="left")
        self.ent_search = ttk.Entry(bar, width=22)
        self.ent_search.pack(side="left", padx=5)
        self.ent_search.bind("<Return>", lambda e: self._populate())
        ttk.Button(bar, text="Find", command=self._populate).pack(side="left")
        ttk.Button(bar, text="Clear", command=self._clear_filter).pack(side="left", padx=(4, 0))

        self.meta_var = tk.StringVar(value="No file loaded.")
        ttk.Label(self.root, textvariable=self.meta_var, foreground="#444444",
                  padding=(10, 0)).pack(fill="x")

        paned = ttk.PanedWindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=8, pady=6)

        left = ttk.Frame(paned)
        paned.add(left, weight=1)

        tools = ttk.Frame(left)
        tools.pack(fill="x")
        ttk.Button(tools, text="Select all", command=self._select_all_checks).pack(side="left")
        ttk.Button(tools, text="Clear checks", command=self._clear_checks).pack(side="left", padx=4)
        self.check_var = tk.StringVar(value="0 checked")
        ttk.Label(tools, textvariable=self.check_var, foreground="#444444").pack(side="right")

        self.tree = ttk.Treeview(left, columns=("sel",), show="tree headings",
                                 selectmode="browse")
        self.tree.heading("#0", text="Section")
        self.tree.heading("sel", text="✓", command=self._toggle_check_all)
        self.tree.column("sel", width=34, minwidth=28, anchor="center", stretch=False)
        yscroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True, pady=(4, 0))
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<space>", self._toggle_check_selected)

        right = ttk.Frame(paned)
        paned.add(right, weight=3)
        self._build_ai_page(right)

        self.status_var = tk.StringVar(value="")
        ttk.Label(self.root, textvariable=self.status_var, relief="sunken",
                  anchor="w", padding=(6, 2)).pack(fill="x", side="bottom")

    # ------------------------------------------------------------- loading -- #
    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select a chunks cache or Processed_chunks JSON",
            filetypes=[("JSON files", "*.json")],
        )
        if path:
            self.ent_path.delete(0, tk.END)
            self.ent_path.insert(0, path)
            self._load(path)

    def _load_from_entry(self):
        path = self.ent_path.get().strip()
        if path:
            self._load(path)

    def _load(self, path):
        if not os.path.exists(path):
            messagebox.showerror("Not found", f"File does not exist:\n{path}")
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, json.JSONDecodeError) as exc:
            messagebox.showerror("Load failed", f"Could not read JSON:\n{exc}")
            return

        sections = sections_from_payload(data)
        if not sections:
            messagebox.showerror(
                "Unexpected format",
                "This file has neither chunks nor 'merged_headings'. Pick a "
                "chunks cache (*_docling_chunks_cache.json) or a "
                "Processed_chunks.json review output.")
            return

        self.json_path = path
        self._sections = sections
        self._checked.clear()
        self.ent_path.delete(0, tk.END)
        self.ent_path.insert(0, path)
        kind = "merged sections" if (isinstance(data, dict) and data.get("merged_headings")) else "chunks"
        self.meta_var.set(f"📄 {Path(path).name}   ·   {len(sections)} {kind}")
        self._populate()

    def _populate(self):
        filt = self.ent_search.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        self.idx_by_iid.clear()
        shown = 0
        for idx, (title, text) in enumerate(self._sections):
            if filt and filt not in title.lower() and filt not in (text or "").lower():
                continue
            glyph = "☑" if idx in self._checked else "☐"
            iid = self.tree.insert("", "end", text=title, values=(glyph,))
            self.idx_by_iid[iid] = idx
            shown += 1
        self._update_check_status()
        if filt:
            self.status_var.set(f"Filter '{filt}': {shown} matching section(s).")
        elif self._sections:
            self.status_var.set(f"Loaded {len(self._sections)} section(s).")

    def _clear_filter(self):
        self.ent_search.delete(0, tk.END)
        self._populate()

    # ---------------------------------------------------------- checkboxes -- #
    def _set_check(self, iid, idx, checked):
        if checked:
            self._checked.add(idx)
        else:
            self._checked.discard(idx)
        self.tree.set(iid, "sel", "☑" if checked else "☐")

    def _on_tree_click(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return None
        if self.tree.identify_column(event.x) != "#1":
            return None
        iid = self.tree.identify_row(event.y)
        idx = self.idx_by_iid.get(iid)
        if idx is None:
            return None
        self._set_check(iid, idx, idx not in self._checked)
        self._update_check_status()
        return "break"

    def _toggle_check_selected(self, event=None):
        for iid in self.tree.selection():
            idx = self.idx_by_iid.get(iid)
            if idx is not None:
                self._set_check(iid, idx, idx not in self._checked)
        self._update_check_status()
        return "break"

    def _select_all_checks(self):
        """Check every visible section (respects an active search filter)."""
        for iid, idx in self.idx_by_iid.items():
            self._set_check(iid, idx, True)
        self._update_check_status()

    def _clear_checks(self):
        self._checked.clear()
        for iid in self.idx_by_iid:
            self.tree.set(iid, "sel", "☐")
        self._update_check_status()

    def _toggle_check_all(self):
        visible = set(self.idx_by_iid.values())
        if visible and visible <= self._checked:
            self._clear_checks()
        else:
            self._select_all_checks()

    def _update_check_status(self):
        self.check_var.set(f"{len(self._checked)} checked")

    # ------------------------------------------------- AIReviewMixin hooks -- #
    def _ai_current_section(self):
        sel = self.tree.selection()
        if not sel:
            return None
        idx = self.idx_by_iid.get(sel[0])
        return self._sections[idx] if idx is not None else None

    def _ai_checked_sections(self):
        return [self._sections[i] for i in sorted(self._checked)]


def main():
    from ..crash_logging import install
    install()
    path = sys.argv[1] if len(sys.argv) > 1 else None
    root = tk.Tk()
    ChunkAIReviewApp(root, json_path=path)
    root.mainloop()


if __name__ == "__main__":
    main()
