"""
Interactive review UI for the EASA structured-extraction JSON.

Opens the JSON produced by EASA_Parser (``{document_metadata, rules_hierarchy}``)
and lets you navigate the regulation hierarchy as an expandable tree, with a
details pane per node: full text, EASA attributes, hyperlinks, and extracted
images/tables (with a native PNG/GIF preview and open-externally support).
Includes a live search that filters the tree to matching branches.

Run standalone:
    python EASA_Json_Review_UI.py [path/to/structured.json]

Also embeddable: pass a container Frame as ``root`` (the Data Extraction Studio
hosts it this way).
"""

import json
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk


def _open_externally(path: str):
    """Open a file with the OS default application (cross-platform)."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: SIM115 - Windows API
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Could not open file", f"{path}\n\n{exc}")


class EASAJsonReviewApp:
    def __init__(self, root, json_path: str = None, logger=None):
        self.root = root
        self.logger = logger
        if isinstance(root, (tk.Tk, tk.Toplevel)):
            root.title("EASA Structured JSON — Review")
            root.geometry("1280x820")
            root.minsize(960, 600)

        self.data = None
        self.json_path = None
        self.assets_base = None       # dir that holds images/ and tables/
        self.node_by_iid = {}
        self._node_count = 0
        self._preview_img = None      # keep a ref so Tk doesn't GC the preview

        self._build_ui()

        if json_path:
            self._load(json_path)

    # ------------------------------------------------------------------ UI -- #
    def _build_ui(self):
        # Top toolbar: file + search
        bar = ttk.Frame(self.root, padding=(8, 6))
        bar.pack(fill="x")

        ttk.Label(bar, text="Structured JSON:").pack(side="left")
        self.ent_path = ttk.Entry(bar, width=60)
        self.ent_path.pack(side="left", padx=5)
        ttk.Button(bar, text="Browse…", command=self._browse).pack(side="left")
        ttk.Button(bar, text="Load", command=self._load_from_entry).pack(side="left", padx=(4, 12))

        ttk.Label(bar, text="Search:").pack(side="left")
        self.ent_search = ttk.Entry(bar, width=26)
        self.ent_search.pack(side="left", padx=5)
        self.ent_search.bind("<Return>", lambda e: self._apply_filter())
        ttk.Button(bar, text="Find", command=self._apply_filter).pack(side="left")
        ttk.Button(bar, text="Clear", command=self._clear_filter).pack(side="left", padx=(4, 0))

        # Document metadata line
        self.meta_var = tk.StringVar(value="No file loaded.")
        ttk.Label(self.root, textvariable=self.meta_var, foreground="#444444",
                  padding=(10, 0)).pack(fill="x")

        # Main split: tree (left) | details (right)
        paned = ttk.PanedWindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=8, pady=6)

        left = ttk.Frame(paned)
        paned.add(left, weight=1)

        tree_tools = ttk.Frame(left)
        tree_tools.pack(fill="x")
        ttk.Button(tree_tools, text="Expand all", command=lambda: self._set_all_open(True)).pack(side="left")
        ttk.Button(tree_tools, text="Collapse all", command=lambda: self._set_all_open(False)).pack(side="left", padx=4)

        self.tree = ttk.Treeview(left, show="tree", selectmode="browse")
        yscroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True, pady=(4, 0))
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        right = ttk.Frame(paned)
        paned.add(right, weight=2)

        self.header_var = tk.StringVar(value="Select a node to see its details.")
        ttk.Label(right, textvariable=self.header_var, font=("TkDefaultFont", 11, "bold"),
                  wraplength=680, justify="left").pack(anchor="w", pady=(0, 4))

        self.detail_nb = ttk.Notebook(right)
        self.detail_nb.pack(fill="both", expand=True)
        self._build_detail_tabs()

        # Status bar
        self.status_var = tk.StringVar(value="")
        ttk.Label(self.root, textvariable=self.status_var, relief="sunken",
                  anchor="w", padding=(6, 2)).pack(fill="x", side="bottom")

    def _build_detail_tabs(self):
        # Overview (ScrolledText must live inside a container frame, not be added
        # to the Notebook directly — its real parent is an internal frame).
        ov = ttk.Frame(self.detail_nb)
        self.tab_overview = scrolledtext.ScrolledText(ov, wrap="word", height=8)
        self.tab_overview.configure(state="disabled")
        self.tab_overview.pack(fill="both", expand=True)
        self.detail_nb.add(ov, text="Overview")

        # Text
        tx = ttk.Frame(self.detail_nb)
        self.tab_text = scrolledtext.ScrolledText(tx, wrap="word")
        self.tab_text.configure(state="disabled")
        self.tab_text.pack(fill="both", expand=True)
        self.detail_nb.add(tx, text="Text")

        # Attributes
        attr_frame = ttk.Frame(self.detail_nb)
        self.tab_attrs = ttk.Treeview(attr_frame, columns=("value",), show="tree headings")
        self.tab_attrs.heading("#0", text="Attribute")
        self.tab_attrs.heading("value", text="Value")
        self.tab_attrs.column("#0", width=220, anchor="w")
        self.tab_attrs.column("value", width=440, anchor="w")
        a_scroll = ttk.Scrollbar(attr_frame, orient="vertical", command=self.tab_attrs.yview)
        self.tab_attrs.configure(yscrollcommand=a_scroll.set)
        a_scroll.pack(side="right", fill="y")
        self.tab_attrs.pack(side="left", fill="both", expand=True)
        self.detail_nb.add(attr_frame, text="Attributes")

        # Hyperlinks
        link_frame = ttk.Frame(self.detail_nb)
        self.tab_links = ttk.Treeview(link_frame, columns=("target",), show="tree headings")
        self.tab_links.heading("#0", text="Link text")
        self.tab_links.heading("target", text="Target")
        self.tab_links.column("#0", width=300, anchor="w")
        self.tab_links.column("target", width=360, anchor="w")
        l_scroll = ttk.Scrollbar(link_frame, orient="vertical", command=self.tab_links.yview)
        self.tab_links.configure(yscrollcommand=l_scroll.set)
        l_scroll.pack(side="right", fill="y")
        self.tab_links.pack(side="left", fill="both", expand=True)
        self.detail_nb.add(link_frame, text="Hyperlinks")

        # Assets (images + tables)
        assets = ttk.Frame(self.detail_nb)
        lists = ttk.Frame(assets)
        lists.pack(side="left", fill="both", expand=True)

        ttk.Label(lists, text="Images:").pack(anchor="w")
        self.lst_images = tk.Listbox(lists, height=8, exportselection=False)
        self.lst_images.pack(fill="both", expand=True)
        self.lst_images.bind("<<ListboxSelect>>", self._on_image_select)
        self.lst_images.bind("<Double-Button-1>", lambda e: self._open_selected_asset(self.lst_images, "images"))

        ttk.Label(lists, text="Tables:").pack(anchor="w", pady=(6, 0))
        self.lst_tables = tk.Listbox(lists, height=6, exportselection=False)
        self.lst_tables.pack(fill="both", expand=True)
        self.lst_tables.bind("<Double-Button-1>", lambda e: self._open_selected_asset(self.lst_tables, "tables"))

        ttk.Label(assets, text="(double-click to open externally)", foreground="#777").pack(
            side="bottom")

        preview = ttk.LabelFrame(assets, text=" Image preview (PNG/GIF) ", padding=6)
        preview.pack(side="right", fill="both", expand=True, padx=(8, 0))
        self.preview_label = ttk.Label(preview, text="Select an image to preview.")
        self.preview_label.pack(fill="both", expand=True)

        self.detail_nb.add(assets, text="Images & Tables")

    # --------------------------------------------------------------- load -- #
    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select EASA structured JSON",
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

        # Accept {document_metadata, rules_hierarchy} or a bare hierarchy list.
        if isinstance(data, dict):
            hierarchy = data.get("rules_hierarchy")
            metadata = data.get("document_metadata", {})
        elif isinstance(data, list):
            hierarchy = data
            metadata = {}
        else:
            hierarchy = None
            metadata = {}

        if not hierarchy:
            messagebox.showerror(
                "Unexpected format",
                "This file has no 'rules_hierarchy'. It does not look like an EASA "
                "structured extraction JSON.",
            )
            return

        self.data = data
        self.json_path = path
        self.assets_base = os.path.dirname(os.path.abspath(path))
        self.ent_path.delete(0, tk.END)
        self.ent_path.insert(0, path)

        title = metadata.get("source-title") or Path(path).stem
        domain = metadata.get("Domain", "")
        self.meta_var.set(f"📄 {title}" + (f"   ·   Domain: {domain}" if domain else ""))

        self._hierarchy = hierarchy
        self._populate_tree()

    # --------------------------------------------------------------- tree -- #
    def _node_label(self, node):
        attrs = node.get("attributes", {}) or {}
        title = attrs.get("source-title") or attrs.get("title")
        etype = node.get("element_type", "node")
        badges = ""
        if node.get("extracted_images"):
            badges += f"  🖼{len(node['extracted_images'])}"
        if node.get("extracted_tables"):
            badges += f"  ▦{len(node['extracted_tables'])}"
        if title:
            return f"{title}  [{etype}]{badges}"
        sid = attrs.get("sdt-id", "")
        return f"{etype} ({sid}){badges}"

    def _subtree_matches(self, node, filt):
        if filt in self._node_label(node).lower():
            return True
        if filt in (node.get("text_content", "") or "").lower():
            return True
        return any(self._subtree_matches(c, filt) for c in (node.get("children") or []))

    def _insert_node(self, parent_iid, node, filt):
        if filt and not self._subtree_matches(node, filt):
            return
        iid = self.tree.insert(parent_iid, "end", text=self._node_label(node))
        self.node_by_iid[iid] = node
        self._node_count += 1
        for child in (node.get("children") or []):
            self._insert_node(iid, child, filt)
        if filt:
            self.tree.item(iid, open=True)

    def _populate_tree(self, filt=None):
        self.tree.delete(*self.tree.get_children())
        self.node_by_iid.clear()
        self._node_count = 0
        for node in self._hierarchy:
            self._insert_node("", node, filt)

        if filt:
            self.status_var.set(f"Filter '{filt}': {self._node_count} matching node(s).")
        else:
            self.status_var.set(f"Loaded {self._node_count} node(s).")
            # Open the first level for orientation.
            for iid in self.tree.get_children():
                self.tree.item(iid, open=True)

    def _apply_filter(self):
        if not getattr(self, "_hierarchy", None):
            return
        filt = self.ent_search.get().strip().lower()
        self._populate_tree(filt or None)

    def _clear_filter(self):
        self.ent_search.delete(0, tk.END)
        if getattr(self, "_hierarchy", None):
            self._populate_tree(None)

    def _set_all_open(self, is_open):
        def walk(iid):
            self.tree.item(iid, open=is_open)
            for c in self.tree.get_children(iid):
                walk(c)
        for iid in self.tree.get_children():
            walk(iid)

    # ------------------------------------------------------------ details -- #
    def _set_text(self, widget, content):
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert(tk.END, content or "")
        widget.configure(state="disabled")

    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        node = self.node_by_iid.get(sel[0])
        if node is None:
            return
        self._show_node(node)

    def _show_node(self, node):
        attrs = node.get("attributes", {}) or {}
        title = attrs.get("source-title") or attrs.get("title") or node.get("element_type", "node")
        self.header_var.set(f"{title}")

        text_content = node.get("text_content", "") or ""
        links = node.get("hyperlinks", []) or []
        images = node.get("extracted_images", []) or []
        tables = node.get("extracted_tables", []) or []

        # Overview
        overview = (
            f"Element type : {node.get('element_type', '')}\n"
            f"sdt-id       : {attrs.get('sdt-id', '')}\n"
            f"Title        : {title}\n"
            f"Children     : {len(node.get('children') or [])}\n"
            f"Hyperlinks   : {len(links)}\n"
            f"Images       : {len(images)}\n"
            f"Tables       : {len(tables)}\n"
            f"Text length  : {len(text_content)} chars\n"
        )
        self._set_text(self.tab_overview, overview)

        # Text
        self._set_text(self.tab_text, text_content)

        # Attributes
        self.tab_attrs.delete(*self.tab_attrs.get_children())
        for k, v in attrs.items():
            self.tab_attrs.insert("", "end", text=str(k), values=(str(v),))

        # Hyperlinks
        self.tab_links.delete(*self.tab_links.get_children())
        for link in links:
            self.tab_links.insert("", "end", text=str(link.get("text", "")),
                                  values=(str(link.get("target", "")),))

        # Assets
        self.lst_images.delete(0, tk.END)
        for img in images:
            self.lst_images.insert(tk.END, img)
        self.lst_tables.delete(0, tk.END)
        for tbl in tables:
            self.lst_tables.insert(tk.END, tbl)
        self.preview_label.configure(image="", text="Select an image to preview.")
        self._preview_img = None

    def _resolve_asset(self, filename, kind):
        """Resolve an image/table filename to an on-disk path next to the JSON."""
        if not self.assets_base:
            return None
        candidate = os.path.join(self.assets_base, kind, filename)
        if os.path.exists(candidate):
            return candidate
        # Fall back to the filename as-is (in case it is already a full path).
        return filename if os.path.exists(filename) else None

    def _on_image_select(self, event=None):
        sel = self.lst_images.curselection()
        if not sel:
            return
        filename = self.lst_images.get(sel[0])
        path = self._resolve_asset(filename, "images")
        if not path:
            self.preview_label.configure(image="", text=f"(file not found)\n{filename}")
            self._preview_img = None
            return
        # Tk's PhotoImage supports PNG/GIF natively; other formats fall back.
        try:
            self._preview_img = tk.PhotoImage(file=path)
            # Downscale very large images so they fit the pane.
            w = self._preview_img.width()
            if w > 520:
                factor = max(1, w // 520)
                self._preview_img = self._preview_img.subsample(factor, factor)
            self.preview_label.configure(image=self._preview_img, text="")
        except Exception:
            self._preview_img = None
            self.preview_label.configure(
                image="", text=f"No inline preview for this format.\nDouble-click to open:\n{filename}"
            )

    def _open_selected_asset(self, listbox, kind):
        sel = listbox.curselection()
        if not sel:
            return
        filename = listbox.get(sel[0])
        path = self._resolve_asset(filename, kind)
        if path:
            _open_externally(path)
        else:
            messagebox.showinfo("Not found", f"Could not locate the file on disk:\n{filename}")


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else None
    root = tk.Tk()
    EASAJsonReviewApp(root, json_path=path)
    root.mainloop()


if __name__ == "__main__":
    main()
