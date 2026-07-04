import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from colorama import Fore

class SectionRewriterUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Docling Cache - Raw Section Re-Writer Tool")
        self.root.geometry("1300x850")
        
        # Data Repositories
        self.loaded_chunks = []
        self.sections_data = []
        self.chunks_file_path = ""
        self.sections_file_path = ""

        self._build_style_and_layout()

    def _build_style_and_layout(self):
        # Apply clean dark-accented style settings
        style = ttk.Style()
        style.theme_use('clam')
        
        # Main File IO Control Header
        top_frame = ttk.LabelFrame(self.root, text=" File Path Management Configuration ", padding=10)
        top_frame.pack(fill="x", padx=10, pady=5)
        
        # Raw Chunk JSON File Path Row
        ttk.Label(top_frame, text="Raw Chunks JSON:").grid(row=0, column=0, sticky="w", pady=2)
        self.ent_chunks_path = ttk.Entry(top_frame, width=90)
        self.ent_chunks_path.grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(top_frame, text="Browse...", command=self.browse_chunks_file).grid(row=0, column=2, padx=2, pady=2)
        
        # Raw Section JSON File Path Row
        ttk.Label(top_frame, text="Raw Sections JSON:").grid(row=1, column=0, sticky="w", pady=2)
        self.ent_sections_path = ttk.Entry(top_frame, width=90)
        self.ent_sections_path.grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(top_frame, text="Browse...", command=self.browse_sections_file).grid(row=1, column=2, padx=2, pady=2)

        ttk.Button(top_frame, text="LOAD DATA SOURCE", command=self.load_files_data).grid(row=0, column=3, rowspan=2, padx=15, sticky="ns")

        # Workspace Splitter (Left: Chunks Reference, Right: Section Target Builder)
        paned_window = ttk.PanedWindow(self.root, orient="horizontal")
        paned_window.pack(fill="both", expand=True, padx=10, pady=5)

        # LEFT WORKSPACE PANEL
        left_frame = ttk.LabelFrame(paned_window, text=" Cached Structural Input Chunks Reference Source ", padding=5)
        paned_window.add(left_frame, weight=1)
        
        # Interactive Treeview list of cache chunks
        self.chunk_tree = ttk.Treeview(left_frame, columns=("page", "heading"), show="headings", selectmode="browse")
        self.chunk_tree.heading("page", text="Page(s)")
        self.chunk_tree.heading("heading", text="Detected Origin Heading Block")
        self.chunk_tree.column("page", width=70, anchor="center")
        self.chunk_tree.column("heading", width=400, anchor="w")
        
        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.chunk_tree.yview)
        self.chunk_tree.configure(yscrollcommand=tree_scroll.set)
        
        self.chunk_tree.pack(fill="both", expand=True, side="top")
        tree_scroll.pack(side="right", fill="y", before=self.chunk_tree)
        self.chunk_tree.bind("<<TreeviewSelect>>", self.on_chunk_selected)

        # Embedded Chunk Text Inspector Panel
        lbl_inspect = ttk.Label(left_frame, text="Selected Individual Chunk Content Text View Window:", font=("TkDefaultFont", 9, "bold"))
        lbl_inspect.pack(anchor="w", pady=(5,0))
        self.txt_chunk_inspector = scrolledtext.ScrolledText(left_frame, height=12, bg="#fcfcfc", wrap="word")
        self.txt_chunk_inspector.pack(fill="x", side="bottom")

        # RIGHT WORKSPACE PANEL
        right_frame = ttk.LabelFrame(paned_window, text=" Target Output Sections Re-Writer Workspace Canvas ", padding=5)
        paned_window.add(right_frame, weight=1)

        # Dynamic Section List Navigator
        sec_nav_frame = ttk.Frame(right_frame)
        sec_nav_frame.pack(fill="x", pady=2)
        ttk.Label(sec_nav_frame, text="Active Editor Target Section:").pack(side="left", padx=2)
        self.cbo_sections = ttk.Combobox(sec_nav_frame, state="readonly", width=45)
        self.cbo_sections.pack(side="left", padx=5, fill="x", expand=True)
        self.cbo_sections.bind("<<ComboboxSelected>>", self.on_section_combo_changed)

        # Section Heading Modification Input Box
        heading_edit_frame = ttk.Frame(right_frame)
        heading_edit_frame.pack(fill="x", pady=4)
        ttk.Label(heading_edit_frame, text="Modify Title / Heading:").pack(side="left", padx=2)
        self.ent_section_heading = ttk.Entry(heading_edit_frame)
        self.ent_section_heading.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(heading_edit_frame, text="Update Header", command=self.update_current_section_header).pack(side="right", padx=2)

        # Main Text Editor Area for target block contents
        ttk.Label(right_frame, text="Edit Drafted Text Body Content for current Section Workspace:").pack(anchor="w", pady=(5,0))
        self.txt_section_editor = scrolledtext.ScrolledText(right_frame, wrap="word", bg="#f9fbf9", font=("Courier New", 10))
        self.txt_section_editor.pack(fill="both", expand=True, pady=4)

        # Action Execution Button Tray Footer
        button_tray = ttk.Frame(self.root, padding=5)
        button_tray.pack(fill="x", side="bottom")

        ttk.Button(button_tray, text="＋ Append New Empty Section Block", command=self.append_new_section_block).pack(side="left", padx=5)
        ttk.Button(button_tray, text="❌ Delete Current Active Section", command=self.delete_current_section).pack(side="left", padx=5)
        ttk.Button(button_tray, text="⚡ Run Baseline Auto-Reconstruction Parsing Logic", command=self.run_baseline_auto_reconstruction).pack(side="left", padx=20)
        
        btn_save = ttk.Button(button_tray, text="💾 SAVE AND EXPORT RAW SECTION JSON FILE", command=self.save_and_export_sections)
        btn_save.pack(side="right", padx=10)

    # File dialog handlers
    def browse_chunks_file(self):
        path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if path:
            self.ent_chunks_path.delete(0, tk.END)
            self.ent_chunks_path.insert(0, path)

    def browse_sections_file(self):
        path = filedialog.asksaveasfilename(filetypes=[("JSON Files", "*.json")])
        if path:
            self.ent_sections_path.delete(0, tk.END)
            self.ent_sections_path.insert(0, path)

    def load_files_data(self):
        self.chunks_file_path = self.ent_chunks_path.get().strip()
        self.sections_file_path = self.ent_sections_path.get().strip()

        if not os.path.exists(self.chunks_file_path):
            messagebox.showerror("IO Missing Error", "Please provide a valid source data file location path for Raw Chunks JSON.")
            return

        try:
            with open(self.chunks_file_path, 'r', encoding='utf-8') as f:
                self.loaded_chunks = json.load(f)
            
            # Clear old layout tree
            for idx in self.chunk_tree.get_children():
                self.chunk_tree.delete(idx)

            # Repopulate interactive reference view
            for index, chunk in enumerate(self.loaded_chunks):
                meta = chunk.get("meta", {})
                pages = meta.get("page_numbers", [])
                page_str = ", ".join(map(str, pages)) if pages else "N/A"
                headings_list = meta.get("headings", [])
                heading_str = headings_list[0] if headings_list else "General / No Heading Detected"
                
                self.chunk_tree.insert("", "end", iid=str(index), values=(page_str, heading_str))

            # Attempt to pull Section Output data if pre-existing
            if os.path.exists(self.sections_file_path):
                with open(self.sections_file_path, 'r', encoding='utf-8') as f:
                    self.sections_data = json.load(f)
                self.refresh_sections_dropdown()
            else:
                self.sections_data = []
                self.refresh_sections_dropdown()

            messagebox.showinfo("Success", f"Successfully loaded {len(self.loaded_chunks)} source document chunks into workspace references.")

        except Exception as e:
            messagebox.showerror("Parsing Error", f"Failed to ingest files: {str(e)}")

    def on_chunk_selected(self, event):
        selected_item = self.chunk_tree.selection()
        if not selected_item:
            return
        chunk_idx = int(selected_item[0])
        chunk = self.loaded_chunks[chunk_idx]
        
        # Display text representation safely
        self.txt_chunk_inspector.delete("1.0", tk.END)
        self.txt_chunk_inspector.insert(tk.END, chunk.get("text", ""))

    def run_baseline_auto_reconstruction(self):
        """
        Emulates standard data sorting fallback extraction logic 
        over raw cached items to quickly bootstrap the layout sections tree.
        """
        if not self.loaded_chunks:
            messagebox.showwarning("Empty Source", "Please load a raw chunk cache file before running baseline algorithms.")
            return

        # Re-build section layout maps using standard heuristics from cache keys
        reconstructed = {}
        for chunk in self.loaded_chunks:
            headings_list = chunk.get("meta", {}).get("headings", [])
            heading_key = headings_list[0].strip() if headings_list else "General"
            text_val = chunk.get("text", "").strip()

            if not text_val:
                continue
            
            if heading_key not in reconstructed:
                reconstructed[heading_key] = []
            reconstructed[heading_key].append(text_val)

        # Merge structural arrays into final document section objects
        self.sections_data = []
        for header, text_list in reconstructed.items():
            section_block = {
                "Section Heading": header,
                "Text_Content": "\n\n".join(text_list),
                "Chunks": [(i+1, text) for i, text in enumerate(text_list)]
            }
            self.sections_data.append(section_block)

        self.refresh_sections_dropdown()
        messagebox.showinfo("Auto Process", "Successfully reconstructed baseline data map from JSON structure keys.")

    def refresh_sections_dropdown(self, select_index=0):
        if not self.sections_data:
            self.cbo_sections['values'] = ()
            self.ent_section_heading.delete(0, tk.END)
            self.txt_section_editor.delete("1.0", tk.END)
            return

        titles = [f"[{i}] {sec.get('Section Heading', 'Untitled')}" for i, sec in enumerate(self.sections_data)]
        self.cbo_sections['values'] = titles
        
        if select_index < len(self.sections_data):
            self.cbo_sections.current(select_index)
            self.on_section_combo_changed(None)

    def on_section_combo_changed(self, event):
        idx = self.cbo_sections.current()
        if idx == -1:
            return
        
        # Pull active selected working draft block
        section = self.sections_data[idx]
        
        # Synchronize interface fields
        self.ent_section_heading.delete(0, tk.END)
        self.ent_section_heading.insert(0, section.get("Section Heading", ""))
        
        self.txt_section_editor.delete("1.0", tk.END)
        self.txt_section_editor.insert(tk.END, section.get("Text_Content", ""))

    def update_current_section_header(self):
        idx = self.cbo_sections.current()
        if idx == -1:
            return
        
        # Save working changes to master active dictionary
        new_header = self.ent_section_heading.get().strip()
        self.sections_data[idx]["Section Heading"] = new_header
        
        # Commit current editor body changes to memory layout too
        self.sections_data[idx]["Text_Content"] = self.txt_section_editor.get("1.0", tk.END).strip()
        self.refresh_sections_dropdown(select_index=idx)

    def append_new_section_block(self):
        new_block = {
            "Section Heading": "New Draft Section Header Block",
            "Text_Content": "",
            "Chunks": []
        }
        self.sections_data.append(new_block)
        self.refresh_sections_dropdown(select_index=len(self.sections_data) - 1)

    def delete_current_section(self):
        idx = self.cbo_sections.current()
        if idx == -1:
            return
        
        if messagebox.askyesno("Confirm Deletion", "Are you sure you want to remove this active section from the output layout canvas?"):
            self.sections_data.pop(idx)
            next_select = max(0, idx - 1)
            self.refresh_sections_dropdown(select_index=next_select if self.sections_data else 0)

    def save_and_export_sections(self):
        # Sync current workspace view parameters to internal model structures first
        idx = self.cbo_sections.current()
        if idx != -1:
            self.sections_data[idx]["Section Heading"] = self.ent_section_heading.get().strip()
            self.sections_data[idx]["Text_Content"] = self.txt_section_editor.get("1.0", tk.END).strip()

        target_out_path = self.ent_sections_path.get().strip()
        if not target_out_path:
            target_out_path = filedialog.asksaveasfilename(filetypes=[("JSON Files", "*.json")])
            if not target_out_path:
                return
            self.ent_sections_path.delete(0, tk.END)
            self.ent_sections_path.insert(0, target_out_path)

        try:
            with open(target_out_path, 'w', encoding='utf-8') as out_f:
                json.dump(self.sections_data, out_f, indent=4, ensure_ascii=False)
            messagebox.showinfo("Export Complete", f"✓ Processed data successfully written to target:\n{target_out_path}")
        except Exception as e:
            messagebox.showerror("Write Failure", f"Failed saving output file down to disk array: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SectionRewriterUI(root)
    root.mainloop()