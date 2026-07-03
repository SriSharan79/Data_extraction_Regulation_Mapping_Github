import json
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import logging

class SectionReviewApp:
    """
    Section review interface that loads merged_headings from chunk review output.
    User selects a section from a dropdown, can edit it, and save changes back.
    Includes cache launcher integration for seamless workflow.
    """
    
    def __init__(self, root, output_file_path, logger=None):
        self.root = root
        self.output_file_path = output_file_path
        self.logger = logger or logging.getLogger("SectionReview")
        
        self.sections_data = []
        self.current_section_index = -1
        self.unsaved_changes = False
        
        # Configure Main Window
        self.root.title("Docling Section Review & Refinement Tool")
        self.root.geometry("1000x750")
        self.root.minsize(900, 650)
        
        # UI Styling
        self.bg_color = "#f4f4f6"
        self.card_color = "#ffffff"
        self.accent_color = "#3498db"
        self.root.configure(bg=self.bg_color)
        
        # Load data first
        self.load_sections_from_output()
        
        # Build UI
        self.setup_ui()
        self.load_section_selector()

    def load_sections_from_output(self):
        """Load merged_headings from chunk review output file."""
        if not self.output_file_path:
            self.logger.warning("No output file path provided for section review.")
            return
        
        try:
            with open(self.output_file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            self.sections_data = data.get("merged_headings", [])
            self.logger.info(f"Loaded {len(self.sections_data)} sections from chunk review output.")
            
            if not self.sections_data:
                messagebox.showwarning("No Sections Found", "No merged sections were found in the chunk review output.\nPlease complete chunk review first.")
                self.root.destroy()
                return
                
        except FileNotFoundError:
            messagebox.showerror("File Not Found", f"Output file not found: {self.output_file_path}")
            self.root.destroy()
        except json.JSONDecodeError as e:
            messagebox.showerror("JSON Error", f"Failed to parse output file: {e}")
            self.root.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sections: {e}")
            self.root.destroy()

    def setup_ui(self):
        """Build the section review interface."""
        
        # ============ TOP HEADER FRAME ============
        header_frame = tk.LabelFrame(self.root, text=" Section Selection & Navigation ", 
                                    bg=self.card_color, font=("Arial", 11, "bold"), 
                                    padx=15, pady=10)
        header_frame.pack(fill="x", padx=15, pady=10)
        
        # Section selector label and dropdown
        tk.Label(header_frame, text="Select Section to Review:", font=("Arial", 10, "bold"), 
                bg=self.card_color).grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        self.section_selector = ttk.Combobox(header_frame, state="readonly", font=("Arial", 10), width=80)
        self.section_selector.grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(0, 10))
        self.section_selector.bind("<<ComboboxSelected>>", self.on_section_selected)
        
        # Navigation buttons
        nav_frame = tk.Frame(header_frame, bg=self.card_color)
        nav_frame.grid(row=1, column=1, sticky="ew", padx=5)
        
        self.btn_prev = tk.Button(nav_frame, text="← Prev", font=("Arial", 9, "bold"), 
                                 bg="#95a5a6", fg="white", width=8, command=self.go_prev_section)
        self.btn_prev.pack(side="left", padx=2)
        
        self.btn_next = tk.Button(nav_frame, text="Next →", font=("Arial", 9, "bold"), 
                                 bg="#95a5a6", fg="white", width=8, command=self.go_next_section)
        self.btn_next.pack(side="left", padx=2)
        
        header_frame.columnconfigure(0, weight=1)
        
        # ============ SECTION METADATA FRAME ============
        meta_frame = tk.LabelFrame(self.root, text=" Section Metadata ", 
                                  bg=self.card_color, font=("Arial", 10, "bold"), 
                                  padx=15, pady=10)
        meta_frame.pack(fill="x", padx=15, pady=5)
        
        # Section heading display
        tk.Label(meta_frame, text="Section Heading:", font=("Arial", 9, "bold"), bg=self.card_color).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.meta_heading = tk.Label(meta_frame, text="N/A", font=("Arial", 10), bg="#ecf0f1", fg="#2c3e50", padx=10, pady=5)
        self.meta_heading.grid(row=0, column=1, sticky="ew", pady=2)
        
        # Chunk indices info
        tk.Label(meta_frame, text="Source Chunks:", font=("Arial", 9, "bold"), bg=self.card_color).grid(row=1, column=0, sticky="w", padx=(0, 10))
        self.meta_chunks = tk.Label(meta_frame, text="N/A", font=("Arial", 9), bg="#ecf0f1", fg="#34495e", padx=10, pady=5)
        self.meta_chunks.grid(row=1, column=1, sticky="ew", pady=2)
        
        # Document item types info
        tk.Label(meta_frame, text="Content Types:", font=("Arial", 9, "bold"), bg=self.card_color).grid(row=2, column=0, sticky="w", padx=(0, 10))
        self.meta_types = tk.Label(meta_frame, text="N/A", font=("Arial", 9), bg="#ecf0f1", fg="#34495e", padx=10, pady=5)
        self.meta_types.grid(row=2, column=1, sticky="ew", pady=2)
        
        # Page numbers info
        tk.Label(meta_frame, text="Page Numbers:", font=("Arial", 9, "bold"), bg=self.card_color).grid(row=3, column=0, sticky="w", padx=(0, 10))
        self.meta_pages = tk.Label(meta_frame, text="N/A", font=("Arial", 9), bg="#ecf0f1", fg="#34495e", padx=10, pady=5)
        self.meta_pages.grid(row=3, column=1, sticky="ew", pady=2)
        
        meta_frame.columnconfigure(1, weight=1)
        
        # ============ HEADING EDITOR FRAME ============
        heading_edit_frame = tk.Frame(self.root, bg=self.bg_color)
        heading_edit_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        tk.Label(heading_edit_frame, text="Edit Section Heading:", font=("Arial", 10, "bold"), bg=self.bg_color).pack(anchor="w")
        self.heading_entry = tk.Entry(heading_edit_frame, font=("Arial", 10), bd=2, relief="groove")
        self.heading_entry.pack(fill="x", pady=(2, 8))
        self.heading_entry.bind("<KeyRelease>", lambda e: self.mark_unsaved())
        
        # ============ CONTENT EDITOR FRAME ============
        content_frame = tk.LabelFrame(self.root, text=" Merged Section Text Content (Editable) ", 
                                     bg=self.card_color, font=("Arial", 10, "bold"), 
                                     padx=10, pady=10)
        content_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        self.text_editor = scrolledtext.ScrolledText(content_frame, font=("Consolas", 10), 
                                                    wrap=tk.WORD, bd=2, relief="groove", 
                                                    bg="#fafafa", fg="#2c3e50")
        self.text_editor.pack(fill="both", expand=True)
        self.text_editor.bind("<KeyRelease>", lambda e: self.mark_unsaved())
        
        # ============ ACTION BUTTON FRAME ============
        btn_frame = tk.Frame(self.root, bg=self.bg_color, pady=10)
        btn_frame.pack(fill="x", padx=15)
        
        # Left side buttons
        left_btn_frame = tk.Frame(btn_frame, bg=self.bg_color)
        left_btn_frame.pack(side="left", anchor="w")
        
        self.btn_add_section = tk.Button(left_btn_frame, text="＋ Add New Section", 
                                        font=("Arial", 10, "bold"), bg="#27ae60", fg="white", 
                                        width=18, command=self.add_new_section, relief="raised", bd=2)
        self.btn_add_section.pack(side="left", padx=5)
        
        self.btn_delete_section = tk.Button(left_btn_frame, text="❌ Delete Section", 
                                           font=("Arial", 10, "bold"), bg="#e74c3c", fg="white", 
                                           width=16, command=self.delete_current_section, relief="raised", bd=2)
        self.btn_delete_section.pack(side="left", padx=5)
        
        # Right side buttons
        right_btn_frame = tk.Frame(btn_frame, bg=self.bg_color)
        right_btn_frame.pack(side="right", anchor="e")
        
        self.btn_save = tk.Button(right_btn_frame, text="💾 Save Changes", 
                                 font=("Arial", 10, "bold"), bg="#f39c12", fg="white", 
                                 width=16, command=self.save_changes, relief="raised", bd=2)
        self.btn_save.pack(side="left", padx=5)
        
        self.btn_export = tk.Button(right_btn_frame, text="📤 Export & Finish", 
                                   font=("Arial", 11, "bold"), bg="#27ae60", fg="white", 
                                   width=18, command=self.export_and_finish, relief="raised", bd=3)
        self.btn_export.pack(side="left", padx=5)
        
        self.btn_close = tk.Button(right_btn_frame, text="Exit", 
                                  font=("Arial", 9), bg="#95a5a6", fg="white", 
                                  width=8, command=self.on_close)
        self.btn_close.pack(side="left", padx=5)
        
        # ============ STATUS BAR ============
        self.status_label = tk.Label(self.root, text="Ready", font=("Arial", 9, "italic"), 
                                    bg=self.bg_color, fg="#7f8c8d", anchor="w")
        self.status_label.pack(fill="x", padx=15, pady=(5, 10))

    def load_section_selector(self):
        """Populate the section dropdown with all available sections."""
        if not self.sections_data:
            return
        
        section_names = []
        for idx, section in enumerate(self.sections_data):
            heading = section.get("heading", [])
            if isinstance(heading, list):
                heading_text = " → ".join(heading) if heading else f"Section {idx + 1}"
            else:
                heading_text = str(heading) if heading else f"Section {idx + 1}"
            
            section_names.append(f"[{idx + 1}] {heading_text}")
        
        self.section_selector['values'] = section_names
        
        # Select first section by default
        if section_names:
            self.section_selector.current(0)
            self.on_section_selected(None)

    def on_section_selected(self, event):
        """Handle section selection from dropdown."""
        idx = self.section_selector.current()
        if idx < 0 or idx >= len(self.sections_data):
            return
        
        # Save current section if there are unsaved changes
        if self.unsaved_changes and self.current_section_index >= 0:
            if messagebox.askyesno("Unsaved Changes", "Save changes to the current section before switching?"):
                self.save_current_section_data()
        
        self.current_section_index = idx
        self.display_current_section()
        self.unsaved_changes = False

    def display_current_section(self):
        """Display the currently selected section's data."""
        if self.current_section_index < 0 or self.current_section_index >= len(self.sections_data):
            return
        
        section = self.sections_data[self.current_section_index]
        
        # Extract heading
        heading = section.get("heading", [])
        if isinstance(heading, list):
            heading_text = " → ".join(heading) if heading else "Untitled"
        else:
            heading_text = str(heading) if heading else "Untitled"
        
        # Update heading entry
        self.heading_entry.delete(0, tk.END)
        self.heading_entry.insert(0, heading_text)
        
        # Update text content
        merged_text = section.get("merged_text", "")
        self.text_editor.delete("1.0", tk.END)
        self.text_editor.insert(tk.END, merged_text)
        
        # Update metadata
        chunk_indices = section.get("chunk_indices", [])
        self.meta_chunks.config(text=", ".join(map(str, chunk_indices)) if chunk_indices else "N/A")
        
        doc_types = section.get("types_of_docitem", [])
        self.meta_types.config(text=", ".join(doc_types) if doc_types else "N/A")
        
        page_nums = section.get("page_nums", [])
        self.meta_pages.config(text=", ".join(map(str, page_nums)) if page_nums else "N/A")
        
        self.meta_heading.config(text=heading_text)
        
        self.update_status(f"Displaying section {self.current_section_index + 1} of {len(self.sections_data)}")
        self.logger.info(f"Loaded section {self.current_section_index + 1}: {heading_text}")

    def save_current_section_data(self):
        """Sync current UI values back to section data."""
        if self.current_section_index < 0 or self.current_section_index >= len(self.sections_data):
            return
        
        section = self.sections_data[self.current_section_index]
        
        # Update heading
        new_heading = self.heading_entry.get().strip()
        section["heading"] = new_heading.split(" → ") if " → " in new_heading else [new_heading]
        
        # Update text content
        new_text = self.text_editor.get("1.0", tk.END).rstrip("\n")
        section["merged_text"] = new_text

    def save_changes(self):
        """Save current section changes to memory."""
        self.save_current_section_data()
        self.unsaved_changes = False
        self.update_status("✓ Changes saved to memory")
        self.logger.info(f"Saved changes for section {self.current_section_index + 1}")
        messagebox.showinfo("Success", "Section changes saved to memory.\n\nClick 'Export & Finish' to save to file.")

    def export_and_finish(self):
        """Export all sections back to the output file and finish."""
        # Save any pending changes
        if self.unsaved_changes:
            self.save_current_section_data()
        
        # Update the output file with modified sections
        try:
            with open(self.output_file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Update merged_headings with modified sections
            data["merged_headings"] = self.sections_data
            
            with open(self.output_file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            
            self.logger.info("Successfully exported all section reviews to output file.")
            messagebox.showinfo("Export Complete", 
                              f"✓ Section review completed!\n\n{len(self.sections_data)} sections saved.\n\nOutput file updated.")
            self.root.destroy()
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export sections: {e}")
            self.logger.error(f"Export failed: {e}")

    def add_new_section(self):
        """Add a new empty section."""
        new_section = {
            "heading": ["New Section"],
            "merged_text": "",
            "chunk_indices": [],
            "types_of_docitem": [],
            "page_nums": []
        }
        self.sections_data.append(new_section)
        self.load_section_selector()
        self.section_selector.current(len(self.sections_data) - 1)
        self.on_section_selected(None)
        self.update_status(f"Added new section. Total sections: {len(self.sections_data)}")
        self.logger.info("Added new empty section")

    def delete_current_section(self):
        """Delete the currently selected section."""
        if self.current_section_index < 0:
            messagebox.showwarning("No Selection", "Please select a section to delete.")
            return
        
        if not messagebox.askyesno("Confirm Deletion", 
                                  f"Are you sure you want to delete this section?\n\n'{self.heading_entry.get()}'"):
            return
        
        section_heading = self.heading_entry.get()
        self.sections_data.pop(self.current_section_index)
        
        self.load_section_selector()
        next_idx = max(0, self.current_section_index - 1)
        if self.sections_data and next_idx < len(self.sections_data):
            self.section_selector.current(next_idx)
            self.on_section_selected(None)
        else:
            self.current_section_index = -1
            self.heading_entry.delete(0, tk.END)
            self.text_editor.delete("1.0", tk.END)
        
        self.update_status(f"Deleted section: {section_heading}")
        self.logger.info(f"Deleted section: {section_heading}")

    def go_prev_section(self):
        """Navigate to previous section."""
        if self.current_section_index > 0:
            self.section_selector.current(self.current_section_index - 1)
            self.on_section_selected(None)

    def go_next_section(self):
        """Navigate to next section."""
        if self.current_section_index < len(self.sections_data) - 1:
            self.section_selector.current(self.current_section_index + 1)
            self.on_section_selected(None)

    def mark_unsaved(self):
        """Mark that there are unsaved changes."""
        self.unsaved_changes = True
        self.update_status("● Unsaved changes...")

    def update_status(self, message):
        """Update the status bar."""
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def on_close(self):
        """Handle window close with unsaved changes check."""
        if self.unsaved_changes:
            if messagebox.askyesno("Unsaved Changes", "You have unsaved changes. Save before exiting?"):
                self.export_and_finish()
            else:
                self.root.destroy()
        else:
            self.root.destroy()