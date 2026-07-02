import json
import tkinter as tk
from tkinter import messagebox, scrolledtext

class ChunkReviewApp:
    def __init__(self, root, chunks_data, logged_chunks, processed_indices, output_file_name, logger):
        self.root = root
        self.chunks_data = chunks_data
        self.logged_chunks = logged_chunks
        self.processed_indices = processed_indices
        self.output_file_name = output_file_name
        self.logger = logger
        
        self.current_pointer = 0
        
        # Configure Main Window
        self.root.title("Docling Chunk Review & Curation Tool")
        self.root.geometry("850x700")
        self.root.minsize(750, 600)
        
        # UI Styling Elements
        self.bg_color = "#f4f4f6"
        self.card_color = "#ffffff"
        self.root.configure(bg=self.bg_color)
        
        self.setup_ui()
        self.advance_to_next_unprocessed()
        self.display_current_chunk()

    def setup_ui(self):
        # Top Metadata Frame
        meta_frame = tk.LabelFrame(self.root, text=" Chunk Metadata ", bg=self.card_color, font=("Arial", 10, "bold"), padx=15, pady=10)
        meta_frame.pack(fill="x", padx=15, pady=10)
        
        # Progress Tracking Label
        self.progress_label = tk.Label(meta_frame, text="Chunk: 0 / 0", font=("Arial", 11, "bold"), bg=self.card_color, fg="#2c3e50")
        self.progress_label.grid(row=0, column=0, sticky="w", pady=2)
        
        self.type_label = tk.Label(meta_frame, text="Chunk Type: N/A", font=("Arial", 10), bg=self.card_color)
        self.type_label.grid(row=1, column=0, sticky="w", pady=2)
        
        self.pages_label = tk.Label(meta_frame, text="Page Number(s): N/A", font=("Arial", 10), bg=self.card_color)
        self.pages_label.grid(row=2, column=0, sticky="w", pady=2)
        
        self.docitem_label = tk.Label(meta_frame, text="DocItem Elements: N/A", font=("Arial", 10), bg=self.card_color)
        self.docitem_label.grid(row=3, column=0, sticky="w", pady=2)

        # Headings Editing Frame
        heading_frame = tk.Frame(self.root, bg=self.bg_color)
        heading_frame.pack(fill="x", padx=15, pady=5)
        
        tk.Label(heading_frame, text="Headings (Comma-separated list):", font=("Arial", 10, "bold"), bg=self.bg_color).pack(anchor="w")
        self.headings_entry = tk.Entry(heading_frame, font=("Arial", 10), bd=2, relief="groove")
        self.headings_entry.pack(fill="x", pady=2)

        # Text Content Editing Frame
        content_frame = tk.Frame(self.root, bg=self.bg_color)
        content_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        tk.Label(content_frame, text="Chunk Text Content (Editable):", font=("Arial", 10, "bold"), bg=self.bg_color).pack(anchor="w")
        self.text_area = scrolledtext.ScrolledText(content_frame, font=("Consolas", 10), wrap=tk.WORD, bd=2, relief="groove")
        self.text_area.pack(fill="both", expand=True, pady=2)

        # Bottom Action Control Frame
        btn_frame = tk.Frame(self.root, bg=self.bg_color, pady=10)
        btn_frame.pack(fill="x", padx=15)
        
        self.btn_log = tk.Button(btn_frame, text="Log Chunk (Save)", font=("Arial", 10, "bold"), bg="#27ae60", fg="white", width=16, command=self.log_chunk, relief="raised", bd=3)
        self.btn_log.pack(side="left", padx=5)
        
        self.btn_skip = tk.Button(btn_frame, text="Skip Chunk", font=("Arial", 10, "bold"), bg="#e67e22", fg="white", width=14, command=self.skip_chunk, relief="raised", bd=3)
        self.btn_skip.pack(side="left", padx=5)

        self.btn_use_prev = tk.Button(btn_frame, text="Use Prev Heading", font=("Arial", 10, "bold"), bg="#34495e", fg="white", width=16, command=self.use_previous_heading, relief="raised", bd=3)
        self.btn_use_prev.pack(side="left", padx=5)

        self.btn_reset = tk.Button(btn_frame, text="Reset & Restart", font=("Arial", 10, "bold"), bg="#c0392b", fg="white", width=16, command=self.reset_all_progress, relief="raised", bd=3)
        self.btn_reset.pack(side="left", padx=5)
        
        self.btn_close = tk.Button(btn_frame, text="Exit Session", font=("Arial", 10), bg="#95a5a6", fg="white", width=12, command=self.root.destroy)
        self.btn_close.pack(side="right", padx=5)

    def advance_to_next_unprocessed(self):
        while self.current_pointer < len(self.chunks_data):
            chunk_idx = self.chunks_data[self.current_pointer]["chunk_index"]
            if chunk_idx in self.processed_indices:
                self.logger.info(f"Chunk {chunk_idx} bypassed automatically via Resume Checkpoint.")
                self.current_pointer += 1
            else:
                break

    def display_current_chunk(self):
        if self.current_pointer >= len(self.chunks_data):
            self.logger.info("All document chunks completed.")
            messagebox.showinfo("Curation Complete", "All chunks have been successfully reviewed and processed!")
            self.root.destroy()
            return
        
        chunk = self.chunks_data[self.current_pointer]
        
        self.progress_label.config(text=f"Chunk Summary: {chunk['chunk_index']} / {len(self.chunks_data)}")
        self.type_label.config(text=f"Chunk Type: {chunk['type']}")
        self.pages_label.config(text=f"Page Number(s): {chunk['page_num']}")
        self.docitem_label.config(text=f"DocItem Elements: {chunk['type_of_docitem']}")
        
        self.headings_entry.delete(0, tk.END)
        heading_val = chunk.get("heading")
        if isinstance(heading_val, list):
            headings_text = ", ".join(str(h) for h in heading_val if h)
        elif isinstance(heading_val, str):
            headings_text = heading_val
        else:
            headings_text = ""
            
        self.headings_entry.insert(0, headings_text)
        
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert(tk.END, chunk["chunk_text"])
        
        self.logger.info(f"UI rendering Chunk index: {chunk['chunk_index']}")

    def save_progress_to_json(self, status, headings, text_body):
        chunk = self.chunks_data[self.current_pointer]
        chunk_idx = chunk["chunk_index"]
        
        log_entry = {
            "chunk_index": chunk_idx,
            "status": status,
            "heading": headings,
            "chunk_text": text_body,
            "type_of_docitem": chunk["type_of_docitem"],
            "page_num": chunk["page_num"]
        }
        
        self.logged_chunks.append(log_entry)
        
        # --- Group and Merge Text Content by Common Headings ---
        merged_dict = {}
        for entry in self.logged_chunks:
            # Skip combining skipped blocks if you only want valid content
            if entry["status"] == "skipped":
                continue
                
            # Convert the heading list to a tuple to use as a reliable dict key
            h_key = tuple(entry["heading"])
            if h_key not in merged_dict:
                merged_dict[h_key] = {
                    "heading": entry["heading"],
                    "merged_text": entry["chunk_text"],
                    "chunk_indices": [entry["chunk_index"]],
                    "types_of_docitem": list(entry["type_of_docitem"]),
                    "page_nums": list(entry["page_num"]) if isinstance(entry["page_num"], list) else [entry["page_num"]]
                }
            else:
                # Merge the text cleanly under the same heading path
                merged_dict[h_key]["merged_text"] += "\n\n" + entry["chunk_text"]
                merged_dict[h_key]["chunk_indices"].append(entry["chunk_index"])
                
                for doc_type in entry["type_of_docitem"]:
                    if doc_type not in merged_dict[h_key]["types_of_docitem"]:
                        merged_dict[h_key]["types_of_docitem"].append(doc_type)
                        
                p_list = entry["page_num"] if isinstance(entry["page_num"], list) else [entry["page_num"]]
                for p in p_list:
                    if p not in merged_dict[h_key]["page_nums"]:
                        merged_dict[h_key]["page_nums"].append(p)

        # Build structured output payload
        output_payload = {
            "merged_headings": list(merged_dict.values()),
            "raw_session_history": self.logged_chunks
        }
        
        try:
            with open(self.output_file_name, 'w', encoding='utf-8') as outfile:
                json.dump(output_payload, outfile, indent=4, ensure_ascii=False)
            self.logger.info(f"Chunk {chunk_idx} saved. Heading merge compiled successfully.")
        except Exception as e:
            self.logger.error(f"Critical disk write error for chunk {chunk_idx}: {e}")

    def log_chunk(self):
        raw_headings = self.headings_entry.get().strip()
        headings_list = [h.strip() for h in raw_headings.split(",") if h.strip()]
        edited_text = self.text_area.get("1.0", tk.END).rstrip("\n")
        
        original_chunk = self.chunks_data[self.current_pointer]
        
        if headings_list != original_chunk["heading"] or edited_text != original_chunk["chunk_text"]:
            status = "logged (edited)"
        else:
            status = "logged"
            
        self.save_progress_to_json(status, headings_list, edited_text)
        
        self.current_pointer += 1
        self.advance_to_next_unprocessed()
        self.display_current_chunk()

    def skip_chunk(self):
        raw_headings = self.headings_entry.get().strip()
        headings_list = [h.strip() for h in raw_headings.split(",") if h.strip()]
        current_text = self.text_area.get("1.0", tk.END).rstrip("\n")
        
        self.save_progress_to_json("skipped", headings_list, current_text)
        
        self.current_pointer += 1
        self.advance_to_next_unprocessed()
        self.display_current_chunk()

    def use_previous_heading(self):
        if not self.logged_chunks:
            messagebox.showwarning("Action Invalid", "There is no previously logged chunk in this session to copy a heading from.")
            return

        last_logged_heading = self.logged_chunks[-1]["heading"]
        heading_str = ", ".join(last_logged_heading)

        self.headings_entry.delete(0, tk.END)
        self.headings_entry.insert(0, heading_str)
        
        self.logger.info(f"Present heading field overwritten with previous heading values: {last_logged_heading}")

    def reset_all_progress(self):
        confirm = messagebox.askyesno(
            "Confirm System Reset", 
            "Are you sure you want to clear all current entries and rewrite the progress file completely from the start?\n\nThis will permanently clear your choices.",
            icon="warning"
        )
        if not confirm:
            return

        self.logged_chunks.clear()
        self.processed_indices.clear()
        self.current_pointer = 0

        try:
            empty_payload = {"merged_headings": [], "raw_session_history": []}
            with open(self.output_file_name, 'w', encoding='utf-8') as outfile:
                json.dump(empty_payload, outfile, indent=4, ensure_ascii=False)
            
            self.logger.info("User triggered full sequence reset. Output tracking structures cleared.")
            messagebox.showinfo("Reset Complete", "The progress log file was cleared successfully. Restarting execution from Chunk 1.")
            self.display_current_chunk()
        except Exception as e:
            self.logger.error(f"Failed clearing progress file target: {e}")
            messagebox.showerror("File IO Error", f"Could not reset log data file: {e}")