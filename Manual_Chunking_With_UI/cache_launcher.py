import hashlib
import os
import json
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

from workspace_config import REGISTRY_FILE

# Import the existing app hand-off hook from your main file
try:
    from Chunk_review_logic import launch_review_app
except ImportError:
    messagebox.showerror(
        "Import Error",
        "Could not find 'Chunk_review_logic' in the current working directory.\n"
        "Please ensure this script is saved in the same folder as your existing modules."
    )
    exit()
class CacheReviewLauncher:

    def __init__(self, root):
        self.root = root
        self.root.title("Docling Cache Curation Review Launcher")
        self.root.geometry("680x280")
        self.root.resizable(False, False)
        
        # Main Container UI Layout
        main_frame = tk.Frame(root, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- Section 1: Input Cache Selection ---
        tk.Label(main_frame, text="Select Pre-Existing JSON Cache File:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.entry_cache = tk.Entry(main_frame, width=65)
        self.entry_cache.grid(row=1, column=0, padx=(0, 10), pady=(0, 15), sticky="we")
        btn_browse_cache = tk.Button(main_frame, text="Browse Cache...", command=self.browse_cache, width=15)
        btn_browse_cache.grid(row=1, column=1, pady=(0, 15))
        
        # --- Section 2: Storage Destination Selection ---
        tk.Label(main_frame, text="Base Storage Destination Directory:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(0, 2))
        self.entry_store = tk.Entry(main_frame, width=65)
        self.entry_store.grid(row=3, column=0, padx=(0, 10), pady=(0, 20), sticky="we")
        btn_browse_store = tk.Button(main_frame, text="Browse Storage...", command=self.browse_storage, width=15)
        btn_browse_store.grid(row=3, column=1, pady=(0, 20))
        
        # --- Section 3: Execution Trigger ---
        self.btn_launch = tk.Button(
            main_frame, 
            text="🚀 Start / Resume Curation Review", 
            font=("Arial", 11, "bold"), 
            bg="#2e7d32", 
            fg="white", 
            activebackground="#1b5e20", 
            activeforeground="white",
            command=self.execute_review_session,
            pady=6
        )
        self.btn_launch.grid(row=4, column=0, columnspan=2, sticky="we")

    # --- Registry & History Lookup Utilities ---
    def get_registered_storage(self, file_path):
        if os.path.exists(REGISTRY_FILE):
            try:
                with open(REGISTRY_FILE, 'r', encoding='utf-8') as f:
                    reg = json.load(f)
                    return reg.get(os.path.abspath(file_path))
            except Exception:
                pass
        return None

    def save_to_registry(self, file_path, storage_path):
        reg = {}
        if os.path.exists(REGISTRY_FILE):
            try:
                with open(REGISTRY_FILE, 'r', encoding='utf-8') as f:
                    reg = json.load(f)
            except Exception:
                pass
        reg[os.path.abspath(file_path)] = os.path.abspath(storage_path)
        try:
            with open(REGISTRY_FILE, 'w', encoding='utf-8') as f:
                json.dump(reg, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

    def auto_populate_storage(self, file_path):
        old_storage = self.get_registered_storage(file_path)
        if old_storage:
            self.entry_store.delete(0, tk.END)
            self.entry_store.insert(0, old_storage)

    # --- UI Event Handlers ---
    def browse_cache(self):
        path = filedialog.askopenfilename(filetypes=[("JSON Cache Files", "*.json")])
        if path:
            self.entry_cache.delete(0, tk.END)
            self.entry_cache.insert(0, path)
            self.auto_populate_storage(path)
            
    def browse_storage(self):
        path = filedialog.askdirectory(title="Select Base Storage Destination Folder")
        if path:
            self.entry_store.delete(0, tk.END)
            self.entry_store.insert(0, path)

    def resolve_paths(self, storage_path, cache_path):
        """Matches the path resolution logic from main.py to share workspaces seamlessly."""
        base_filename = os.path.splitext(os.path.basename(cache_path))[0]
        
        if base_filename.endswith("_docling_chunks_cache"):
            doc_name = os.path.basename(os.path.dirname(cache_path))
        else:
            doc_name = base_filename

        target_root = os.path.join(storage_path, doc_name)
        current_date_str = datetime.now().strftime("%Y-%m-%d")
        dated_subfolder = os.path.join(target_root, current_date_str)
        base_hash = hashlib.md5(doc_name.encode()).hexdigest()[:8]
        
        return {
            "root": target_root,
            "dated_folder": dated_subfolder,
            "cache_file": os.path.join(target_root, f"{base_hash}_docling_chunks_cache.json"),
            "output_file": os.path.join(dated_subfolder, f"{base_hash}_docling_logged_chunks.json"),
            "log_file": os.path.join(dated_subfolder, f"{base_hash}_docling_execution.log"),
            "tables_path": os.path.join(dated_subfolder, "tables"),
            "images_path": os.path.join(dated_subfolder, "images")
        }

    def validate_and_confirm_storage(self, storage_path, cache_path):
        paths = self.resolve_paths(storage_path, cache_path)
        
        if os.path.exists(paths["output_file"]) or os.path.exists(paths["cache_file"]):
            use_older = messagebox.askyesno(
                "Existing Storage Footprint",
                f"Processing footprints already exist for this reference inside:\n{paths['root']}\n\n"
                "Do you want to continue using this previous destination folder?\n"
                "(Selecting 'No' redirects you to choose a new base location track entirely.)"
            )
            if not use_older:
                chosen_dir = filedialog.askdirectory(title="Select New Base Storage Destination Folder")
                if chosen_dir:
                    self.entry_store.delete(0, tk.END)
                    self.entry_store.insert(0, chosen_dir)
                    storage_path = chosen_dir
                    
        self.save_to_registry(cache_path, storage_path)
        return storage_path

    def execute_review_session(self):
        cache_path = self.entry_cache.get().strip()
        storage_path = self.entry_store.get().strip()
        
        if not cache_path or not storage_path or not os.path.exists(cache_path):
            messagebox.showerror("Error", "Please clarify a valid cache file path and base storage folder target.")
            return

        storage_path = self.validate_and_confirm_storage(storage_path, cache_path)
        paths = self.resolve_paths(storage_path, cache_path)
        os.makedirs(paths["dated_folder"], exist_ok=True)

        # Setup runtime log inside the common dated subfolder
        logger = logging.getLogger("CacheLauncher")
        logger.setLevel(logging.INFO)
        if not logger.handlers:
            fh = logging.FileHandler(paths["log_file"], encoding="utf-8")
            fh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            logger.addHandler(fh)

        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                payload = json.load(f)
                chunks_data = payload.get("chunks", []) if isinstance(payload, dict) else payload
        except Exception as e:
            messagebox.showerror("Cache Parsing Error", f"Failed to ingest raw JSON structure details: {e}")
            return

        if not chunks_data:
            messagebox.showerror("Data Error", "No raw text items or structure mappings extracted in cache arrays.")
            return

        logged_chunks = []
        processed_indices = set()

        if os.path.exists(paths["output_file"]):
            try:
                with open(paths["output_file"], 'r', encoding='utf-8') as infile:
                    existing_data = json.load(infile)

                history_records = existing_data.get("raw_session_history", []) if isinstance(existing_data, dict) else existing_data

                if len(history_records) > 0:
                    ask_resume = messagebox.askyesno(
                        "Previous Progress Detected",
                        f"Found an existing progress track containing {len(history_records)} completed blocks.\n\n"
                        "Do you want to CONTINUE from where you left off?\n"
                        "(Selecting 'No' permanently clears progress structures to reset back to Chunk 1.)"
                    )
                    if ask_resume:
                        logged_chunks = history_records
                        for entry in logged_chunks:
                            if "chunk_index" in entry:
                                processed_indices.add(entry["chunk_index"])
                    else:
                        with open(paths["output_file"], 'w', encoding='utf-8') as outfile:
                            json.dump({"merged_headings": [], "raw_session_history": []}, outfile, indent=4, ensure_ascii=False)
            except Exception as e:
                logger.warning(f"Error checking historical track logs: {e}")

        if not os.path.exists(paths["output_file"]):
            try:
                with open(paths["output_file"], 'w', encoding='utf-8') as outfile:
                    json.dump({"merged_headings": [], "raw_session_history": []}, outfile, indent=4, ensure_ascii=False)
            except Exception as e:
                logger.error(f"Failed to safely allocate empty target tracking file space: {e}")

        self.root.destroy()
        launch_review_app(chunks_data, logged_chunks, processed_indices, paths["output_file"], logger)

if __name__ == "__main__":
    main_window = tk.Tk()
    app = CacheReviewLauncher(main_window)
    main_window.mainloop()