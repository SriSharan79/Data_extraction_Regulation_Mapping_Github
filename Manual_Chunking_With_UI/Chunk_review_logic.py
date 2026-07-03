import sys

sys.path.extend([
    r'src',
    r'src/COLLECTION',
    r'Working_Code',
    r'src/DATA_ANALYSIS',
    r'src/COMMON',
    r'src/Command_Line_UI'
])
import os
import json
import re
import logging
import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox,filedialog, scrolledtext
from colorama import Fore, Style, init
from docling.document_converter import DocumentConverter
from docling.chunking import HybridChunker
from chunk_review_ui import ChunkReviewApp
import traceback
import hashlib


from Table_image_extractor import DoclingExtractor
from workspace_config import REGISTRY_FILE
# --- Initialize Colorama for background processing logs ---
init(autoreset=True)

class ExtractionLauncherUI:
    """Launcher screen to customize processing goals and path configurations dynamically."""
    def __init__(self, root):
        self.root = root
        self.root.title("Docling Automation & Curation Setup Launcher")
        self.root.geometry("680x420")
        self.root.configure(bg="#f4f4f6")
        
        # Section A: Fresh Document Processing
        frame_pdf = tk.LabelFrame(root, text=" Pipeline Run: Process PDF Document ", bg="white", font=("Arial", 10, "bold"), padx=10, pady=10)
        frame_pdf.pack(fill="x", padx=15, pady=10)
        
        tk.Label(frame_pdf, text="PDF File Path:", bg="white").grid(row=0, column=0, sticky="w")
        self.entry_pdf = tk.Entry(frame_pdf, width=50)
        self.entry_pdf.grid(row=0, column=1, padx=5, pady=2)
        tk.Button(frame_pdf, text="Browse...", command=self.browse_pdf).grid(row=0, column=2, padx=2)
        
        tk.Label(frame_pdf, text="Storage Destination:", bg="white").grid(row=1, column=0, sticky="w")
        self.entry_store_pdf = tk.Entry(frame_pdf, width=50)
        self.entry_store_pdf.grid(row=1, column=1, padx=5, pady=2)
        tk.Button(frame_pdf, text="Browse...", command=self.browse_storage_pdf).grid(row=1, column=2, padx=2)
        
        btn_box = tk.Frame(frame_pdf, bg="white")
        btn_box.grid(row=2, column=1, pady=8, sticky="w")
        self.btn_cache_only = tk.Button(btn_box, text="Generate Cache Only", bg="#3498db", fg="white", font=("Arial", 9, "bold"), command=self.run_cache_only)
        self.btn_cache_only.pack(side="left", padx=5)
        self.btn_full_pipeline = tk.Button(btn_box, text="Run Extraction + Review Process", bg="#2ecc71", fg="white", font=("Arial", 9, "bold"), command=self.run_full_pipeline)
        self.btn_full_pipeline.pack(side="left", padx=5)

        # Section B: Independent Review Process from Pre-existing Cache
        frame_cache = tk.LabelFrame(root, text=" Curation Only: Review Pre-existing Cache File From Scratch ", bg="white", font=("Arial", 10, "bold"), padx=10, pady=10)
        frame_cache.pack(fill="x", padx=15, pady=10)
        
        tk.Label(frame_cache, text="Cache JSON File:", bg="white").grid(row=0, column=0, sticky="w")
        self.entry_cache = tk.Entry(frame_cache, width=50)
        self.entry_cache.grid(row=0, column=1, padx=5, pady=2)
        tk.Button(frame_cache, text="Browse...", command=self.browse_cache).grid(row=0, column=2, padx=2)
        
        tk.Label(frame_cache, text="Storage Destination:", bg="white").grid(row=1, column=0, sticky="w")
        self.entry_store_cache = tk.Entry(frame_cache, width=50)
        self.entry_store_cache.grid(row=1, column=1, padx=5, pady=2)
        tk.Button(frame_cache, text="Browse...", command=self.browse_storage_cache).grid(row=1, column=2, padx=2)
        
        self.btn_review_cache = tk.Button(frame_cache, text="Launch Curation Review From Start", bg="#9b59b6", fg="white", font=("Arial", 9, "bold"), command=self.run_review_from_cache)
        self.btn_review_cache.grid(row=2, column=1, pady=8, sticky="w", padx=5)

        # Buttons disabled while a background conversion is running.
        self.action_buttons = [self.btn_cache_only, self.btn_full_pipeline, self.btn_review_cache]
        self._busy = False

        # Status line for background-task feedback.
        self.status_label = tk.Label(root, text="", bg="#f4f4f6", fg="#2c3e50", font=("Arial", 9, "italic"), anchor="w")
        self.status_label.pack(fill="x", padx=15, pady=(0, 8))

    # --- Background-task helpers ------------------------------------------- #
    def _set_busy(self, busy, message=""):
        """Toggle the UI busy state (main thread only)."""
        self._busy = busy
        state = "disabled" if busy else "normal"
        for btn in getattr(self, "action_buttons", []):
            try:
                btn.config(state=state)
            except tk.TclError:
                pass
        try:
            self.status_label.config(text=message)
            self.root.update_idletasks()
        except tk.TclError:
            pass

    def _run_in_background(self, work_fn, on_done, busy_message="Processing, please wait..."):
        """Run work_fn() off the UI thread; call on_done(result, error) back on
        the UI thread via after(). Keeps the docling conversion from freezing
        the window."""
        if self._busy:
            messagebox.showinfo("Please wait", "A processing task is already running.")
            return

        self._set_busy(True, busy_message)
        holder = {}

        def worker():
            try:
                holder["result"] = work_fn()
            except Exception as exc:  # surfaced to on_done on the UI thread
                holder["error"] = exc
            finally:
                holder["done"] = True

        threading.Thread(target=worker, daemon=True).start()

        def poll():
            if holder.get("done"):
                self._set_busy(False, "")
                on_done(holder.get("result"), holder.get("error"))
            else:
                self.root.after(120, poll)

        self.root.after(120, poll)

    def browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Documents", "*.pdf")])
        if path:
            self.entry_pdf.delete(0, tk.END)
            self.entry_pdf.insert(0, path)
            # Check for history and automatically fill out destination space
            self.auto_populate_storage(path, self.entry_store_pdf)
            
    def browse_cache(self):
        path = filedialog.askopenfilename(filetypes=[("JSON Cache Files", "*.json")])
        if path:
            self.entry_cache.delete(0, tk.END)
            self.entry_cache.insert(0, path)
            # Check for history and automatically fill out destination space
            self.auto_populate_storage(path, self.entry_store_cache)
            
    def browse_storage_pdf(self):
        path = filedialog.askdirectory()
        if path:
            self.entry_store_pdf.delete(0, tk.END)
            self.entry_store_pdf.insert(0, path)

            
    def browse_storage_cache(self):
        path = filedialog.askdirectory()
        if path:
            self.entry_store_cache.delete(0, tk.END)
            self.entry_store_cache.insert(0, path)

    def resolve_directory_structure(self, storage_path, reference_path):
        """Unified path resolution utility shared across the ecosystem."""
        base_filename = os.path.splitext(os.path.basename(reference_path))[0]
        
        # If running from a cache file, strip the suffix to recover the root PDF name
        if base_filename.endswith("_docling_chunks_cache"):
            doc_name = os.path.basename(os.path.dirname(reference_path))
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

    def init_runtime_logger(self, log_path):
        logger = logging.getLogger("DoclingRuntimeUI")
        logger.setLevel(logging.INFO)
        # Clear past execution streams safely
        if logger.hasHandlers():
            logger.handlers.clear()
        handler = logging.FileHandler(log_path, encoding='utf-8')
        handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(handler)
        return logger

    def run_cache_only(self):
        pdf_path = self.entry_pdf.get().strip()
        storage_path = self.entry_store_pdf.get().strip()
        
        if not pdf_path or not storage_path:
            messagebox.showerror("Error", "Please clarify valid target PDF and storage folders.")
            return
            
        storage_path = self.validate_and_confirm_storage(storage_path, pdf_path, self.entry_store_pdf)
        paths = self.resolve_directory_structure(storage_path, pdf_path)
        
        # Ensure directories exist before generating logs
        os.makedirs(paths["dated_folder"], exist_ok=True)
        logger = self.init_runtime_logger(paths["log_file"])

        print(Fore.YELLOW + "Executing background extraction task pipeline...")

        def work():
            return generate_and_cache_document(pdf_path, paths, logger)

        def done(chunks, error):
            if error is not None:
                logger.critical(f"Cache generation crashed: {error}")
                messagebox.showerror("Execution Fail", f"Extraction pipeline crashed:\n{error}")
                return
            if chunks:
                messagebox.showinfo("Success", f"Cache structure created successfully!\nFile path:\n{paths['cache_file']}")
            else:
                messagebox.showerror("Execution Fail", "Extraction pipelines crashed. Review execution log.")

        self._run_in_background(work, done, "Generating cache — converting document, please wait...")

    def run_full_pipeline(self):
        pdf_path = self.entry_pdf.get().strip()
        storage_path = self.entry_store_pdf.get().strip()
        
        if not pdf_path or not storage_path:
            messagebox.showerror("Error", "Please clarify valid target PDF and storage folders.")
            return
            
        storage_path = self.validate_and_confirm_storage(storage_path, pdf_path, self.entry_store_pdf)
        paths = self.resolve_directory_structure(storage_path, pdf_path)
        
        os.makedirs(paths["dated_folder"], exist_ok=True)
        logger = self.init_runtime_logger(paths["log_file"])

        print(Fore.CYAN + "Loading active document tracking logs...")

        def work():
            return generate_and_cache_document(pdf_path, paths, logger)

        def done(chunks_data, error):
            if error is not None:
                logger.critical(f"Extraction crashed: {error}")
                messagebox.showerror("Error", f"Extraction process failed:\n{error}")
                return
            if not chunks_data:
                messagebox.showerror("Error", "Extraction process failed.")
                return

            logged_chunks, processed_indices = load_existing_progress_log(paths["output_file"], logger)

            self.root.destroy()
            launch_review_app(chunks_data, logged_chunks, processed_indices, paths["output_file"], logger)

        self._run_in_background(work, done, "Running extraction — converting document, please wait...")

    def run_review_from_cache(self):
        cache_path = self.entry_cache.get().strip()
        storage_path = self.entry_store_cache.get().strip()
        
        if not cache_path or not storage_path or not os.path.exists(cache_path):
            messagebox.showerror("Error", "Please provide a valid pre-existing cache file and storage directory destination.")
            return
            
        storage_path = self.validate_and_confirm_storage(storage_path, cache_path, self.entry_store_cache)
        paths = self.resolve_directory_structure(storage_path, cache_path)
        
        os.makedirs(paths["dated_folder"], exist_ok=True)
        logger = self.init_runtime_logger(paths["log_file"])
        
        print(Fore.MAGENTA + "Reading configuration tracks from chosen custom cache file...")
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                payload = json.load(f)
                chunks_data = payload.get("chunks", []) if isinstance(payload, dict) else payload
        except Exception as e:
            messagebox.showerror("Cache Error", f"Failed parsing chosen cache file format: {e}")
            return
            
        if not chunks_data:
            messagebox.showerror("Data Error", "No data content allocations parsed inside the provided cache file.")
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
                        f"Found an existing progress log with {len(history_records)} reviewed chunks.\n\n"
                        "Do you want to CONTINUE from where you left off?\n"
                        "(Selecting 'No' will reset all progress and start from the beginning.)"
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
                logger.warning(f"Error parsing existing output progress logs: {e}")
        
        if not os.path.exists(paths["output_file"]):
            try:
                with open(paths["output_file"], 'w', encoding='utf-8') as outfile:
                    json.dump({"merged_headings": [], "raw_session_history": []}, outfile, indent=4, ensure_ascii=False)
            except Exception as e:
                logger.error(f"Failed creating empty log target: {e}")

        self.root.destroy()
        launch_review_app(chunks_data, logged_chunks, processed_indices, paths["output_file"], logger)  



    def get_registered_storage(self, file_path):
        """Reads the global registry to find a previously paired storage location."""
        if os.path.exists(REGISTRY_FILE):
            try:
                with open(REGISTRY_FILE, 'r', encoding='utf-8') as f:
                    reg = json.load(f)
                    return reg.get(os.path.abspath(file_path))
            except Exception:
                pass
        return None

    def save_to_registry(self, file_path, storage_path):
        """Saves or updates the file-to-storage mapping in the registry."""
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

    def auto_populate_storage(self, file_path, entry_widget):
        """Auto-fills the UI storage entry if a history matching the file exists."""
        old_storage = self.get_registered_storage(file_path)
        if old_storage:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, old_storage)
            print(Fore.GREEN + f"[Registry] Auto-loaded previous workspace track: {old_storage}")


    def validate_and_confirm_storage(self, storage_path, reference_path, entry_widget):
        """Validates storage targets and syncs with workspace memory maps."""
        paths = self.resolve_directory_structure(storage_path, reference_path)
        
        if os.path.exists(paths["output_file"]) or os.path.exists(paths["cache_file"]):
            use_older = messagebox.askyesno(
                "Existing Storage Footprint",
                f"Processing files or logged histories already exist for this reference inside:\n{paths['root']}\n\n"
                "Do you want to continue using this older destination folder?\n"
                "(Selecting 'No' allows you to select a new base storage folder track entirely.)"
            )
            
            if not use_older:
                chosen_dir = filedialog.askdirectory(title="Select New Base Storage Destination Folder")
                if chosen_dir:
                    entry_widget.delete(0, tk.END)
                    entry_widget.insert(0, chosen_dir)
                    storage_path = chosen_dir

        self.save_to_registry(reference_path, storage_path)
        return storage_path


def load_existing_progress_log(output_file_name, logger):
    logged_chunks = []
    processed_indices = set()
    if os.path.exists(output_file_name):
        try:
            with open(output_file_name, 'r', encoding='utf-8') as infile:
                data = json.load(infile)
                
            # Safely unpack raw session history from new dict structure; fallback to legacy lists
            if isinstance(data, dict) and "raw_session_history" in data:
                logged_chunks = data["raw_session_history"]
            elif isinstance(data, list):
                logged_chunks = data
                
            for entry in logged_chunks:
                if "chunk_index" in entry:
                    processed_indices.add(entry["chunk_index"])
            print(Fore.CYAN + f">>> Resume validation checkpoint: found {len(processed_indices)} handled tracking elements.")
        except Exception as e:
            logger.warning(f"Failed to read existing progress log data tracks: {e}")
    return logged_chunks, processed_indices


def generate_and_cache_document(pdf_path, paths, logger):
    """Processes document files through pipeline architectures and dumps metrics directly into the structured cache layer."""
    chunks_data = []
    tables_data = []
    images_data = []
    headings_data = []

    if os.path.exists(paths["cache_file"]):
        print(Fore.GREEN + Style.BRIGHT + f">>> Fetching structural entries from existing track cache path...")
        try:
            with open(paths["cache_file"], 'r', encoding='utf-8') as cache_file:
                cache_payload = json.load(cache_file)
                if isinstance(cache_payload, dict) and "chunks" in cache_payload:
                    return cache_payload.get("chunks", [])
                return cache_payload
        except Exception as e:
            logger.warning(f"Failed to cleanly process storage cache files: {e}")
            traceback.print_exc()

    # Cache miss execution route
    logger.info("Cache miss. Initializing DoclingExtractor parsing pipeline loops.")
    try:
        chunker = HybridChunker()
        extractor = DoclingExtractor(
            input_path=pdf_path,
            tables_output_path=paths["tables_path"],
            images_output_path=paths["images_path"]
        )
        
        logger.info(f"Starting single-pass document conversion logic configuration for: {pdf_path}")
        conversion_result = extractor.doc_converter.convert(pdf_path)
        doc = conversion_result.document
        
        logger.info("Extracting tables to sub-dated storage track...")
        tables_data = extractor._extract_tables(conversion_result)
        
        logger.info("Extracting images to sub-dated storage track...")
        images_data = extractor._extract_images(conversion_result)
        
        logger.info("Extracting layout headings records...")
        headings_data = extractor._extract_headings(conversion_result)
        
        logger.info("Executing text chunker token parsing engine rules...")
        raw_chunks = list(chunker.chunk(dl_doc=doc))
        
        for i, chunk in enumerate(raw_chunks):
            chunk_pages = set()
            doc_items_list = []
            
            if hasattr(chunk.meta, 'doc_items'):
                for item in chunk.meta.doc_items:
                    item_text = None
                    self_ref = getattr(item, 'self_ref', None)
                    if self_ref:
                        match = re.match(r'#/texts/(\d+)', self_ref)
                        if match:
                            text_index = int(match.group(1))
                            try:
                                item_text = doc.texts[text_index].text
                            except (IndexError, AttributeError):
                                item_text = None

                    item_pages = []
                    prov_list = getattr(item, 'prov', [])
                    if prov_list:
                        for prov_item in prov_list:
                            p_no = prov_item.get("page_no") if isinstance(prov_item, dict) else getattr(prov_item, "page_no", None)
                            if p_no is not None:
                                item_pages.append(p_no)
                                chunk_pages.add(p_no)

                    item_dict = {
                        "label": getattr(item, 'label', 'N/A'),
                        "self_ref": self_ref,
                        "actual_text": item_text,
                        "page_no": item_pages[0] if item_pages else None,
                        "all_pages": item_pages
                    }
                    doc_items_list.append(item_dict)
            
            sorted_pages = sorted(list(chunk_pages))
            headings = chunk.meta.headings if hasattr(chunk.meta, 'headings') else []
            doc_item_labels = [item["label"] for item in doc_items_list]

            chunk_dict = {
                "chunk_index": i + 1,
                "type": type(chunk).__name__,
                "heading": headings,
                "chunk_text": chunk.text,
                "type_of_docitem": doc_item_labels if doc_item_labels else ["N/A"],
                "page_num": sorted_pages,
                "text": chunk.text,
                "meta": {
                    "headings": headings,
                    "doc_items": doc_items_list,
                    "page_numbers": sorted_pages
                }
            }
            chunks_data.append(chunk_dict)

        logger.info("Cache extraction sequence run complete.")
        cache_payload = {
            "chunks": chunks_data,
            "tables": tables_data,
            "images": images_data,
            "headings": headings_data
        }
        
        with open(paths["cache_file"], 'w', encoding='utf-8') as cache_file:
            json.dump(cache_payload, cache_file, indent=4, ensure_ascii=False)
        print(Fore.GREEN + f"Backend parsing complete. Structure saved directly to cache path location.")
        return chunks_data
    except Exception as e:
        logger.critical(f"Pipeline processing failed to execute cleanly: {e}")
        traceback.print_exc()
        return []


def launch_review_app(chunks_data, logged_chunks, processed_indices, output_file, logger):
    """Triggers the curation user interface window."""
    review_window = tk.Tk()
    app = ChunkReviewApp(
        root=review_window,
        chunks_data=chunks_data,
        logged_chunks=logged_chunks,
        processed_indices=processed_indices,
        output_file_name=output_file,
        logger=logger
    )
    review_window.mainloop()


def main():
    root = tk.Tk()
    launcher = ExtractionLauncherUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()