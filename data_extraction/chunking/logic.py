import os
import json
import re
import logging
import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, filedialog, scrolledtext, ttk
from colorama import Fore, Style, init
# NOTE: docling is heavyweight and only needed when an extraction actually runs.
# It is imported lazily inside generate_and_cache_document() so this module (and
# the launcher UIs that import it) load even when docling is not installed.
from .chunk_review_ui import ChunkReviewApp
from .section_review_ui import SectionReviewApp
import traceback
import hashlib


from .table_image_extractor import DoclingExtractor
from .workspace_config import REGISTRY_FILE
# --- Initialize Colorama for background processing logs ---
init(autoreset=True)

class ExtractionLauncherUI:
    """Launcher screen to customize processing goals and path configurations dynamically."""
    def __init__(self, root):
        self.root = root
        self.root.title("Docling Automation & Curation Setup Launcher")
        self.root.geometry("760x460")

        # Section A: Fresh Document Processing
        frame_pdf = ttk.LabelFrame(
            root, text=" Pipeline Run: Process PDF Document ", padding=10)
        frame_pdf.pack(fill="x", padx=10, pady=10)

        ttk.Label(frame_pdf, text="PDF file:").grid(row=0, column=0, sticky="w")
        self.entry_pdf = ttk.Entry(frame_pdf, width=78)
        self.entry_pdf.grid(row=0, column=1, padx=5, pady=3, sticky="we")
        ttk.Button(frame_pdf, text="Browse...", command=self.browse_pdf).grid(row=0, column=2)

        ttk.Label(frame_pdf, text="Storage destination:").grid(row=1, column=0, sticky="w")
        self.entry_store_pdf = ttk.Entry(frame_pdf, width=78)
        self.entry_store_pdf.grid(row=1, column=1, padx=5, pady=3, sticky="we")
        ttk.Button(frame_pdf, text="Browse...", command=self.browse_storage_pdf).grid(row=1, column=2)

        btn_box = ttk.Frame(frame_pdf)
        btn_box.grid(row=2, column=1, pady=8, sticky="w")
        self.btn_cache_only = ttk.Button(btn_box, text="Generate Cache Only",
                                         command=self.run_cache_only)
        self.btn_cache_only.pack(side="left")
        self.btn_full_pipeline = ttk.Button(btn_box, text="Run Extraction + Review Process",
                                            command=self.run_full_pipeline)
        self.btn_full_pipeline.pack(side="left", padx=5)
        frame_pdf.columnconfigure(1, weight=1)

        # Section B: Independent Review Process from Pre-existing Cache
        frame_cache = ttk.LabelFrame(
            root, text=" Curation of the Extracted Chunks (from an existing cache) ", padding=10)
        frame_cache.pack(fill="x", padx=10, pady=10)

        ttk.Label(frame_cache, text="Cache JSON file:").grid(row=0, column=0, sticky="w")
        self.entry_cache = ttk.Entry(frame_cache, width=78)
        self.entry_cache.grid(row=0, column=1, padx=5, pady=3, sticky="we")
        ttk.Button(frame_cache, text="Browse...", command=self.browse_cache).grid(row=0, column=2)

        ttk.Label(frame_cache, text="Storage destination:").grid(row=1, column=0, sticky="w")
        self.entry_store_cache = ttk.Entry(frame_cache, width=78)
        self.entry_store_cache.grid(row=1, column=1, padx=5, pady=3, sticky="we")
        ttk.Button(frame_cache, text="Browse...", command=self.browse_storage_cache).grid(row=1, column=2)

        ttk.Label(
            frame_cache,
            text="Always choose the storage folder that was used to generate the cache file.",
            foreground="#b3541e",
        ).grid(row=2, column=1, sticky="w", pady=(4, 0))
        ttk.Label(
            frame_cache,
            text=("Note: raw chunks in the cache JSON are processed into a structured "
                  "output;\nonly the processed chunks land in Processed_chunks.json "
                  "inside the dated folder."),
            foreground="#666666",
            justify="left",
        ).grid(row=3, column=1, sticky="w", pady=(2, 0))

        self.btn_review_cache = ttk.Button(frame_cache, text="Launch Chunk Curation Review",
                                           command=self.run_review_from_cache)
        self.btn_review_cache.grid(row=4, column=1, pady=8, sticky="w")
        frame_cache.columnconfigure(1, weight=1)

        # Buttons disabled while a background conversion is running.
        self.action_buttons = [self.btn_cache_only, self.btn_full_pipeline, self.btn_review_cache]
        self._busy = False

        # Status line for background-task feedback.
        self.status_label = ttk.Label(root, text="", foreground="#2c3e50", anchor="w")
        self.status_label.pack(fill="x", padx=12, pady=(0, 8))

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
            "output_file": os.path.join(dated_subfolder, f"Processed_chunks.json"),
            "log_file": os.path.join(target_root, f"{base_hash}_Review_execution.log"),
            "tables_path": os.path.join(dated_subfolder, "tables"),
            "images_path": os.path.join(dated_subfolder, "images")
        }

    def validate_and_confirm_storage(self, storage_path, reference_path):
        """Warn when the chosen storage already holds a processing footprint
        for this reference and offer to pick a different base folder; then
        remember the mapping in the registry (adopted from the former
        CacheReviewLauncher so its behaviour lives on here)."""
        paths = self.resolve_directory_structure(storage_path, reference_path)

        if os.path.exists(paths["output_file"]) or os.path.exists(paths["cache_file"]):
            use_older = messagebox.askyesno(
                "Existing Storage Footprint",
                f"Processing footprints already exist for this reference inside:\n{paths['root']}\n\n"
                "Do you want to continue using this previous destination folder?\n"
                "(Selecting 'No' redirects you to choose a new base location entirely.)"
            )
            if not use_older:
                chosen_dir = filedialog.askdirectory(title="Select New Base Storage Destination Folder")
                if chosen_dir:
                    self.entry_store_cache.delete(0, tk.END)
                    self.entry_store_cache.insert(0, chosen_dir)
                    storage_path = chosen_dir

        self.save_to_registry(reference_path, storage_path)
        return storage_path

    def run_cache_only(self):
        pdf_path = self.entry_pdf.get().strip()
        storage_path = self.entry_store_pdf.get().strip()

        if not pdf_path or not storage_path:
            messagebox.showerror("Missing input", "Please provide a PDF file and storage destination.")
            return
        if not os.path.exists(pdf_path):
            messagebox.showerror("File not found", f"PDF file does not exist:\n{pdf_path}")
            return
        self.save_to_registry(pdf_path, storage_path)

        def job():
            paths = self.resolve_directory_structure(storage_path, pdf_path)
            os.makedirs(paths["dated_folder"], exist_ok=True)
            
            logger = logging.getLogger("ExtractionLauncher")
            logger.setLevel(logging.INFO)
            if not logger.handlers:
                fh = logging.FileHandler(paths["log_file"], encoding="utf-8")
                fh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
                logger.addHandler(fh)

            chunks_data = generate_and_cache_document(pdf_path, paths, logger)
            print(Fore.GREEN + f"✓ Cache generation complete. {len(chunks_data)} chunks saved.")

        def on_done(result, error):
            if error:
                messagebox.showerror("Extraction Error", f"Extraction failed:\n{error}")
            else:
                messagebox.showinfo("Success", "Cache generation complete!")

        self._run_in_background(job, on_done, "Generating cache, please wait...")

    def run_full_pipeline(self):
        pdf_path = self.entry_pdf.get().strip()
        storage_path = self.entry_store_pdf.get().strip()
        
        if not pdf_path or not storage_path:
            messagebox.showerror("Missing input", "Please provide a PDF file and storage destination.")
            return
        if not os.path.exists(pdf_path):
            messagebox.showerror("File not found", f"PDF file does not exist:\n{pdf_path}")
            return
        self.save_to_registry(pdf_path, storage_path)

        def job():
            paths = self.resolve_directory_structure(storage_path, pdf_path)
            os.makedirs(paths["dated_folder"], exist_ok=True)

            logger = logging.getLogger("ExtractionLauncher")
            logger.setLevel(logging.INFO)
            if not logger.handlers:
                fh = logging.FileHandler(paths["log_file"], encoding="utf-8")
                fh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
                logger.addHandler(fh)

            chunks_data = generate_and_cache_document(pdf_path, paths, logger)
            print(Fore.GREEN + f"✓ Extraction complete. {len(chunks_data)} chunks extracted.")
            return (chunks_data, paths)

        def on_done(result, error):
            if error:
                messagebox.showerror("Extraction Error", f"Extraction failed:\n{error}")
                return
            
            chunks_data, paths = result
            self.root.destroy()
            launch_review_app(chunks_data, [], set(), paths["output_file"], logging.getLogger("ExtractionLauncher"))

        self._run_in_background(job, on_done, "Extracting document, please wait...")

    def run_review_from_cache(self):
        cache_path = self.entry_cache.get().strip()
        storage_path = self.entry_store_cache.get().strip()
        
        if not cache_path or not storage_path or not os.path.exists(cache_path):
            messagebox.showerror("Error", "Please provide a valid cache file path and base storage folder.")
            return

        storage_path = self.validate_and_confirm_storage(storage_path, cache_path)
        paths = self.resolve_directory_structure(storage_path, cache_path)
        os.makedirs(paths["dated_folder"], exist_ok=True)

        logger = logging.getLogger("ExtractionLauncher")
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
            messagebox.showerror("Cache Parsing Error", f"Failed to parse cache file:\n{e}")
            return

        if not chunks_data:
            messagebox.showerror("Data Error", "No chunks found in cache file.")
            return

        logged_chunks = []
        processed_indices = set()

        if os.path.exists(paths["output_file"]):
            try:
                with open(paths["output_file"], 'r', encoding='utf-8') as infile:
                    existing_data = json.load(infile)
                    history_records = existing_data.get("raw_session_history", []) if isinstance(existing_data, dict) else existing_data

                    if history_records:
                        ask_resume = messagebox.askyesno(
                            "Previous Progress Detected",
                            f"Found {len(history_records)} completed chunks.\n\nContinue from where you left off?"
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
                logger.warning(f"Error checking existing progress: {e}")

        if not os.path.exists(paths["output_file"]):
            try:
                with open(paths["output_file"], 'w', encoding='utf-8') as outfile:
                    json.dump({"merged_headings": [], "raw_session_history": []}, outfile, indent=4, ensure_ascii=False)
            except Exception as e:
                logger.error(f"Failed to create output file: {e}")

        self.root.destroy()
        launch_review_app(chunks_data, logged_chunks, processed_indices, paths["output_file"], logger)

    def auto_populate_storage(self, file_path, entry_widget):
        """Auto-populate storage from registry."""
        if os.path.exists(REGISTRY_FILE):
            try:
                with open(REGISTRY_FILE, 'r', encoding='utf-8') as f:
                    reg = json.load(f)
                    old_storage = reg.get(os.path.abspath(file_path))
                    if old_storage:
                        entry_widget.delete(0, tk.END)
                        entry_widget.insert(0, old_storage)
            except Exception:
                pass

    def save_to_registry(self, file_path, storage_path):
        """Save file-storage mapping to registry."""
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


def load_existing_progress_log(output_file_name, logger):
    """Load existing progress from output file."""
    logged_chunks = []
    processed_indices = set()
    if os.path.exists(output_file_name):
        try:
            with open(output_file_name, 'r', encoding='utf-8') as infile:
                data = json.load(infile)
                
            if isinstance(data, dict) and "raw_session_history" in data:
                logged_chunks = data["raw_session_history"]
            elif isinstance(data, list):
                logged_chunks = data
                
            for entry in logged_chunks:
                if "chunk_index" in entry:
                    processed_indices.add(entry["chunk_index"])
            print(Fore.CYAN + f">>> Resume checkpoint: found {len(processed_indices)} processed chunks.")
        except Exception as e:
            logger.warning(f"Failed to read existing progress: {e}")
    return logged_chunks, processed_indices


def generate_and_cache_document(pdf_path, paths, logger):
    """Process PDF and generate cache file."""
    chunks_data = []
    tables_data = []
    images_data = []
    headings_data = []

    if os.path.exists(paths["cache_file"]):
        print(Fore.GREEN + Style.BRIGHT + f">>> Using existing cache file...")
        try:
            with open(paths["cache_file"], 'r', encoding='utf-8') as cache_file:
                cache_payload = json.load(cache_file)
                if isinstance(cache_payload, dict) and "chunks" in cache_payload:
                    return cache_payload.get("chunks", [])
                return cache_payload
        except Exception as e:
            logger.warning(f"Failed to load cache: {e}")

    logger.info("Cache miss. Initializing DoclingExtractor parsing pipeline.")
    try:
        from docling.chunking import HybridChunker  # heavy import, deferred to run time

        chunker = HybridChunker()
        extractor = DoclingExtractor(
            input_path=pdf_path,
            tables_output_path=paths["tables_path"],
            images_output_path=paths["images_path"]
        )
        
        logger.info(f"Starting document conversion for: {pdf_path}")
        conversion_result = extractor.doc_converter.convert(pdf_path)
        doc = conversion_result.document
        
        logger.info("Extracting tables...")
        tables_data = extractor._extract_tables(conversion_result)
        
        logger.info("Extracting images...")
        images_data = extractor._extract_images(conversion_result)
        
        logger.info("Extracting headings...")
        headings_data = extractor._extract_headings(conversion_result)
        
        logger.info("Executing chunking engine...")
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

        logger.info("Cache extraction complete.")
        cache_payload = {
            "chunks": chunks_data,
            "tables": tables_data,
            "images": images_data,
            "headings": headings_data
        }
        
        with open(paths["cache_file"], 'w', encoding='utf-8') as cache_file:
            json.dump(cache_payload, cache_file, indent=4, ensure_ascii=False)
        print(Fore.GREEN + f">>> Cache saved successfully.")
        return chunks_data
    except Exception as e:
        logger.critical(f"Pipeline processing failed: {e}")
        traceback.print_exc()
        return []


def _triage_llm(logger):
    """An ``llm(prompt, system_prompt)`` callable for the triage's heading
    check, or None when the LLM layer is unavailable — the triage then falls
    back to the Table of Contents / deterministic rules on its own."""
    try:
        from data_extraction.ai_utils.llm_utils import (get_selected_model,
                                                        llm_call)
    except Exception as exc:  # noqa: BLE001
        logger.info(f"Triage runs without an LLM ({exc}).")
        return None

    def call(prompt, system_prompt):
        service = "o"
        try:
            model = get_selected_model(service)
        except Exception:  # noqa: BLE001
            model = None
        return llm_call(prompt, system_prompt, service, model)

    return call


def launch_review_app(chunks_data, logged_chunks, processed_indices, output_file,
                      logger, bulk=True):
    """
    Launch chunk review with automatic section review continuation.
    When chunk review completes, section review launches automatically.

    ``bulk`` (the default) opens the triage screen: the chunks are sorted
    automatically against the document's Table of Contents (or an LLM-checked
    heading list) and only the uncertain ones need attention. Pass
    ``bulk=False`` for the original one-chunk-at-a-time tool.
    """
    def on_chunk_review_complete():
        """Callback when chunk review finishes - auto-launch section review."""
        logger.info("Chunk review completed. Launching section review...")
        section_window = tk.Tk()
        section_window.title("Section Review")
        section_app = SectionReviewApp(
            root=section_window,
            output_file_path=output_file,
            logger=logger
        )
        section_window.mainloop()

    review_window = tk.Tk()
    if bulk:
        from .chunk_triage_ui import ChunkTriageApp
        app = ChunkTriageApp(
            root=review_window,
            chunks_data=chunks_data,
            output_file_name=output_file,
            logger=logger,
            on_complete_callback=on_chunk_review_complete,
            llm=_triage_llm(logger),
        )
    else:
        app = ChunkReviewApp(
            root=review_window,
            chunks_data=chunks_data,
            logged_chunks=logged_chunks,
            processed_indices=processed_indices,
            output_file_name=output_file,
            logger=logger,
            on_complete_callback=on_chunk_review_complete
        )
    review_window.mainloop()


def main():
    from ..crash_logging import install
    install()
    root = tk.Tk()
    launcher = ExtractionLauncherUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()