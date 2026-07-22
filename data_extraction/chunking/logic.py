import glob
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
    """PDF extraction + review in ONE page: pick the document and the storage
    destination, hit Run, and the pre-sorted chunk triage appears right below
    for review and saving — no separate review window.

    Resume/versioning: when the chosen document + storage match a previous
    run and a reviewed ``Processed_chunks*.json`` already exists anywhere
    under that document's storage root, its decisions are pre-applied to the
    triage; accepting with changes writes a NEW file (the earlier review is
    never overwritten), accepting without changes keeps the existing one."""

    def __init__(self, root):
        self.root = root
        self.root.title("PDF Extraction & Review")
        self.root.geometry("1200x800")

        self.triage = None            # the embedded ChunkTriageApp (if any)

        # -- Source & storage ------------------------------------------------
        frame_src = ttk.LabelFrame(
            root, text=" Document & storage ", padding=10)
        frame_src.pack(fill="x", padx=10, pady=(10, 4))

        ttk.Label(frame_src, text="PDF file:").grid(row=0, column=0, sticky="w")
        self.entry_pdf = ttk.Entry(frame_src, width=78)
        self.entry_pdf.grid(row=0, column=1, padx=5, pady=3, sticky="we")
        ttk.Button(frame_src, text="Browse...",
                   command=self.browse_pdf).grid(row=0, column=2)

        ttk.Label(frame_src, text="…or cache JSON:").grid(row=1, column=0, sticky="w")
        self.entry_cache = ttk.Entry(frame_src, width=78)
        self.entry_cache.grid(row=1, column=1, padx=5, pady=3, sticky="we")
        ttk.Button(frame_src, text="Browse...",
                   command=self.browse_cache).grid(row=1, column=2)

        ttk.Label(frame_src, text="Storage destination:").grid(row=2, column=0, sticky="w")
        self.entry_store = ttk.Entry(frame_src, width=78)
        self.entry_store.grid(row=2, column=1, padx=5, pady=3, sticky="we")
        ttk.Button(frame_src, text="Browse...",
                   command=self.browse_storage).grid(row=2, column=2)

        btn_box = ttk.Frame(frame_src)
        btn_box.grid(row=3, column=1, pady=8, sticky="w")
        self.btn_run = ttk.Button(btn_box, text="▶ Run Extraction & Review",
                                  command=self.run_extraction_review)
        self.btn_run.pack(side="left")
        self.btn_cache_only = ttk.Button(btn_box, text="Generate Cache Only",
                                         command=self.run_cache_only)
        self.btn_cache_only.pack(side="left", padx=5)
        ttk.Label(
            frame_src,
            text=("Give a cache JSON to skip re-extraction (its storage must be "
                  "the folder the cache was generated with); with only a PDF, "
                  "Docling converts it first.\nThe chunks then appear below, "
                  "pre-sorted against the document's Table of Contents (or a "
                  "validated heading list) — review, adjust, Accept & save."),
            foreground="#666666",
            justify="left",
        ).grid(row=4, column=1, sticky="w", pady=(2, 0))
        frame_src.columnconfigure(1, weight=1)

        # -- Embedded triage review ------------------------------------------
        frame_triage = ttk.LabelFrame(
            root, text=" Chunk triage & review ", padding=4)
        frame_triage.pack(fill="both", expand=True, padx=10, pady=(4, 4))
        self.triage_host = ttk.Frame(frame_triage)
        self.triage_host.pack(fill="both", expand=True)
        self._triage_placeholder = ttk.Label(
            self.triage_host, foreground="#666666", justify="left", padding=12,
            text="No document loaded yet.\n\nPick a PDF (or an existing cache "
                 "JSON) and the storage destination above, then Run — the "
                 "pre-sorted chunks appear here.\nIf this document was "
                 "reviewed before in that storage, the previous decisions are "
                 "loaded automatically; saving with changes creates a new "
                 "file, the earlier review stays untouched.")
        self._triage_placeholder.pack(anchor="nw")

        # Buttons disabled while a background conversion is running.
        self.action_buttons = [self.btn_run, self.btn_cache_only]
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
            self.auto_populate_storage(path, self.entry_store)

    def browse_cache(self):
        path = filedialog.askopenfilename(filetypes=[("JSON Cache Files", "*.json")])
        if path:
            self.entry_cache.delete(0, tk.END)
            self.entry_cache.insert(0, path)
            # Check for history and automatically fill out destination space
            self.auto_populate_storage(path, self.entry_store)

    def browse_storage(self):
        path = filedialog.askdirectory()
        if path:
            self.entry_store.delete(0, tk.END)
            self.entry_store.insert(0, path)

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
                    self.entry_store.delete(0, tk.END)
                    self.entry_store.insert(0, chosen_dir)
                    storage_path = chosen_dir

        self.save_to_registry(reference_path, storage_path)
        return storage_path

    def run_cache_only(self):
        pdf_path = self.entry_pdf.get().strip()
        storage_path = self.entry_store.get().strip()

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

    def _make_logger(self, paths):
        logger = logging.getLogger("ExtractionLauncher")
        logger.setLevel(logging.INFO)
        if not logger.handlers:
            fh = logging.FileHandler(paths["log_file"], encoding="utf-8")
            fh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            logger.addHandler(fh)
        return logger

    def run_extraction_review(self):
        """The ONE run button: extract (or load the cache), then show the
        pre-sorted triage right below for review and saving."""
        pdf_path = self.entry_pdf.get().strip()
        cache_path = self.entry_cache.get().strip()
        storage_path = self.entry_store.get().strip()

        if not storage_path:
            messagebox.showerror("Missing input", "Please provide a storage destination.")
            return

        # -- cache route: no conversion needed, show the triage immediately --
        if cache_path:
            if not os.path.exists(cache_path):
                messagebox.showerror("File not found", f"Cache file does not exist:\n{cache_path}")
                return
            storage_path = self.validate_and_confirm_storage(storage_path, cache_path)
            paths = self.resolve_directory_structure(storage_path, cache_path)
            os.makedirs(paths["dated_folder"], exist_ok=True)
            logger = self._make_logger(paths)
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
            self._open_triage(chunks_data, paths, logger)
            return

        # -- PDF route: Docling conversion on a background thread ------------
        if not pdf_path:
            messagebox.showerror("Missing input", "Please provide a PDF file (or a cache JSON).")
            return
        if not os.path.exists(pdf_path):
            messagebox.showerror("File not found", f"PDF file does not exist:\n{pdf_path}")
            return
        self.save_to_registry(pdf_path, storage_path)

        def job():
            paths = self.resolve_directory_structure(storage_path, pdf_path)
            os.makedirs(paths["dated_folder"], exist_ok=True)
            logger = self._make_logger(paths)
            chunks_data = generate_and_cache_document(pdf_path, paths, logger)
            print(Fore.GREEN + f"✓ Extraction complete. {len(chunks_data)} chunks extracted.")
            return (chunks_data, paths, logger)

        def on_done(result, error):
            if error:
                messagebox.showerror("Extraction Error", f"Extraction failed:\n{error}")
                return
            chunks_data, paths, logger = result
            if not chunks_data:
                messagebox.showerror("Data Error", "Extraction produced no chunks.")
                return
            self._open_triage(chunks_data, paths, logger)

        self._run_in_background(job, on_done, "Extracting document, please wait...")

    def _find_prior_review(self, paths):
        """The most recent reviewed Processed_chunks*.json anywhere under
        this document's storage root (reviews live in dated subfolders, so
        an earlier day's review is found too). Returns (path, history)."""
        pattern = os.path.join(paths["root"], "*", "Processed_chunks*.json")
        for candidate in sorted(glob.glob(pattern), key=os.path.getmtime,
                                reverse=True):
            try:
                with open(candidate, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                history = (data.get("raw_session_history")
                           if isinstance(data, dict) else data) or []
            except Exception:  # noqa: BLE001 - a broken file is just skipped
                continue
            if history:
                return candidate, history
        return None, []

    def _open_triage(self, chunks_data, paths, logger):
        """Build the embedded triage review below the inputs. A previous
        review of this document+storage (if any) is pre-applied; accepting
        with changes saves a NEW file, unchanged keeps the existing one."""
        from .chunk_triage_ui import ChunkTriageApp

        prior_path, prior_history = self._find_prior_review(paths)
        if prior_path:
            self.status_label.config(
                text=f"Previous review found — decisions loaded from "
                     f"{os.path.basename(os.path.dirname(prior_path))}/"
                     f"{os.path.basename(prior_path)}")
            logger.info(f"Prior review pre-applied from {prior_path}")
        else:
            self.status_label.config(text="No previous review for this "
                                          "document — fresh triage.")

        for child in self.triage_host.winfo_children():
            child.destroy()
        self.triage = ChunkTriageApp(
            root=self.triage_host,
            chunks_data=chunks_data,
            output_file_name=paths["output_file"],
            logger=logger,
            on_complete_callback=self._open_section_review,
            llm=_triage_llm(logger),
            prior_history=prior_history,
            prior_path=prior_path,
            version_on_change=True,
            destroy_on_accept=False,      # embedded: save and stay
        )

    def _open_section_review(self, saved_path=None):
        """Chain into Section Review (modal window) on the file the triage
        actually saved."""
        from .section_review_ui import SectionReviewApp

        path = saved_path or (self.triage.saved_path if self.triage else None)
        if not path or not os.path.exists(path):
            return
        top = self.root.winfo_toplevel()
        win = tk.Toplevel(top)
        win.transient(top)
        SectionReviewApp(root=win, output_file_path=path,
                         logger=logging.getLogger("ExtractionLauncher"))

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
            prior_history=logged_chunks,   # resume: earlier decisions pre-applied
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