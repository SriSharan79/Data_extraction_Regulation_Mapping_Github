import glob
import os
import json
import re
import logging
import subprocess
import sys
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
        ttk.Button(frame_src, text="Open PDF",
                   command=self.open_selected_pdf).grid(row=0, column=3,
                                                        padx=(4, 0))
        ttk.Button(frame_src, text="View TOC",
                   command=self.show_pdf_toc).grid(row=0, column=4,
                                                   padx=(2, 0))

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

        # LLM used by the triage as the LAST resort: it is asked to read the
        # Table of Contents out of the contents pages only when the PDF, the
        # chunks and the extracted tables all fail to give one, and to
        # validate the heading list when there is no TOC at all. Populated by
        # a background probe: a service appears ONLY when its API key is
        # stored AND its live model list answers — BlaBla first.
        ttk.Label(frame_src, text="LLM (TOC / heading check):").grid(row=3, column=0, sticky="w")
        llm_row = ttk.Frame(frame_src)
        llm_row.grid(row=3, column=1, pady=(2, 0), sticky="w")
        self.llm_service = ttk.Combobox(llm_row, state="readonly", width=12,
                                        values=[])
        self.llm_service.pack(side="left")
        self.llm_service.bind("<<ComboboxSelected>>", self._llm_service_picked)
        ttk.Label(llm_row, text="Model:").pack(side="left", padx=(8, 0))
        self.llm_model = ttk.Combobox(llm_row, state="readonly", width=42,
                                      values=[])
        self.llm_model.pack(side="left", padx=4)
        ttk.Button(llm_row, text="↻", width=3,
                   command=self._probe_llm_services).pack(side="left")
        self.llm_hint = ttk.Label(llm_row, foreground="#666666",
                                  text="checking services…")
        self.llm_hint.pack(side="left", padx=(8, 0))

        btn_box = ttk.Frame(frame_src)
        btn_box.grid(row=4, column=1, pady=8, sticky="w")
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
                  "pre-sorted against the document's Table of Contents — read "
                  "from the PDF, else the chunks, else a contents table, else "
                  "the LLM — or against a validated heading list.\nReview, "
                  "adjust, Accept & save."),
            foreground="#666666",
            justify="left",
        ).grid(row=5, column=1, sticky="w", pady=(2, 0))
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

        # Which LLM services are usable (probed in the background at start).
        self._llm_avail = {}
        try:
            self.root.after(300, self._probe_llm_services)
        except tk.TclError:
            pass

    # --- LLM service/model picker (for the triage heading check) ----------- #
    def _probe_llm_services(self):
        """Check in the background which services have a stored API key AND
        an answering model list; only those become pickable (BlaBla first)."""
        try:
            self.llm_hint.config(text="checking services…")
        except tk.TclError:
            pass

        def work():
            try:
                from data_extraction.ai_utils import llm_utils as _lu
                probe = getattr(_lu, "probe_available_services", None)
                avail = probe() if probe else []
            except Exception:  # noqa: BLE001 - offline just means none usable
                avail = []
            try:
                self.root.after(0, lambda a=avail: self._apply_llm_services(a))
            except (RuntimeError, tk.TclError):
                pass

        threading.Thread(target=work, daemon=True).start()

    def _apply_llm_services(self, avail):
        self._llm_avail = {label: models for label, models in (avail or [])}
        if not self._llm_avail:
            self.llm_service.configure(values=[])
            self.llm_service.set("")
            self.llm_model.configure(values=[])
            self.llm_model.set("")
            self.llm_hint.config(
                text="no usable LLM service (API key missing / models "
                     "unreachable) — triage uses the TOC / rules")
            return
        labels = list(self._llm_avail)          # probe returns BlaBla first
        self.llm_service.configure(values=labels)
        self.llm_service.set("BlaBla" if "BlaBla" in labels else labels[0])
        self.llm_hint.config(text="")
        self._llm_service_picked()

    def _llm_service_picked(self, _event=None):
        """Fill the model picker from the probed list of the chosen service."""
        label = self.llm_service.get().strip()
        models = self._llm_avail.get(label) or []
        self.llm_model.configure(values=models)
        current = ""
        try:
            from data_extraction.ai_utils.llm_utils import get_selected_model
            current = get_selected_model(label) or ""
        except Exception:  # noqa: BLE001
            pass
        self.llm_model.set(current if current in models
                           else (models[0] if models else ""))

    def _selected_triage_llm(self, logger):
        """The llm callable for the triage from the tab's pickers, or None
        when no usable service is selected (TOC/rules only)."""
        label = self.llm_service.get().strip()
        if not label or label not in self._llm_avail:
            return None
        code = {"DLR Ollama": "o", "Chat AI": "c"}.get(label, "b")
        model = self.llm_model.get().strip() or None
        return _triage_llm(logger, code, model)

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

    def open_selected_pdf(self):
        """Open the chosen PDF with the system's default viewer."""
        path = self.entry_pdf.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showinfo("No PDF", "Pick an existing PDF file first.")
            return
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", path])
            elif os.name == "nt":
                os.startfile(path)  # noqa: S606 - user-chosen file
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Open PDF", f"Could not open the PDF:\n{exc}")

    def _cache_payload_for_toc(self):
        """The cached extraction for the chosen document as
        ``(chunks, tables)``, or ``([], [])`` when there is none.

        The chunk-, table- and LLM-based TOC fallbacks all need the
        extraction; without a cache the TOC viewer can only read the PDF
        itself. Looked for in the cache JSON the user picked, then in the
        cache this document+storage would have written."""
        candidates = []
        cache_path = self.entry_cache.get().strip()
        if cache_path:
            candidates.append(cache_path)
        pdf_path = self.entry_pdf.get().strip()
        storage = self.entry_store.get().strip()
        if pdf_path and storage:
            try:
                candidates.append(self.resolve_directory_structure(
                    storage, pdf_path)["cache_file"])
            except Exception:  # noqa: BLE001 - a bad path is just no cache
                pass
        for path in candidates:
            if not path or not os.path.exists(path):
                continue
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    payload = json.load(f)
            except Exception:  # noqa: BLE001 - a broken cache is skipped
                continue
            if isinstance(payload, dict):
                return (payload.get("chunks") or [],
                        payload.get("tables") or [])
            if isinstance(payload, list):
                return payload, []
        return [], []

    def show_pdf_toc(self):
        """Show the Table of Contents the triage will actually use, resolved
        through the whole cascade: the PDF's embedded outline, the printed
        TOC in the PDF text, the TOC reassembled from the chunks, an
        extracted table sitting on the contents pages, and finally the LLM
        reading the contents pages.

        The last three need an extraction, so the cache (the picked JSON, or
        the one this document+storage wrote) is loaded when available.
        Resolved in the background — the LLM stage only runs if every earlier
        stage came back empty and a service is selected."""
        path = self.entry_pdf.get().strip()
        has_pdf = bool(path) and os.path.exists(path)
        chunks, tables = self._cache_payload_for_toc()
        if not has_pdf and not chunks:
            messagebox.showinfo(
                "Nothing to read",
                "Pick an existing PDF file — or a cache JSON from an earlier "
                "extraction — first.")
            return

        logger = logging.getLogger("ExtractionLauncher")
        llm = self._selected_triage_llm(logger)
        self.status_label.config(text="Resolving the Table of Contents…")

        def work():
            lines = []
            try:
                from .chunk_triage import resolve_toc
                toc = resolve_toc(pdf_path=path if has_pdf else None,
                                  chunks=chunks, tables=tables, llm=llm,
                                  log=lines.append)
            except Exception as exc:  # noqa: BLE001
                toc = {"entries": [], "source": None, "attempts": [],
                       "error": str(exc)}
            toc["log"] = lines
            toc["counts"] = {"chunks": len(chunks), "tables": len(tables),
                             "llm": bool(llm)}
            try:
                self.root.after(0, lambda t=toc: self._show_toc_dialog(path, t))
            except (RuntimeError, tk.TclError):
                pass

        threading.Thread(target=work, daemon=True).start()

    def _show_toc_dialog(self, path, toc):
        """The TOC viewer: which source won, what every stage produced (and
        why a contents table was rejected), then the entries themselves."""
        from .chunk_triage import TOC_SOURCE_LABEL

        self.status_label.config(text="")
        entries = toc.get("entries") or []
        source = toc.get("source")
        title = os.path.basename(path) if path else "the cached extraction"

        dlg = tk.Toplevel(self.root.winfo_toplevel())
        dlg.title(f"Table of Contents — {title}")
        dlg.transient(self.root.winfo_toplevel())
        dlg.geometry("760x640")
        frm = ttk.Frame(dlg, padding=10)
        frm.pack(fill="both", expand=True)

        if entries:
            headline = (f"{len(entries)} entries — from "
                        f"{TOC_SOURCE_LABEL.get(source, source or '?')}")
        else:
            headline = "No usable Table of Contents"
        ttk.Label(frm, font=("Arial", 10, "bold"), text=headline).pack(
            anchor="w")
        ttk.Label(frm, foreground="#666666", justify="left", wraplength=720,
                  text=("This is the section reference the triage verifies the "
                        "Docling chunk headings against. Without one, the "
                        "extracted heading list is validated instead (by the "
                        "LLM when a service is selected, else by rules).")
                  ).pack(anchor="w", pady=(2, 6))

        # -- what each stage of the cascade produced ------------------------
        attempts = toc.get("attempts") or []
        if attempts or toc.get("error"):
            box = ttk.LabelFrame(frm, text=" How this was resolved ",
                                 padding=6)
            box.pack(fill="x", pady=(0, 8))
            for att in attempts:
                if att.get("usable"):
                    mark, colour = "✓", "#1f7a1f"
                elif att.get("skipped"):
                    mark, colour = "–", "#888888"
                else:
                    mark, colour = "✗", "#b00020"
                detail = f"{att.get('entries', 0)} entries"
                if att.get("entries"):
                    detail += f" ({att.get('with_pages', 0)} with a page)"
                if att.get("note"):
                    detail += f" — {att['note']}"
                ttk.Label(box, foreground=colour, justify="left",
                          wraplength=700,
                          text=f"{mark}  {att.get('label', att.get('source'))}"
                               f": {detail}").pack(anchor="w")
            counts = toc.get("counts") or {}
            if counts:
                ttk.Label(box, foreground="#666666",
                          text=(f"(from {counts.get('chunks', 0)} cached "
                                f"chunks and {counts.get('tables', 0)} "
                                f"extracted tables; LLM "
                                f"{'selected' if counts.get('llm') else 'not selected'})")
                          ).pack(anchor="w", pady=(4, 0))
            if toc.get("error"):
                ttk.Label(box, foreground="#b00020", wraplength=700,
                          justify="left",
                          text=f"error: {toc['error']}").pack(anchor="w")

        # -- the entries ----------------------------------------------------
        wrap = ttk.Frame(frm)
        wrap.pack(fill="both", expand=True)
        box = tk.Text(wrap, wrap="none", font=("Consolas", 11))
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=box.yview)
        box.configure(yscrollcommand=vsb.set)
        box.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")
        if entries:
            for entry in entries:
                indent = "    " * max(0, int(entry.get("level") or 1) - 1)
                page = entry.get("page")
                stamp = (("p." + str(page)).rjust(7) if page is not None
                         else "      —")
                box.insert("end", f"{stamp}  {indent}{entry['title']}\n")
        else:
            box.insert("end",
                       "No Table of Contents could be resolved.\n\n"
                       "The lines above say what each stage found. The usual "
                       "reasons:\n"
                       "  • the PDF carries no bookmarks and no printed "
                       "contents page;\n"
                       "  • the contents pages have not been extracted yet — "
                       "run the extraction first;\n"
                       "  • the contents table did not extract cleanly (the "
                       "note above says so);\n"
                       "  • no LLM service is selected for the last-resort "
                       "pass.\n")
        box.configure(state="disabled")
        ttk.Button(frm, text="Close", command=dlg.destroy).pack(anchor="e",
                                                                pady=(6, 0))

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
                is_dict = isinstance(payload, dict)
                chunks_data = payload.get("chunks", []) if is_dict else payload
                # The extracted tables travel with the chunks: a contents
                # table is one of the sources the TOC is resolved from.
                tables_data = payload.get("tables", []) if is_dict else []
            except Exception as e:
                messagebox.showerror("Cache Parsing Error", f"Failed to parse cache file:\n{e}")
                return
            if not chunks_data:
                messagebox.showerror("Data Error", "No chunks found in cache file.")
                return
            self._open_triage(chunks_data, paths, logger, tables=tables_data)
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

    def _open_triage(self, chunks_data, paths, logger, tables=None):
        """Build the embedded triage review below the inputs. A previous
        review of this document+storage (if any) is pre-applied; accepting
        with changes saves a NEW file, unchanged keeps the existing one.

        ``tables`` are the extractor's table records; they are read back from
        the cache when not supplied, because a contents table is one of the
        sources the Table of Contents is resolved from."""
        from .chunk_triage_ui import ChunkTriageApp

        if tables is None:
            tables = load_cached_tables(paths.get("cache_file"))
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
            llm=self._selected_triage_llm(logger),
            prior_history=prior_history,
            prior_path=prior_path,
            version_on_change=True,
            destroy_on_accept=False,      # embedded: save and stay
            # with the PDF at hand the TOC is read from the PDF itself and
            # the chunk headings are verified against it; the tables let the
            # cascade fall back to a contents table when it is not
            pdf_path=self.entry_pdf.get().strip() or None,
            tables=tables,
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


def load_cached_tables(cache_file):
    """The extractor's table records from a cache file, or ``[]``.

    The triage needs them so a Table of Contents that was typeset as a table
    (and therefore extracted as a table rather than as text) can still be
    used as the section reference."""
    if not cache_file or not os.path.exists(cache_file):
        return []
    try:
        with open(cache_file, 'r', encoding='utf-8') as f:
            payload = json.load(f)
    except Exception:  # noqa: BLE001 - a broken cache just means no tables
        return []
    if isinstance(payload, dict):
        return payload.get("tables") or []
    return []


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


def _triage_llm(logger, service_code="b", model=None):
    """An ``llm(prompt, system_prompt)`` callable for the triage's heading
    check, or None when the LLM layer is unavailable — the triage then falls
    back to the Table of Contents / deterministic rules on its own.
    ``service_code`` is 'b' (BlaBla, the preferred default), 'c' (Chat AI)
    or 'o' (DLR Ollama); ``model`` overrides the service's selected model."""
    try:
        from data_extraction.ai_utils.llm_utils import (get_selected_model,
                                                        llm_call)
    except Exception as exc:  # noqa: BLE001
        logger.info(f"Triage runs without an LLM ({exc}).")
        return None

    label = {"o": "DLR Ollama", "c": "Chat AI"}.get(str(service_code).lower(), "BlaBla")

    def call(prompt, system_prompt):
        m = model
        if not m:
            try:
                m = get_selected_model(label)
            except Exception:  # noqa: BLE001
                m = None
        return llm_call(prompt, system_prompt, service_code, m)

    return call


def launch_review_app(chunks_data, logged_chunks, processed_indices, output_file,
                      logger, bulk=True, tables=None, pdf_path=None):
    """
    Launch chunk review with automatic section review continuation.
    When chunk review completes, section review launches automatically.

    ``bulk`` (the default) opens the triage screen: the chunks are sorted
    automatically against the document's Table of Contents (or an LLM-checked
    heading list) and only the uncertain ones need attention. Pass
    ``bulk=False`` for the original one-chunk-at-a-time tool.

    ``tables`` (the extractor's table records) and ``pdf_path`` widen the TOC
    cascade: with them the TOC can come from the PDF itself or from a
    contents table, not only from the chunks.
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
            pdf_path=pdf_path,
            tables=tables,                 # lets a contents TABLE serve as the TOC
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