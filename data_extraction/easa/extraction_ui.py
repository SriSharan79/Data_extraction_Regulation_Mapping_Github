"""
Standalone EASA XML Extraction UI.

A self-contained Tkinter window that extracts one EASA e-Rules XML ZIP (or a
whole folder of them) into a workspace: structured JSON, Master Excel index,
images and tables, and optionally the Cosmograph node/edge graph.

Run standalone:
    python EASA_Extraction_UI.py

It is also embeddable: pass a container Frame as ``root`` (the Data Extraction
Studio hosts it this way). The heavy backend runs on a worker thread so the
window stays responsive, and its stdout is streamed into the log pane.
"""

import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

_SENTINEL = object()


class _QueueStream:
    """stdout replacement that forwards writes to a queue for the UI to drain."""

    def __init__(self, q):
        self._q = q

    def write(self, s):
        if s:
            self._q.put(s)

    def flush(self):
        pass


class EASAExtractionApp:
    def __init__(self, root):
        self.root = root
        # Only touch window chrome on a real top-level window (standalone use);
        # when embedded in a Frame these are skipped.
        if isinstance(root, tk.Tk) or isinstance(root, tk.Toplevel):
            root.title("EASA XML Extraction")
            root.geometry("900x640")
            root.minsize(720, 480)

        self._q = None
        self._produced = []  # structured-JSON paths from the last run

        self._build_ui()

    # ---- UI ---------------------------------------------------------------- #
    def _build_ui(self):
        form = ttk.LabelFrame(
            self.root,
            text=" EASA XML ZIP  ->  Structured JSON / Master Excel / Images / Tables ",
            padding=10,
        )
        form.pack(fill="x", padx=10, pady=10)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Source (.zip file or a folder of .zip files):").grid(
            row=0, column=0, sticky="w"
        )
        self.ent_src = ttk.Entry(form)
        self.ent_src.grid(row=0, column=1, padx=5, pady=3, sticky="we")
        ttk.Button(form, text="File...", command=self._browse_file).grid(row=0, column=2)
        ttk.Button(form, text="Folder...", command=self._browse_folder).grid(row=0, column=3, padx=(2, 0))

        ttk.Label(form, text="Workspace / storage directory:").grid(row=1, column=0, sticky="w")
        self.ent_store = ttk.Entry(form)
        self.ent_store.grid(row=1, column=1, padx=5, pady=3, sticky="we")
        ttk.Button(form, text="Browse...", command=self._browse_store).grid(row=1, column=2)

        self.var_graph = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            form,
            text="Also build Cosmograph node/edge CSVs + Excel from each output JSON",
            variable=self.var_graph,
        ).grid(row=2, column=1, sticky="w", pady=(4, 0))

        btns = ttk.Frame(form)
        btns.grid(row=3, column=1, sticky="w", pady=8)
        self.run_btn = ttk.Button(btns, text="Run EASA Extraction", command=self._run)
        self.run_btn.pack(side="left")
        self.review_btn = ttk.Button(
            btns, text="Review output JSON…", command=self._open_review, state="disabled"
        )
        self.review_btn.pack(side="left", padx=8)

        ttk.Label(self.root, text="Log output:").pack(anchor="w", padx=10)
        self.txt = scrolledtext.ScrolledText(self.root, height=18, wrap="word")
        self.txt.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # ---- browse handlers --------------------------------------------------- #
    def _browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("ZIP archives", "*.zip")])
        if path:
            self.ent_src.delete(0, tk.END)
            self.ent_src.insert(0, path)

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Select a folder containing .zip files")
        if path:
            self.ent_src.delete(0, tk.END)
            self.ent_src.insert(0, path)

    def _browse_store(self):
        path = filedialog.askdirectory(title="Select a workspace directory")
        if path:
            self.ent_store.delete(0, tk.END)
            self.ent_store.insert(0, path)

    # ---- run --------------------------------------------------------------- #
    def _log(self, s):
        self.txt.insert("end", s)
        self.txt.see("end")

    def _run(self):
        src = self.ent_src.get().strip()
        store = self.ent_store.get().strip()
        if not src or not store:
            messagebox.showerror("Missing input", "Please provide both a source and a storage directory.")
            return
        if not os.path.exists(src):
            messagebox.showerror("Not found", f"Source path does not exist:\n{src}")
            return

        build_graph = self.var_graph.get()
        self._q = queue.Queue()
        self.run_btn.config(state="disabled")
        self.review_btn.config(state="disabled")
        self._produced = []

        def worker():
            old_stdout = sys.stdout
            sys.stdout = _QueueStream(self._q)
            produced = []
            try:
                from . import run_main
                produced = run_main.main(src, store, build_cosmograph=build_graph)
            except Exception as exc:  # noqa: BLE001 - surfaced to the log
                print(f"\n[ERROR] {exc}\n{traceback.format_exc()}\n")
            finally:
                sys.stdout = old_stdout
                self._q.put(("produced", produced))
                self._q.put(_SENTINEL)

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(100, self._poll)

    def _poll(self):
        try:
            while True:
                item = self._q.get_nowait()
                if item is _SENTINEL:
                    self.run_btn.config(state="normal")
                    if self._produced:
                        self.review_btn.config(state="normal")
                    return
                if isinstance(item, tuple) and item[0] == "produced":
                    self._produced = [p for p in (item[1] or []) if p]
                    continue
                self._log(item)
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    # ---- review handoff ---------------------------------------------------- #
    def _open_review(self):
        # Prefer an output that actually exists on disk.
        target = next((p for p in self._produced if p and os.path.exists(p)), None)
        if target is None:
            target = self._produced[0] if self._produced else None
        if target is None:
            messagebox.showinfo("Nothing to review", "No structured JSON output was produced yet.")
            return

        try:
            from .json_review_ui import EASAJsonReviewApp
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Review UI unavailable", f"Could not load the review UI:\n{exc}")
            return

        win = tk.Toplevel(self.root)
        EASAJsonReviewApp(win, json_path=target)


def main():
    from ..crash_logging import install
    install()
    root = tk.Tk()
    EASAExtractionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
