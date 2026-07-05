"""Modal Tk dialog to view, add, edit and remove the stored LLM API keys.

Open it from any UI with::

    from ..ai_utils.key_manager_ui import open_api_key_dialog
    open_api_key_dialog(parent_widget)

Saved keys go through :func:`LLM_Config.set_api_key`, so they are exported
to the current process environment *and* persisted to
``API_keys_config.json`` for the next launch.  Clearing a field removes the
key from both places.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox

from .LLM_Config import (API_keys_config, KEY_ENV_NAMES, delete_api_key,
                         get_stored_api_key, set_api_key)

# Friendlier row labels than the raw config names.
_SERVICE_LABELS = {
    'BlaBla Door': "Blablador (BlaBla Door)",
    'DLR Ollama': "DLR Ollama",
}


def open_api_key_dialog(parent):
    """Show the API-key manager as a modal dialog owned by *parent*."""
    win = tk.Toplevel(parent)
    win.title("API Keys")
    win.transient(parent.winfo_toplevel())
    win.resizable(False, False)

    frm = ttk.Frame(win, padding=12)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="Keys are used for AI Review calls, exported to the "
                        "environment for this session\nand saved to:\n"
                        f"{API_keys_config}").grid(
        row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

    entries: dict[str, ttk.Entry] = {}
    status: dict[str, ttk.Label] = {}

    def _refresh_status(service):
        stored = bool(get_stored_api_key(service))
        status[service].config(text="stored" if stored else "not set",
                               foreground="green" if stored else "#a00")

    for row, service in enumerate(KEY_ENV_NAMES, start=1):
        ttk.Label(frm, text=_SERVICE_LABELS.get(service, service) + ":").grid(
            row=row, column=0, sticky="w", pady=3)
        ent = ttk.Entry(frm, width=44, show="•")
        ent.insert(0, get_stored_api_key(service) or "")
        ent.grid(row=row, column=1, sticky="we", padx=6, pady=3)
        entries[service] = ent

        shown = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frm, text="Show", variable=shown,
            command=lambda e=ent, v=shown: e.config(show="" if v.get() else "•"),
        ).grid(row=row, column=2, padx=(0, 6))

        status[service] = ttk.Label(frm, width=8)
        status[service].grid(row=row, column=3, sticky="w")
        _refresh_status(service)

    def _save():
        changed = []
        for service, ent in entries.items():
            value = ent.get().strip()
            stored = get_stored_api_key(service) or ""
            if value and value != stored:
                set_api_key(service, value)
                changed.append(service)
            elif not value and stored:
                delete_api_key(service)
                changed.append(service)
            _refresh_status(service)
        if changed:
            messagebox.showinfo(
                "API Keys", "Updated: " + ", ".join(changed), parent=win)
        if win.winfo_exists():
            win.destroy()

    btns = ttk.Frame(frm)
    btns.grid(row=len(KEY_ENV_NAMES) + 1, column=0, columnspan=4,
              sticky="e", pady=(10, 0))
    ttk.Button(btns, text="Save", command=_save).pack(side="left", padx=4)
    ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left")

    win.grab_set()
    entries[next(iter(KEY_ENV_NAMES))].focus_set()
    return win
