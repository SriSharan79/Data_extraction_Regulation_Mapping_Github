"""
EASA Studio — launcher for the EASA tools (extraction + JSON review).

Reuses the shared ``_BaseStudio`` base class and tab builders, so this file
only picks which two tabs to show.

Run from the repo root:
    python run_easa_studio.py
    # or: python -m data_extraction.studio.easa
"""

import tkinter as tk

from data_extraction.studio.base import _BaseStudio


class EASAStudio(_BaseStudio):
    """Launcher for the EASA extraction and structured-JSON review tabs."""

    WINDOW_TITLE = "EASA Studio — Extraction & Review"
    HEADER = "EASA XML extraction and structured-JSON review."
    GEOMETRY = "1280x820"
    TAB_SPECS = [
        ("EASA XML Extraction", "_build_easa_tab"),
        ("EASA JSON Review", "_build_easa_review_tab"),
    ]


def main():
    from ..crash_logging import install
    install()
    root = tk.Tk()
    EASAStudio(root)
    root.mainloop()


if __name__ == "__main__":
    main()
