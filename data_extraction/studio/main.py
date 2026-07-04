"""
Data Extraction Studio — main launcher (non-EASA tools).

Hosts the PDF extraction & review, cache review, section review, and
PDF -> Markdown tabs. The EASA tabs live in ``data_extraction.studio.easa``.

Run from the repo root:
    python run_studio.py
    # or: python -m data_extraction.studio.main
"""

import tkinter as tk

from data_extraction.studio.base import _BaseStudio


class DataExtractionStudio(_BaseStudio):
    """Launcher for the non-EASA tools."""

    WINDOW_TITLE = "Data Extraction Studio"
    HEADER = "Data Extraction Studio — pick a tab to begin."
    TAB_SPECS = [
        ("PDF Extraction & Review", "_build_extraction_tab"),
        ("Cache Review Launcher", "_build_cache_tab"),
        ("Section Review", "_build_section_review_tab"),
        ("PDF -> Markdown", "_build_markdown_tab"),
    ]


def main():
    root = tk.Tk()
    DataExtractionStudio(root)
    root.mainloop()


if __name__ == "__main__":
    main()
