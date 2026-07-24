"""
Data Extraction Studio — main launcher (non-EASA tools).

Hosts the PDF extraction & review, section review, AI review, and
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
        ("Extract & Review", "_build_extract_review_tab"),
        ("PDF Extraction & Review", "_build_extraction_tab"),
        ("Section Review", "_build_section_review_tab"),
        ("AI Review", "_build_chunk_ai_tab"),
        ("Data & Analysis", "_build_data_analysis_tab"),
        ("PDF -> Markdown", "_build_markdown_tab"),
    ]


def main():
    from ..crash_logging import install
    install()
    root = tk.Tk()
    DataExtractionStudio(root)
    root.mainloop()


if __name__ == "__main__":
    main()
