"""
data_extraction.db
==================

A two-tier SQLite persistence layer for the extraction tools, kept alongside
the existing file-based outputs (JSON caches, Processed_chunks*.json, EASA
structured JSON, Excel indexes) rather than replacing them.

Tiers
-----
* **Per-workspace data DB** — ``<workspace>/extraction_index.db`` holds the
  heavy data: every extracted document, its sections/nodes, the AI-review
  outputs, parsed entity chains, and cross-document analysis results. Data
  travels with the regulatory dataset it describes.

* **Global registry DB** — ``<home>/registry.db`` (``~/.data_extraction`` by
  default, overridable with ``DATA_EXTRACTION_HOME``) is a small index of
  which workspaces exist and which documents live in each, so the tool can
  list everything across projects without opening every workspace.

Only the Python standard library ``sqlite3`` is used — no extra dependencies.
The stores are pure data-access objects (no Tkinter), so they test headless.

Typical use from the UI / pipelines goes through :mod:`data_extraction.db.facade`:

    from data_extraction.db import facade
    ws = facade.open_workspace("/path/to/workspace")   # store + registered
    facade.persist_pdf_review(ws, doc_meta, payload)
"""

from .store import ExtractionStore, WORKSPACE_DB_NAME
from .registry import WorkspaceRegistry, default_home, registry_db_path
from .analysis import (run_cross_document_analysis, build_corpus,
                       default_llm, available_llm)

__all__ = [
    "ExtractionStore",
    "WORKSPACE_DB_NAME",
    "WorkspaceRegistry",
    "default_home",
    "registry_db_path",
    "run_cross_document_analysis",
    "build_corpus",
    "default_llm",
    "available_llm",
]
