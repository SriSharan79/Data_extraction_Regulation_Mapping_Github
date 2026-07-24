"""
Per-workspace SQLite data store.

One ``extraction_index.db`` lives inside each workspace folder and holds every
extracted document, its sections/nodes, the AI-review outputs, parsed entity
chains, and cross-document analysis results for that workspace.

Design (mirrors the proven patterns in the alr ``sql_store``):
  * Migrations only ever ADD columns (``PRAGMA table_info`` + ``ALTER TABLE``),
    so an older DB file keeps working after a schema grows.
  * ``upsert_document`` preserves enrichment columns with COALESCE, so a plain
    re-extract never wipes review/analysis data written later.
  * The DB path is injectable, so tests route it to a temp file instead of a
    real workspace.

No Tkinter, no third-party deps — pure ``sqlite3`` so it tests headless.
"""

from __future__ import annotations

import json
import os
import sqlite3
from datetime import datetime
from pathlib import Path

WORKSPACE_DB_NAME = "extraction_index.db"


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _dumps(value) -> str:
    """JSON-encode a value for a TEXT column (empty string for None)."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    try:
        return json.dumps(value, ensure_ascii=False)
    except (TypeError, ValueError):
        return str(value)


def _loads(text):
    """Best-effort JSON decode of a TEXT column; returns None for empty."""
    if text in (None, ""):
        return None
    try:
        return json.loads(text)
    except (TypeError, ValueError):
        return text


# --------------------------------------------------------------------------- #
# Schema — column lists are the single source of truth for each table.        #
# --------------------------------------------------------------------------- #
# documents: one row per extracted document (PDF or EASA).
DOCUMENT_COLUMNS = [
    "doc_uuid",          # stable id (hash of source path + type)
    "doc_name",
    "doc_type",          # 'pdf' | 'easa'
    "source_path",       # the PDF / zip / cache the doc came from
    "storage_root",      # the per-document storage folder
    "content_hash",
    "num_sections",
    "num_chunks",
    "num_images",
    "num_tables",
    "toc_source",        # where the TOC/headings were resolved from
    "status",            # 'extracted' | 'reviewed'
    "review_path",       # Processed_chunks*.json / structured JSON path
    "metadata_json",     # document_metadata blob (EASA) / free extras
    "created_at",
    "updated_at",
]
# Columns filled by later passes (review / analysis) — preserved on re-extract.
DOCUMENT_ENRICHMENT = {"status", "review_path", "num_sections", "metadata_json"}

# sections: one row per merged section (PDF) or hierarchy node (EASA).
SECTION_COLUMNS = [
    "doc_uuid",
    "ordinal",           # document order (0-based)
    "node_id",           # EASA node id, or synthesized for PDF
    "parent_node_id",
    "level",
    "heading",
    "text",
    "page",
    "toc_match",         # 'yes' | 'no' | ''
    "source",            # heading source / provenance
    "decision",          # 'log'/'skip'/'review' for PDF triage
    "labels_json",       # docitem labels / types
    "extra_json",
    "created_at",
    "updated_at",
]

# ai_runs: one row per AI-review or analysis invocation.
AI_RUN_COLUMNS = [
    "doc_uuid",          # NULL for cross-document runs
    "scope",             # 'document' | 'section' | 'cross-document'
    "mode",              # 'free-form' | 'column' | 'analysis'
    "service_requested",
    "model_requested",
    "prompt",
    "created_at",
]

# ai_reviews: one row per AI-reviewed section (or free-form result).
AI_REVIEW_COLUMNS = [
    "run_id",
    "doc_uuid",
    "section_ordinal",   # links to sections.ordinal when known (else NULL)
    "section_title",
    "column_values_json",  # {column: value} for column analysis
    "response_text",       # free-form response
    "service_used",
    "model_used",
    "fallback",            # 'yes' | ''
    "created_at",
]

# entities: parsed entity-chain components (Specific-entities column).
ENTITY_COLUMNS = [
    "doc_uuid",
    "run_id",
    "section_title",
    "reference",
    "system_info",
    "process",
    "personal",
    "physical_quantity",
    "quantity_value",
    "chain",
    "created_at",
]

# analysis_results: cross-document AI analysis outputs (the new capability).
ANALYSIS_COLUMNS = [
    "name",
    "scope_json",        # which docs/sections fed the analysis
    "prompt",
    "service_used",
    "model_used",
    "result_text",
    "result_json",
    "created_at",
]

_TABLES = {
    "documents": (DOCUMENT_COLUMNS, "doc_uuid"),
    "sections": (SECTION_COLUMNS, None),
    "ai_runs": (AI_RUN_COLUMNS, None),
    "ai_reviews": (AI_REVIEW_COLUMNS, None),
    "entities": (ENTITY_COLUMNS, None),
    "analysis_results": (ANALYSIS_COLUMNS, None),
}


class ExtractionStore:
    """Data-access object for one workspace's ``extraction_index.db``."""

    def __init__(self, db_path):
        self.db_path = str(db_path)
        Path(self.db_path).parent.mkdir(parents=True, exist_ok=True)
        self.init_db()

    # -- infrastructure ----------------------------------------------------- #
    def _connect(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def init_db(self):
        """Create tables if missing, and ADD any columns a newer schema needs
        (migrations only ever add — an older DB file keeps working)."""
        with self._connect() as conn:
            for table, (cols, pk) in _TABLES.items():
                defs = []
                if pk is None:
                    defs.append("id INTEGER PRIMARY KEY AUTOINCREMENT")
                for c in cols:
                    if c == pk:
                        defs.append(f"{c} TEXT PRIMARY KEY")
                    elif c in ("ordinal", "level", "run_id", "section_ordinal",
                               "num_sections", "num_chunks", "num_images",
                               "num_tables"):
                        defs.append(f"{c} INTEGER")
                    else:
                        defs.append(f"{c} TEXT")
                conn.execute(
                    f"CREATE TABLE IF NOT EXISTS {table} ({', '.join(defs)})")
                existing = {r[1] for r in
                            conn.execute(f"PRAGMA table_info({table})").fetchall()}
                for c in cols:
                    if c not in existing:
                        conn.execute(f"ALTER TABLE {table} ADD COLUMN {c} TEXT")
            # Helpful indexes for the list/browse views.
            conn.execute("CREATE INDEX IF NOT EXISTS ix_sections_doc "
                         "ON sections(doc_uuid, ordinal)")
            conn.execute("CREATE INDEX IF NOT EXISTS ix_reviews_doc "
                         "ON ai_reviews(doc_uuid)")
            conn.execute("CREATE INDEX IF NOT EXISTS ix_entities_doc "
                         "ON entities(doc_uuid)")

    # -- documents ---------------------------------------------------------- #
    def upsert_document(self, record: dict) -> str:
        """Insert or update a document by ``doc_uuid``. ``created_at`` is set
        once; enrichment columns (review status/path, section count, metadata)
        are preserved with COALESCE when the incoming record omits them, so a
        plain re-extract never wipes review data."""
        now = _now()
        data = {c: record.get(c) for c in DOCUMENT_COLUMNS}
        data["created_at"] = now
        data["updated_at"] = now
        if "metadata_json" in record and not isinstance(record["metadata_json"], str):
            data["metadata_json"] = _dumps(record["metadata_json"])

        placeholders = ", ".join("?" for _ in DOCUMENT_COLUMNS)
        col_list = ", ".join(DOCUMENT_COLUMNS)
        update_cols = [c for c in DOCUMENT_COLUMNS if c not in ("doc_uuid", "created_at")]
        update_sql = ", ".join(
            f"{c}=COALESCE(excluded.{c}, {c})" if c in DOCUMENT_ENRICHMENT
            else f"{c}=excluded.{c}"
            for c in update_cols)
        sql = (f"INSERT INTO documents ({col_list}) VALUES ({placeholders}) "
               f"ON CONFLICT(doc_uuid) DO UPDATE SET {update_sql}")
        with self._connect() as conn:
            conn.execute(sql, [data[c] for c in DOCUMENT_COLUMNS])
        return record.get("doc_uuid")

    def get_document(self, doc_uuid: str):
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM documents WHERE doc_uuid=?",
                               (doc_uuid,)).fetchone()
        return dict(row) if row else None

    def find_document_by_name(self, doc_name: str, doc_type: str = None):
        """The most-recently-updated document matching a name (used to attach
        AI reviews to a document that extraction already recorded, whose
        doc_uuid keys on a different source path than the review's JSON)."""
        if not doc_name:
            return None
        sql = "SELECT * FROM documents WHERE doc_name=?"
        args = [doc_name]
        if doc_type:
            sql += " AND doc_type=?"
            args.append(doc_type)
        sql += " ORDER BY COALESCE(updated_at, created_at) DESC LIMIT 1"
        with self._connect() as conn:
            row = conn.execute(sql, args).fetchone()
        return dict(row) if row else None

    def list_documents(self, doc_type: str = None):
        sql = "SELECT * FROM documents"
        args = []
        if doc_type:
            sql += " WHERE doc_type=?"
            args.append(doc_type)
        sql += " ORDER BY COALESCE(updated_at, created_at) DESC"
        with self._connect() as conn:
            return [dict(r) for r in conn.execute(sql, args).fetchall()]

    # -- sections ----------------------------------------------------------- #
    def replace_sections(self, doc_uuid: str, sections: list):
        """Replace all sections for a document (extraction is idempotent — a
        re-run rewrites the section set rather than appending duplicates)."""
        now = _now()
        with self._connect() as conn:
            conn.execute("DELETE FROM sections WHERE doc_uuid=?", (doc_uuid,))
            for ordinal, sec in enumerate(sections):
                rec = {c: sec.get(c) for c in SECTION_COLUMNS}
                rec["doc_uuid"] = doc_uuid
                rec["ordinal"] = sec.get("ordinal", ordinal)
                rec["labels_json"] = _dumps(sec.get("labels")
                                            if "labels_json" not in sec
                                            else sec.get("labels_json"))
                rec["extra_json"] = _dumps(sec.get("extra")
                                           if "extra_json" not in sec
                                           else sec.get("extra_json"))
                rec["created_at"] = now
                rec["updated_at"] = now
                cols = ", ".join(SECTION_COLUMNS)
                ph = ", ".join("?" for _ in SECTION_COLUMNS)
                conn.execute(f"INSERT INTO sections ({cols}) VALUES ({ph})",
                             [rec.get(c) for c in SECTION_COLUMNS])

    def get_sections(self, doc_uuid: str):
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM sections WHERE doc_uuid=? ORDER BY ordinal",
                (doc_uuid,)).fetchall()
        out = []
        for r in rows:
            d = dict(r)
            d["labels"] = _loads(d.get("labels_json"))
            d["extra"] = _loads(d.get("extra_json"))
            out.append(d)
        return out

    # -- AI runs / reviews / entities --------------------------------------- #
    def start_ai_run(self, doc_uuid=None, scope="document", mode="column",
                     service_requested=None, model_requested=None, prompt=None) -> int:
        rec = {
            "doc_uuid": doc_uuid, "scope": scope, "mode": mode,
            "service_requested": service_requested,
            "model_requested": model_requested, "prompt": prompt,
            "created_at": _now(),
        }
        cols = ", ".join(AI_RUN_COLUMNS)
        ph = ", ".join("?" for _ in AI_RUN_COLUMNS)
        with self._connect() as conn:
            cur = conn.execute(f"INSERT INTO ai_runs ({cols}) VALUES ({ph})",
                               [rec.get(c) for c in AI_RUN_COLUMNS])
            return cur.lastrowid

    def add_ai_review(self, run_id, doc_uuid=None, section_ordinal=None,
                      section_title=None, column_values=None, response_text=None,
                      service_used=None, model_used=None, fallback=None) -> int:
        rec = {
            "run_id": run_id, "doc_uuid": doc_uuid,
            "section_ordinal": section_ordinal, "section_title": section_title,
            "column_values_json": _dumps(column_values),
            "response_text": response_text or "",
            "service_used": service_used or "", "model_used": model_used or "",
            "fallback": fallback or "", "created_at": _now(),
        }
        cols = ", ".join(AI_REVIEW_COLUMNS)
        ph = ", ".join("?" for _ in AI_REVIEW_COLUMNS)
        with self._connect() as conn:
            cur = conn.execute(f"INSERT INTO ai_reviews ({cols}) VALUES ({ph})",
                               [rec.get(c) for c in AI_REVIEW_COLUMNS])
            return cur.lastrowid

    def add_entity(self, doc_uuid=None, run_id=None, section_title=None,
                   components: dict = None, chain=None) -> int:
        components = components or {}
        rec = {
            "doc_uuid": doc_uuid, "run_id": run_id, "section_title": section_title,
            "reference": components.get("reference") or components.get("Reference"),
            "system_info": components.get("system_info") or components.get("System Info"),
            "process": components.get("process") or components.get("Process"),
            "personal": components.get("personal") or components.get("Personal"),
            "physical_quantity": (components.get("physical_quantity")
                                  or components.get("Physical Quantity")),
            "quantity_value": (components.get("quantity_value")
                               or components.get("QuantityValue")),
            "chain": chain or "", "created_at": _now(),
        }
        cols = ", ".join(ENTITY_COLUMNS)
        ph = ", ".join("?" for _ in ENTITY_COLUMNS)
        with self._connect() as conn:
            cur = conn.execute(f"INSERT INTO entities ({cols}) VALUES ({ph})",
                               [rec.get(c) for c in ENTITY_COLUMNS])
            return cur.lastrowid

    def get_ai_reviews(self, doc_uuid: str):
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM ai_reviews WHERE doc_uuid=? ORDER BY id",
                (doc_uuid,)).fetchall()
        out = []
        for r in rows:
            d = dict(r)
            d["column_values"] = _loads(d.get("column_values_json"))
            out.append(d)
        return out

    def get_entities(self, doc_uuid: str):
        with self._connect() as conn:
            return [dict(r) for r in conn.execute(
                "SELECT * FROM entities WHERE doc_uuid=? ORDER BY id",
                (doc_uuid,)).fetchall()]

    # -- cross-document analysis -------------------------------------------- #
    def record_analysis(self, name, scope=None, prompt=None, service_used=None,
                        model_used=None, result_text=None, result=None) -> int:
        rec = {
            "name": name, "scope_json": _dumps(scope), "prompt": prompt,
            "service_used": service_used or "", "model_used": model_used or "",
            "result_text": result_text or "", "result_json": _dumps(result),
            "created_at": _now(),
        }
        cols = ", ".join(ANALYSIS_COLUMNS)
        ph = ", ".join("?" for _ in ANALYSIS_COLUMNS)
        with self._connect() as conn:
            cur = conn.execute(
                f"INSERT INTO analysis_results ({cols}) VALUES ({ph})",
                [rec.get(c) for c in ANALYSIS_COLUMNS])
            return cur.lastrowid

    def list_analyses(self):
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM analysis_results ORDER BY id DESC").fetchall()
        out = []
        for r in rows:
            d = dict(r)
            d["scope"] = _loads(d.get("scope_json"))
            d["result"] = _loads(d.get("result_json"))
            out.append(d)
        return out

    # -- corpus helpers for cross-document analysis ------------------------- #
    def iter_sections_across_documents(self, doc_type=None, limit_per_doc=None):
        """Yield ``(document_row, [section_rows])`` for every document — the
        corpus a cross-document AI analysis reads from."""
        for doc in self.list_documents(doc_type=doc_type):
            secs = self.get_sections(doc["doc_uuid"])
            if limit_per_doc:
                secs = secs[:limit_per_doc]
            yield doc, secs
