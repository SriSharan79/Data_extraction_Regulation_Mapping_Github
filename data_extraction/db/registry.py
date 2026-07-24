"""
Global workspace registry.

A small SQLite index — ``<home>/registry.db`` — that records which workspaces
exist and mirrors each workspace's document list, so the tool can show
everything across projects without opening every per-workspace database.

``home`` defaults to ``~/.data_extraction`` and can be redirected with the
``DATA_EXTRACTION_HOME`` environment variable (tests point it at a temp dir).
The registry holds only lightweight index rows; the heavy data stays in each
workspace's :class:`~data_extraction.db.store.ExtractionStore`.
"""

from __future__ import annotations

import os
import sqlite3
from datetime import datetime
from pathlib import Path

REGISTRY_DB_NAME = "registry.db"


def default_home() -> Path:
    """The global data-extraction home dir (``DATA_EXTRACTION_HOME`` or
    ``~/.data_extraction``)."""
    env = os.getenv("DATA_EXTRACTION_HOME")
    return Path(env) if env else Path.home() / ".data_extraction"


def registry_db_path() -> Path:
    return default_home() / REGISTRY_DB_NAME


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


WORKSPACE_COLUMNS = ["path", "db_path", "first_seen", "last_seen"]
DOCINDEX_COLUMNS = ["workspace_id", "doc_uuid", "doc_name", "doc_type",
                    "status", "updated_at"]


class WorkspaceRegistry:
    """The global index of workspaces and their documents."""

    def __init__(self, db_path=None):
        self.db_path = str(db_path or registry_db_path())
        Path(self.db_path).parent.mkdir(parents=True, exist_ok=True)
        self.init_db()

    def _connect(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def init_db(self):
        with self._connect() as conn:
            conn.execute(
                "CREATE TABLE IF NOT EXISTS workspaces ("
                "id INTEGER PRIMARY KEY AUTOINCREMENT, "
                "path TEXT UNIQUE, db_path TEXT, first_seen TEXT, last_seen TEXT)")
            conn.execute(
                "CREATE TABLE IF NOT EXISTS documents_index ("
                "id INTEGER PRIMARY KEY AUTOINCREMENT, "
                "workspace_id INTEGER, doc_uuid TEXT, doc_name TEXT, "
                "doc_type TEXT, status TEXT, updated_at TEXT, "
                "UNIQUE(workspace_id, doc_uuid))")
            # ADD-only migrations for either table.
            for table, cols in (("workspaces", WORKSPACE_COLUMNS),
                                ("documents_index", DOCINDEX_COLUMNS)):
                existing = {r[1] for r in
                            conn.execute(f"PRAGMA table_info({table})").fetchall()}
                for c in cols:
                    if c not in existing:
                        conn.execute(f"ALTER TABLE {table} ADD COLUMN {c} TEXT")

    # -- workspaces --------------------------------------------------------- #
    def register_workspace(self, path: str, db_path: str) -> int:
        """Record (or refresh) a workspace and return its registry id."""
        path = str(Path(path).resolve())
        now = _now()
        with self._connect() as conn:
            conn.execute(
                "INSERT INTO workspaces (path, db_path, first_seen, last_seen) "
                "VALUES (?, ?, ?, ?) "
                "ON CONFLICT(path) DO UPDATE SET db_path=excluded.db_path, "
                "last_seen=excluded.last_seen",
                (path, str(db_path), now, now))
            row = conn.execute("SELECT id FROM workspaces WHERE path=?",
                               (path,)).fetchone()
            return row["id"] if row else None

    def list_workspaces(self):
        with self._connect() as conn:
            return [dict(r) for r in conn.execute(
                "SELECT * FROM workspaces ORDER BY last_seen DESC").fetchall()]

    # -- document index ----------------------------------------------------- #
    def index_document(self, workspace_id: int, doc_uuid: str, doc_name: str,
                       doc_type: str, status: str):
        with self._connect() as conn:
            conn.execute(
                "INSERT INTO documents_index "
                "(workspace_id, doc_uuid, doc_name, doc_type, status, updated_at) "
                "VALUES (?, ?, ?, ?, ?, ?) "
                "ON CONFLICT(workspace_id, doc_uuid) DO UPDATE SET "
                "doc_name=excluded.doc_name, doc_type=excluded.doc_type, "
                "status=excluded.status, updated_at=excluded.updated_at",
                (workspace_id, doc_uuid, doc_name, doc_type, status, _now()))

    def list_all_documents(self):
        """Every indexed document across all workspaces (for a global browse)."""
        with self._connect() as conn:
            return [dict(r) for r in conn.execute(
                "SELECT d.*, w.path AS workspace_path, w.db_path AS workspace_db "
                "FROM documents_index d JOIN workspaces w ON w.id=d.workspace_id "
                "ORDER BY d.updated_at DESC").fetchall()]
