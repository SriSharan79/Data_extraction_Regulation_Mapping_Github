"""
High-level persistence facade.

The one entry point the UI and the extraction / AI-review flows use. It bundles
a workspace's :class:`ExtractionStore` with its :class:`WorkspaceRegistry`
entry, and maps the tools' domain payloads (Docling ``merged_headings``, EASA
``rules_hierarchy``, AI-review row dicts, entity chains) into store rows.

Everything here is pure data-mapping — no Tkinter — so callers can persist from
a worker thread and it all tests headless. Callers should treat persistence as
best-effort: wrap calls so a DB error never blocks a review (see
:func:`safe`).
"""

from __future__ import annotations

import hashlib
import os
from pathlib import Path

from .store import ExtractionStore, WORKSPACE_DB_NAME
from .registry import WorkspaceRegistry


class Workspace:
    """A workspace's data store plus its global-registry identity."""

    def __init__(self, path, store: ExtractionStore, registry: WorkspaceRegistry,
                 workspace_id):
        self.path = str(path)
        self.store = store
        self.registry = registry
        self.workspace_id = workspace_id


def workspace_db_path(workspace_dir) -> str:
    return str(Path(workspace_dir) / WORKSPACE_DB_NAME)


def open_workspace(workspace_dir, registry: WorkspaceRegistry = None) -> Workspace:
    """Open (creating if needed) a workspace's data DB and register it in the
    global registry. ``registry`` is injectable for tests."""
    workspace_dir = str(Path(workspace_dir))
    db_path = workspace_db_path(workspace_dir)
    store = ExtractionStore(db_path)
    reg = registry or WorkspaceRegistry()
    ws_id = reg.register_workspace(workspace_dir, db_path)
    return Workspace(workspace_dir, store, reg, ws_id)


def doc_uuid_for(source_path: str, doc_type: str) -> str:
    """A stable id for a document from its source path + type."""
    key = f"{doc_type}::{os.path.abspath(source_path)}" if source_path else f"{doc_type}::"
    return hashlib.sha1(key.encode("utf-8")).hexdigest()[:16]


def safe(fn, logger=None):
    """Run a persistence call, swallowing any error (best-effort: a DB problem
    must never break a review). Returns the result or None."""
    try:
        return fn()
    except Exception as exc:  # noqa: BLE001 - persistence is best-effort
        if logger:
            try:
                logger.warning(f"DB persistence skipped: {exc}")
            except Exception:  # noqa: BLE001
                pass
        return None


# --------------------------------------------------------------------------- #
# PDF (Docling) side                                                          #
# --------------------------------------------------------------------------- #
def merged_headings_to_sections(merged_headings: list) -> list:
    """Map the triage ``merged_headings`` payload to section rows."""
    sections = []
    for i, sec in enumerate(merged_headings or []):
        heading = sec.get("heading") or []
        if isinstance(heading, list):
            heading_text = " / ".join(str(h) for h in heading if h)
        else:
            heading_text = str(heading)
        pages = sec.get("page_nums") or []
        sections.append({
            "ordinal": i,
            "node_id": f"sec_{i}",
            "parent_node_id": None,
            "level": 1,
            "heading": heading_text,
            "text": sec.get("merged_text", ""),
            "page": ", ".join(str(p) for p in pages) if pages else "",
            "toc_match": "",
            "source": "triage",
            "decision": "log",
            "labels_json": sec.get("types_of_docitem") or [],
            "extra_json": {"chunk_indices": sec.get("chunk_indices") or []},
        })
    return sections


def persist_pdf_review(ws: Workspace, *, source_path, storage_root, payload,
                       review_path=None, doc_name=None, num_chunks=None,
                       toc_source=None, logger=None) -> str:
    """Persist a reviewed PDF document + its merged sections. ``payload`` is
    the triage output ({merged_headings, raw_session_history})."""
    merged = (payload or {}).get("merged_headings", []) if isinstance(payload, dict) else []
    doc_uuid = doc_uuid_for(source_path, "pdf")
    name = doc_name or (Path(source_path).stem if source_path else "document")

    def _do():
        ws.store.upsert_document({
            "doc_uuid": doc_uuid,
            "doc_name": name,
            "doc_type": "pdf",
            "source_path": source_path,
            "storage_root": storage_root,
            "num_sections": len(merged),
            "num_chunks": num_chunks,
            "toc_source": toc_source,
            "status": "reviewed",
            "review_path": review_path,
        })
        ws.store.replace_sections(doc_uuid, merged_headings_to_sections(merged))
        ws.registry.index_document(ws.workspace_id, doc_uuid, name, "pdf", "reviewed")
        return doc_uuid

    return safe(_do, logger)


# --------------------------------------------------------------------------- #
# EASA (XML → structured JSON) side                                           #
# --------------------------------------------------------------------------- #
def _node_title(node: dict) -> str:
    for key in ("title", "heading", "name", "label"):
        v = node.get(key)
        if v:
            return str(v)
    return ""


def _node_text(node: dict) -> str:
    for key in ("text", "content", "body"):
        v = node.get(key)
        if isinstance(v, str) and v.strip():
            return v
    return ""


def easa_hierarchy_to_sections(hierarchy: list) -> list:
    """Flatten an EASA ``rules_hierarchy`` (nested nodes with ``children``)
    into ordered section rows, one per node, keeping parent/level."""
    sections = []
    counter = {"n": 0}

    def walk(node, parent_id, level):
        if not isinstance(node, dict):
            return
        idx = counter["n"]
        counter["n"] += 1
        node_id = str(node.get("node_id") or node.get("id") or f"node_{idx}")
        imgs = node.get("extracted_images") or node.get("images") or []
        tbls = node.get("extracted_tables") or node.get("tables") or []
        sections.append({
            "ordinal": idx,
            "node_id": node_id,
            "parent_node_id": parent_id,
            "level": level,
            "heading": _node_title(node),
            "text": _node_text(node),
            "page": str(node.get("page") or ""),
            "toc_match": "",
            "source": "easa-xml",
            "decision": "",
            "labels_json": node.get("type") or node.get("node_type") or "",
            "extra_json": {
                "attributes": node.get("attributes"),
                "hyperlinks": node.get("hyperlinks"),
                "images": imgs,
                "tables": tbls,
            },
        })
        for child in (node.get("children") or []):
            walk(child, node_id, level + 1)

    for top in (hierarchy or []):
        walk(top, None, 1)
    return sections


def persist_easa_document(ws: Workspace, *, source_path, storage_root,
                          structured_data, review_path=None, doc_name=None,
                          logger=None) -> str:
    """Persist an EASA document + its flattened hierarchy nodes.
    ``structured_data`` is the loaded structured JSON (dict with
    ``rules_hierarchy`` + ``document_metadata``) or a bare hierarchy list."""
    if isinstance(structured_data, dict):
        hierarchy = structured_data.get("rules_hierarchy") or []
        metadata = structured_data.get("document_metadata") or {}
    elif isinstance(structured_data, list):
        hierarchy, metadata = structured_data, {}
    else:
        hierarchy, metadata = [], {}

    doc_uuid = doc_uuid_for(source_path, "easa")
    name = doc_name or (metadata.get("title") if isinstance(metadata, dict) else None) \
        or (Path(source_path).stem if source_path else "document")
    sections = easa_hierarchy_to_sections(hierarchy)
    num_images = sum(len((s.get("extra_json") or {}).get("images") or [])
                     for s in sections)
    num_tables = sum(len((s.get("extra_json") or {}).get("tables") or [])
                     for s in sections)

    def _do():
        ws.store.upsert_document({
            "doc_uuid": doc_uuid,
            "doc_name": name,
            "doc_type": "easa",
            "source_path": source_path,
            "storage_root": storage_root,
            "num_sections": len(sections),
            "num_images": num_images,
            "num_tables": num_tables,
            "status": "extracted",
            "review_path": review_path,
            "metadata_json": metadata,
        })
        ws.store.replace_sections(doc_uuid, sections)
        ws.registry.index_document(ws.workspace_id, doc_uuid, name, "easa", "extracted")
        return doc_uuid

    return safe(_do, logger)


# --------------------------------------------------------------------------- #
# AI review side                                                              #
# --------------------------------------------------------------------------- #
def persist_ai_review_rows(ws: Workspace, *, doc_uuid, rows, mode="column",
                           service_requested=None, model_requested=None,
                           prompt=None, entity_column=None, logger=None):
    """Persist a batch of AI-review rows under a new run.

    ``rows`` are the review-panel row dicts. Each may carry provenance columns
    ('Service Used' / 'Model Used' / 'Fallback?') and, when ``entity_column``
    is given, that column's chains are parsed and stored in ``entities``.
    Returns the ``run_id`` (or None on error)."""
    def _do():
        run_id = ws.store.start_ai_run(
            doc_uuid=doc_uuid, scope="section" if doc_uuid else "document",
            mode=mode, service_requested=service_requested,
            model_requested=model_requested, prompt=prompt)
        for row in rows or []:
            title = row.get("Section") or row.get("title") or ""
            values = {k: v for k, v in row.items()
                      if k not in ("Service Used", "Model Used", "Fallback?")}
            ws.store.add_ai_review(
                run_id, doc_uuid=doc_uuid, section_title=title,
                column_values=values,
                response_text=row.get("response"),
                service_used=row.get("Service Used") or row.get("service_used"),
                model_used=row.get("Model Used") or row.get("model_used"),
                fallback=row.get("Fallback?") or row.get("fallback"))
            if entity_column and row.get(entity_column):
                _store_entities(ws, doc_uuid, run_id, title, row[entity_column])
        return run_id

    return safe(_do, logger)


def persist_ai_reviews_for(ws: Workspace, *, source_path=None, doc_name=None,
                           doc_type=None, rows, mode="column",
                           service_requested=None, model_requested=None,
                           prompt=None, entity_column=None, logger=None):
    """Persist AI-review rows, attaching them to the document they belong to.

    The AI-review UIs load a cache / structured / merged-headings JSON whose
    path differs from the source path extraction keyed the document on, so we
    resolve the document by *name* first (attaching to the extracted row when
    present) and only mint a lightweight document row when none exists yet.
    Returns the ``run_id`` (or None)."""
    def _do():
        doc = ws.store.find_document_by_name(doc_name) if doc_name else None
        if doc:
            doc_uuid = doc["doc_uuid"]
        else:
            dt = doc_type or ("easa" if str(source_path or "").lower().endswith(".zip")
                              else "pdf")
            doc_uuid = doc_uuid_for(source_path or doc_name or "", dt)
            ws.store.upsert_document({
                "doc_uuid": doc_uuid, "doc_name": doc_name or "document",
                "doc_type": dt, "source_path": source_path, "status": "extracted"})
            ws.registry.index_document(ws.workspace_id, doc_uuid,
                                       doc_name or "document", dt, "extracted")
        return persist_ai_review_rows(
            ws, doc_uuid=doc_uuid, rows=rows, mode=mode,
            service_requested=service_requested, model_requested=model_requested,
            prompt=prompt, entity_column=entity_column, logger=logger)

    return safe(_do, logger)


def _store_entities(ws, doc_uuid, run_id, title, cell):
    """Parse an entity-chain cell and store one row per chain (best-effort;
    uses entity_chains.parse_chains when available)."""
    try:
        from data_extraction.ai_utils.entity_chains import parse_chains
    except Exception:  # noqa: BLE001
        return
    # parse_chains yields component dicts with the Title-case COMPONENTS keys
    # ("Reference", "System Info", …, "QuantityValue") plus the raw "Chain";
    # add_entity() already accepts exactly those keys.
    for parsed in parse_chains(cell):
        ws.store.add_entity(doc_uuid=doc_uuid, run_id=run_id, section_title=title,
                            components=parsed, chain=parsed.get("Chain") or "")
