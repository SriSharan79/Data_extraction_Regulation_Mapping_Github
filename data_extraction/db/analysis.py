"""
Cross-document AI analysis over the workspace store.

Reads the sections (and, optionally, the AI-review outputs) of many stored
documents into one corpus, asks an AI model a question about that corpus, and
records the answer in ``analysis_results`` so it can be reviewed in the tool
and reused downstream.

The LLM is injected as a ``llm(prompt, system_prompt) -> str`` callable (the
same convention the triage uses), so this module is pure logic and tests
headless. :func:`default_llm` builds one from ``llm_utils.llm_call`` for real
runs, and :func:`available_llm` reports whether a usable one can be built.
"""

from __future__ import annotations

SYSTEM_PROMPT_CROSS_DOC = (
    "You are a regulatory-analysis assistant. You are given excerpts from "
    "several regulatory documents, each introduced by a '### DOCUMENT' header "
    "with its name. Answer the user's question using ONLY the provided "
    "excerpts, citing document names and section headings where relevant. If "
    "the excerpts do not contain the answer, say so plainly."
)

# Corpus size guards (characters) so a large workspace can't build a prompt
# bigger than the model will accept.
DEFAULT_MAX_CHARS_PER_SECTION = 1200
DEFAULT_MAX_SECTIONS_PER_DOC = 40
DEFAULT_MAX_TOTAL_CHARS = 60000


def build_corpus(store, doc_type=None, doc_uuids=None,
                 max_chars_per_section=DEFAULT_MAX_CHARS_PER_SECTION,
                 max_sections_per_doc=DEFAULT_MAX_SECTIONS_PER_DOC,
                 max_total_chars=DEFAULT_MAX_TOTAL_CHARS):
    """Build ``(corpus_text, scope)`` from the store's documents.

    ``scope`` records which documents/sections fed the corpus (persisted with
    the result). Sections and per-document counts are truncated to the guards
    so the prompt stays within a sane size."""
    parts = []
    scope = {"documents": [], "doc_type": doc_type, "total_sections": 0}
    total = 0
    wanted = set(doc_uuids) if doc_uuids else None

    for doc in store.list_documents(doc_type=doc_type):
        if wanted is not None and doc["doc_uuid"] not in wanted:
            continue
        secs = store.get_sections(doc["doc_uuid"])[:max_sections_per_doc]
        if not secs:
            continue
        header = f"\n### DOCUMENT: {doc.get('doc_name') or doc['doc_uuid']}\n"
        block = [header]
        used = 0
        for sec in secs:
            heading = (sec.get("heading") or "").strip()
            text = (sec.get("text") or "").strip()[:max_chars_per_section]
            line = f"- {heading}: {text}" if heading else f"- {text}"
            block.append(line)
            used += 1
            total += len(line)
            if total >= max_total_chars:
                break
        parts.append("\n".join(block))
        scope["documents"].append({
            "doc_uuid": doc["doc_uuid"],
            "doc_name": doc.get("doc_name"),
            "doc_type": doc.get("doc_type"),
            "sections_used": used,
        })
        scope["total_sections"] += used
        if total >= max_total_chars:
            scope["truncated"] = True
            break

    return "\n".join(parts), scope


def run_cross_document_analysis(store, prompt, llm, *, name=None, doc_type=None,
                                doc_uuids=None, system_prompt=None,
                                get_provenance=None, logger=None,
                                **corpus_kwargs):
    """Run one cross-document analysis and record it.

    ``store``   — an :class:`ExtractionStore`.
    ``prompt``  — the analyst's question.
    ``llm``     — ``llm(prompt, system_prompt) -> str`` callable (injected).
    ``get_provenance`` — optional ``() -> dict`` returning the service/model
                that actually answered (defaults to ``llm_utils.get_last_call_info``).

    Returns ``{"analysis_id", "name", "result_text", "scope", "service_used",
    "model_used"}`` or None when there was nothing to analyse / no LLM."""
    if llm is None:
        if logger:
            logger.info("Cross-document analysis skipped: no usable LLM.")
        return None

    corpus, scope = build_corpus(store, doc_type=doc_type, doc_uuids=doc_uuids,
                                 **corpus_kwargs)
    if not corpus.strip() or not scope["documents"]:
        if logger:
            logger.info("Cross-document analysis skipped: no stored sections.")
        return None

    full_prompt = (f"{prompt}\n\n=== CORPUS ({scope['total_sections']} sections "
                   f"from {len(scope['documents'])} document(s)) ===\n{corpus}")
    result_text = llm(full_prompt, system_prompt or SYSTEM_PROMPT_CROSS_DOC)
    if result_text is None:
        if logger:
            logger.info("Cross-document analysis produced no answer.")
        return None

    service_used = model_used = None
    prov = get_provenance
    if prov is None:
        try:
            from data_extraction.ai_utils.llm_utils import get_last_call_info as prov
        except Exception:  # noqa: BLE001
            prov = None
    if prov is not None:
        try:
            info = prov() or {}
            service_used = info.get("service_used")
            model_used = info.get("model_used")
        except Exception:  # noqa: BLE001
            pass

    name = name or "cross-document analysis"
    analysis_id = store.record_analysis(
        name=name, scope=scope, prompt=prompt, service_used=service_used,
        model_used=model_used, result_text=result_text,
        result={"answer": result_text})
    return {
        "analysis_id": analysis_id, "name": name, "result_text": result_text,
        "scope": scope, "service_used": service_used, "model_used": model_used,
    }


def available_llm() -> bool:
    """Whether a usable remote LLM can be built (a service key is stored and
    its model list answers). Cheap-ish network check — call off the UI thread."""
    try:
        from data_extraction.ai_utils.llm_utils import probe_available_services
        return bool(probe_available_services())
    except Exception:  # noqa: BLE001
        return False


def default_llm(service_code="b", model=None):
    """Build an ``llm(prompt, system_prompt)`` callable from ``llm_utils`` for a
    real run, or None when the LLM layer is unavailable. ``service_code`` is
    'b' (BlaBla), 'c' (Chat AI) or 'o' (DLR Ollama)."""
    try:
        from data_extraction.ai_utils.llm_utils import llm_call
    except Exception:  # noqa: BLE001
        return None

    def call(prompt, system_prompt):
        return llm_call(prompt, system_prompt, service_code, model)

    return call
