"""
alr.analysis_evaluation.metric_evaluator
========================================

Batch metric evaluation over a storage space -- the batch counterpart of the
manual two-text comparison panel, following the same space-wide pattern as
:mod:`data_evaluator` (which handles the substring/grounding check).

For every analyzed document it compares each extracted section item against the
document's identified reference text (abstract or introduction) using the
selected metric kinds:

* ``"lexical"``  -- Jaccard, ROUGE-1/2/L, BLEU (:mod:`Lexical_Overlap_Metrics`).
* ``"distance"`` -- Levenshtein distance/ratio + word error rate
  (``Distance_w_Structural _Alignment`` -- imported via importlib because of the
  space in the module filename).
* ``"cosine"``   -- embedding cosine similarity, reusing the same functions as
  :mod:`alr.rag_builders.query_executor`: existing per-section FAISS ``.bin``
  indexes are reused (vectors reconstructed by row), created from the section
  text-DB JSON when missing, and only when neither exists (e.g. the intro
  target, which has no vector DBs) is the item embedded directly.

Results are recorded in the storage space as **one dated workbook per metric
kind** (e.g. ``{date}_Abstract_Lexical_Metrics.xlsx`` /
``_Distance_Metrics.xlsx`` / ``_Cosine_Metrics.xlsx``) **plus a combined
overview workbook** holding all metric data together
(``{date}_Abstract_Metrics_Overview.xlsx`` — and the ``Introduction_*``
counterparts). Every workbook has one sheet per section plus an ``Overview``
sheet of per-document averages. The per-document averages are also merged into
the SQLite ``metrics_json`` column so the Review tool's overviews can use them.
Workbook locations come from :func:`alr.common.sections.build_metric_workbooks_map`.
"""

from __future__ import annotations

import importlib
from pathlib import Path

from colorama import Fore

from alr.common.file_manager import DataAnalyzeManager, Vec_DB_Manager
from alr.common.sections import (
    ALR_SECTIONS, INTRO_SECTIONS, ABSTRACT_TEXT_KEY, INTRO_TEXT_KEY,
    build_sections_map_vdb, build_metric_workbooks_map,
)
from alr.analysis_evaluation.data_evaluator import (
    _write_section_sheet_flat, _fetch_metadata,
    _load_abstract_json, _load_intro_json,
    _load_recorded_abstracts, _load_recorded_intros,
)

METRIC_KINDS = ("lexical", "distance", "cosine")

# Metric columns produced per kind (order preserved in the workbooks).
_KIND_COLUMNS = {
    "lexical": ("jaccard", "rouge1", "rouge2", "rougeL", "bleu"),
    "distance": ("levenshtein_distance", "similarity_ratio", "word_error_rate"),
    "cosine": ("cosine_similarity",),
}

_punkt_ready = None  # tri-state: None = unchecked, True/False = usable


def _ensure_punkt() -> bool:
    """BLEU needs nltk punkt data; try a one-time quiet download if missing."""
    global _punkt_ready
    if _punkt_ready is None:
        try:
            from nltk.tokenize import word_tokenize
            word_tokenize("probe sentence")
            _punkt_ready = True
        except Exception:
            try:
                import nltk
                nltk.download("punkt", quiet=True)
                nltk.download("punkt_tab", quiet=True)
                from nltk.tokenize import word_tokenize
                word_tokenize("probe sentence")
                _punkt_ready = True
            except Exception as e:
                print(Fore.YELLOW + f"⚠️ BLEU disabled (nltk punkt unavailable): {e}")
                _punkt_ready = False
    return _punkt_ready


def _lexical_metrics(reference, candidate) -> dict:
    """Jaccard/ROUGE/BLEU; each guarded so one failure records None, not an abort."""
    from alr.analysis_evaluation import Lexical_Overlap_Metrics as lom

    out = {c: None for c in _KIND_COLUMNS["lexical"]}
    try:
        out["jaccard"] = round(lom.calculate_jaccard_similarity(reference, candidate), 4)
    except Exception as e:
        print(Fore.YELLOW + f"⚠️ Jaccard failed: {e}")
    try:
        rouge = lom.calculate_rouge_scores(reference, candidate)
        out["rouge1"] = round(rouge.get("ROUGE-1", 0.0), 4)
        out["rouge2"] = round(rouge.get("ROUGE-2", 0.0), 4)
        out["rougeL"] = round(rouge.get("ROUGE-L", 0.0), 4)
    except Exception as e:
        print(Fore.YELLOW + f"⚠️ ROUGE failed: {e}")
    if _ensure_punkt():
        try:
            out["bleu"] = round(lom.calculate_bleu_score(reference, candidate), 4)
        except Exception as e:
            print(Fore.YELLOW + f"⚠️ BLEU failed: {e}")
    return out


def _distance_metrics(reference, candidate) -> dict:
    """Levenshtein + WER from the structural-alignment module (space in filename)."""
    out = {c: None for c in _KIND_COLUMNS["distance"]}
    try:
        mod = importlib.import_module("alr.analysis_evaluation.Distance_w_Structural _Alignment")
        res = mod.calculate_edit_distance_metrics(reference, candidate)
        out["levenshtein_distance"] = res["character_level"]["levenshtein_distance"]
        out["similarity_ratio"] = round(res["character_level"]["similarity_ratio"], 4)
        out["word_error_rate"] = round(res["word_level"]["word_error_rate"], 4)
    except Exception as e:
        print(Fore.YELLOW + f"⚠️ Distance metrics failed: {e}")
    return out


class _CosineContext:
    """
    Embedding cosine similarity against a per-document reference vector,
    reusing the storage space's per-section FAISS indexes where possible
    (same functions and index/metadata handling as query_executor).
    """

    def __init__(self, vdb, target):
        # {section_key: (bin_path, json_path)} -- abstract sections only; the
        # intro target has no vector DBs, so it always embeds directly.
        self.section_paths = build_sections_map_vdb(vdb) if target == "abstract" else {}
        self._cache = {}  # section_key -> (index, contents) or None

    @staticmethod
    def _normalize(vec):
        import numpy as np
        vec = np.asarray(vec, dtype="float32").reshape(-1)
        norm = float(np.linalg.norm(vec))
        return vec / norm if norm else vec

    def embed(self, text):
        from alr.rag_builders.vector_db_updater import vectorize_strings
        return self._normalize(vectorize_strings([str(text)])[0])

    def _section_index(self, section_key):
        """Load (index, contents) for a section; create the index if missing."""
        if section_key in self._cache:
            return self._cache[section_key]
        result = None
        paths = self.section_paths.get(section_key)
        if paths:
            bin_path, json_path = str(paths[0]), str(paths[1])
            try:
                from alr.common.json_utils import get_key_from_file
                from alr.rag_builders.vector_db_updater import (
                    load_index_file, vectorize_strings, create_faiss_index_cosine, save_index_file,
                )
                contents = get_key_from_file(json_path, "Content") if Path(json_path).exists() else None
                if contents:
                    index = load_index_file(bin_path)
                    if index is None:
                        # Vector DB missing -> create it from the text DB, then use it.
                        print(Fore.CYAN + f"🆕 Building vector index for '{section_key}'…")
                        embeds = vectorize_strings([str(c) for c in contents])
                        index = create_faiss_index_cosine(embeds)
                        save_index_file(index, bin_path)
                    if index is not None and getattr(index, "ntotal", 0):
                        result = (index, [str(c) for c in contents])
            except Exception as e:
                print(Fore.YELLOW + f"⚠️ Vector index unavailable for '{section_key}': {e}")
        self._cache[section_key] = result
        return result

    def item_vector(self, section_key, item_text):
        """Reconstruct the item's stored vector from the section index, else embed it."""
        entry = self._section_index(section_key)
        if entry:
            index, contents = entry
            try:
                pos = contents.index(str(item_text))
                if pos < index.ntotal:
                    return self._normalize(index.reconstruct(pos))
            except ValueError:
                pass  # item not in the text DB (e.g. never synced) -> embed directly
        return self.embed(item_text)

    def similarity(self, section_key, item_text, ref_vec):
        import numpy as np
        try:
            return round(float(np.dot(self.item_vector(section_key, item_text), ref_vec)), 4)
        except Exception as e:
            print(Fore.YELLOW + f"⚠️ Cosine failed for '{section_key}': {e}")
            return None


def _target_config(MF, VDB, target):
    """Section keys, JSON loader, reference-text key, workbooks and UUID source per target."""
    workbooks = {kind: Path(p) for kind, p in build_metric_workbooks_map(VDB, target).items()}
    if target == "intro":
        return {
            "section_keys": [key for key, _ in INTRO_SECTIONS],
            "loader": _load_intro_json,
            "text_key": INTRO_TEXT_KEY,
            "workbooks": workbooks,
            "recorded": _load_recorded_intros(MF),
            "sql_label": "introduction",
        }
    return {
        "section_keys": [spec.key for spec in ALR_SECTIONS],
        "loader": _load_abstract_json,
        "text_key": ABSTRACT_TEXT_KEY,
        "workbooks": workbooks,
        "recorded": _load_recorded_abstracts(MF),
        "sql_label": "abstract",
    }


def _metric_row(kinds, cosine_ctx, section_key, reference, candidate, ref_vec) -> dict:
    """Compute the selected metric columns for one item."""
    row = {}
    if "lexical" in kinds:
        row.update(_lexical_metrics(reference, candidate))
    if "distance" in kinds:
        row.update(_distance_metrics(reference, candidate))
    if "cosine" in kinds and cosine_ctx is not None and ref_vec is not None:
        row["cosine_similarity"] = cosine_ctx.similarity(section_key, candidate, ref_vec)
    return row


def _push_metrics_to_sql(uuid, averages, sql_label, db_path=None) -> bool:
    """Merge per-document metric averages into the metrics_json column (if row exists)."""
    import json
    from alr.common.sql_store import AnalyzedDataStore, DB_PATH

    store = AnalyzedDataStore(db_path or DB_PATH)
    doc = store.get_document(uuid)
    if not doc:
        return False
    try:
        merged = json.loads(doc.get("metrics_json") or "{}")
        if not isinstance(merged, dict):
            merged = {}
    except (ValueError, TypeError):
        merged = {}
    merged[sql_label] = averages
    store.update_document(uuid, {"metrics_json": json.dumps(merged)})
    return True


def _existing_metrics(uuid, metric_cols, sql_label, db_path=None):
    """
    Return the stored per-document metric averages for ``uuid`` if its
    ``metrics_json[sql_label]`` already covers every selected metric column,
    else ``None``. Used by ``mode="copy"`` (mirroring data_evaluator's
    ``_existing_evaluation``) to reuse prior metrics instead of recomputing.
    A document evaluated before with only SOME of the selected kinds (e.g.
    lexical done, cosine newly requested) does not count and is recomputed.
    """
    import json
    from alr.common.sql_store import AnalyzedDataStore, DB_PATH

    try:
        store = AnalyzedDataStore(db_path or DB_PATH)
        row = store.get_document(uuid)
    except Exception:
        return None
    if not row:
        return None
    try:
        merged = json.loads(row.get("metrics_json") or "{}")
    except (ValueError, TypeError):
        return None
    stored = merged.get(sql_label)
    if not isinstance(stored, dict):
        return None
    if all(col in stored for col in metric_cols):
        return stored
    return None


def evaluate_space_metrics(storage_path, kinds, target="abstract", db_path=None,
                           progress_callback=None, should_cancel=None, mode="generate") -> int:
    """
    Batch-compute the selected metric ``kinds`` (subset of :data:`METRIC_KINDS`)
    for every analyzed document in a storage space, against the ``target`` data
    ("abstract" or "intro"). Returns the number of documents evaluated.

    ``mode="copy"`` reuses prior metrics: documents whose SQL ``metrics_json``
    already covers every selected metric column are skipped (only new/partially
    evaluated documents are computed). ``mode="generate"`` (default, previous
    behaviour) recomputes every document and updates the workbook rows in place.

    Results go to the storage space as one dated workbook **per metric kind**
    plus a **combined overview workbook** with all metric data (each workbook:
    one sheet per section with per-item metric columns, plus an ``Overview``
    sheet of per-document averages), and into the SQLite ``metrics_json``
    column. Re-runs update rows in place (no duplicates).
    ``progress_callback(done, total)`` / ``should_cancel`` support progress
    reporting and cancellation.
    """
    kinds = {k for k in kinds if k in METRIC_KINDS}
    if not kinds:
        print("No metric kinds selected; nothing to do.")
        return 0

    MF = storage_path if isinstance(storage_path, DataAnalyzeManager) else DataAnalyzeManager(storage_path)
    VDB = Vec_DB_Manager(MF.folder)
    cfg = _target_config(MF, VDB, target)
    if not cfg["recorded"]:
        print(f"No recorded {target} analyses found; nothing to evaluate.")
        return 0

    cosine_ctx = _CosineContext(VDB, target) if "cosine" in kinds else None
    metric_cols = [c for k in ("lexical", "distance", "cosine") if k in kinds for c in _KIND_COLUMNS[k]]

    count = 0
    total = len(cfg["recorded"])
    for i, uuid in enumerate(cfg["recorded"], 1):
        if should_cancel is not None and should_cancel():
            print("Metric evaluation cancelled by user.")
            break
        try:
            if mode == "copy":
                prior = _existing_metrics(uuid, metric_cols, cfg["sql_label"], db_path)
                if prior is not None:
                    print(Fore.YELLOW + f"⏭️ Reusing existing {target} metrics for {uuid} (copy mode).")
                    count += 1
                    if progress_callback:
                        progress_callback(i, total)
                    continue

            MF.update_id_files(uuid)
            title, file_name = _fetch_metadata(MF, uuid)
            json_data = cfg["loader"](MF, uuid)
            if not json_data:
                continue
            reference = str(json_data.get(cfg["text_key"], "") or "")
            if not reference.strip():
                continue

            ref_vec = None
            if cosine_ctx is not None:
                try:
                    ref_vec = cosine_ctx.embed(reference)
                except Exception as e:
                    print(Fore.YELLOW + f"⚠️ Reference embedding failed for {uuid}: {e}")

            # Per-metric running sums for the document-level averages.
            sums = {c: [0.0, 0] for c in metric_cols}  # col -> [total, n]
            workbooks = cfg["workbooks"]  # {kind: per-kind workbook, "overview": combined}
            base = {"UUID": str(uuid), "Title": title, "Filename": file_name}

            for key in cfg["section_keys"]:
                value = json_data.get(key, None)
                if value is None or value == "Not Found":
                    continue
                items = [str(v) for v in value] if isinstance(value, list) else [str(value)]

                # One row per document: a per-kind entry for each metric's own
                # workbook, and a combined entry carrying all metric data.
                kind_entries = {k: dict(base) for k in kinds}
                combined_entry = dict(base)
                for n, item in enumerate(items, 1):
                    metrics = _metric_row(kinds, cosine_ctx, key, reference, item, ref_vec)
                    prefix = f"Item {n} " if len(items) > 1 else ""
                    combined_entry[f"{prefix}Content"] = item
                    for kind in kinds:
                        kind_entries[kind][f"{prefix}Content"] = item
                        for col in _KIND_COLUMNS[kind]:
                            val = metrics.get(col)
                            kind_entries[kind][f"{prefix}{col}"] = val
                            combined_entry[f"{prefix}{col}"] = val
                            if isinstance(val, (int, float)):
                                sums[col][0] += float(val)
                                sums[col][1] += 1
                for kind in kinds:
                    _write_section_sheet_flat(workbooks[kind], key, kind_entries[kind])
                _write_section_sheet_flat(workbooks["overview"], key, combined_entry)

            averages = {col: (round(t / n, 4) if n else None) for col, (t, n) in sums.items()}
            # Per-kind Overview sheets carry that kind's averages; the combined
            # overview workbook's Overview sheet carries all of them.
            for kind in kinds:
                kind_avgs = {f"avg_{c}": averages.get(c) for c in _KIND_COLUMNS[kind]}
                _write_section_sheet_flat(workbooks[kind], "Overview", {**base, **kind_avgs})
            _write_section_sheet_flat(workbooks["overview"], "Overview",
                                      {**base, **{f"avg_{c}": v for c, v in averages.items()}})

            try:
                _push_metrics_to_sql(uuid, averages, cfg["sql_label"], db_path)
            except Exception as e:
                print(Fore.YELLOW + f"⚠️ Could not push metrics to SQL for {uuid}: {e}")

            count += 1
            print(Fore.GREEN + f"✅ Metrics ({', '.join(sorted(kinds))}) recorded for {uuid} ({target}).")
        except Exception as e:
            print(Fore.YELLOW + f"⚠️ Metric evaluation failed for {uuid}: {e}")
        if progress_callback:
            progress_callback(i, total)

    written = [str(cfg["workbooks"][k]) for k in sorted(kinds)] + [str(cfg["workbooks"]["overview"])]
    print(Fore.GREEN + f"✅ Metric evaluation finished: {count} document(s) ({target}).")
    for path in written:
        print(Fore.GREEN + f"   📄 {path}")
    return count