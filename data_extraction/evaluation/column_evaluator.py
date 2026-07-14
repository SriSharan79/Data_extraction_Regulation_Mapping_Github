"""
data_extraction.evaluation.column_evaluator
===========================================

Evaluate AI Review **column analysis** output workbooks against the section
content that generated them — the adaptation of the alr evaluation tools in
this folder (``data_evaluator.py``'s substring/grounding check and
``metric_evaluator.py``'s guarded lexical/distance metrics) to this repo's
data.

Input is the accumulating ``.xlsx`` workbook the AI Review tab writes
(``Run N <timestamp>`` snapshot sheets, one row per section: ``Section`` +
one cell per extracted column) plus a reference map
``{section title: section text}`` — the exact text that was sent to the LLM.

Every evaluation is individually selectable via ``metrics``:

* ``grounding`` — each cell is split into its items and each item checked as
  a case-insensitive substring of the section's reference text
  (data_evaluator's ``Is_Subset``), giving per-column True/False counts and a
  per-section ``Coverage %``.
* ``jaccard`` — token-set overlap (no third-party library needed).
* ``rouge1`` / ``rouge2`` / ``rougeL`` and ``bleu`` — via
  ``Lexical_Overlap_Metrics`` when nltk/rouge_score are installed.
* ``levenshtein_distance`` / ``similarity_ratio`` / ``word_error_rate`` — via
  ``Distance_w_Structural _Alignment`` when Levenshtein/jiwer are installed
  (a stdlib ``difflib`` ratio fills ``similarity_ratio`` otherwise).

For the **per-item** metric sheets both sides are split: each cell into its
items (``; ``) and each reference section into sentences (sentence-aware,
protecting decimals / abbreviations / citations). Every item is scored
against every reference sentence and only the **most relevant** sentence's
value is kept — the highest for the similarity-style metrics, the lowest for
the edit-distance / error-rate metrics — alongside the sentence that produced
it. The **overview** (``Eval <run sheet>``) is unchanged: it still scores each
whole cell against the whole reference.

Each metric is guarded: a missing library or failure records ``None``
instead of aborting. The **overview** (``Eval <run sheet>``) keeps the
established shape: grounding counts per column plus every other metric
computed on the whole cell text, closed by an average row. In addition,
every cell is split into its items (:func:`split_items`) and **every
selected metric is evaluated per item**; each metric gets its **own sheet**
(``<code> <run sheet>``, e.g. ``Grd …``, ``Jac …`` — see
:data:`METRIC_SHEET_CODES`) with one row per item carrying the section
name, the source column, the item content and the item's metric value,
closed by an average row (coverage for grounding). All sheets are rebuilt
on every call — either in the **same workbook** or, via ``out_path``, in a
**new evaluation workbook** (which then also gets a copy of the run sheet
so it is self-contained).

The unique-value evaluation reuses
``scripts/excel_file_utils.save_unique_elements_to_new_sheet``: the
``Uniq <run sheet>`` sheet is refreshed to identify the unique values of the
generated content (element + count), then every unique element is grounded
against the combined reference text of the run's sections into a
``Uniq Eval <run sheet>`` sheet with a per-column coverage summary.

Standalone (references come from the sections JSON the run was made from):

    python -m data_extraction.evaluation.column_evaluator results.xlsx sections.json

where ``sections.json`` is a chunks cache (``*_docling_chunks_cache.json``)
or a ``Processed_chunks.json`` review output. Inside the app the AI Review
tab's *Evaluation* page drives the same functions (and can use the last
run's in-memory references, which also covers the EASA studio's sections).
"""

from __future__ import annotations

import argparse
import importlib
import importlib.util
import json
import os
import re
import sys
from pathlib import Path

# Every selectable evaluation, in output order.
ALL_METRICS = (
    "grounding",
    "jaccard", "rouge1", "rouge2", "rougeL", "bleu",
    "levenshtein_distance", "similarity_ratio", "word_error_rate",
    # Semantic-embedding metrics (see the embedding section below): these need a
    # backend + service and are configured once per run via configure_embeddings.
    "embedding_cosine", "bertscore_p", "bertscore_r", "bertscore_f1",
)
_LEXICAL = ("jaccard", "rouge1", "rouge2", "rougeL", "bleu")
_DISTANCE = ("levenshtein_distance", "similarity_ratio", "word_error_rate")
_EMBEDDING = ("embedding_cosine", "bertscore_p", "bertscore_r", "bertscore_f1")

# Short prefixes for the per-metric item sheets (Excel's 31-char sheet-name
# limit leaves no room for the full metric name next to the run-sheet name).
METRIC_SHEET_CODES = {
    "grounding": "Grd",
    "jaccard": "Jac",
    "rouge1": "R1",
    "rouge2": "R2",
    "rougeL": "RL",
    "bleu": "Bleu",
    "levenshtein_distance": "Lev",
    "similarity_ratio": "Sim",
    "word_error_rate": "WER",
    "embedding_cosine": "Cos",
    "bertscore_p": "BsP",
    "bertscore_r": "BsR",
    "bertscore_f1": "BsF",
}

# Candidate item separators inside one extracted cell (the AI Review tab
# joins list answers with "; "), mirroring excel_file_utils' detection.
_SEPARATORS = (
    ";", 
    # ","
    )

# Reference text is split into sentences so each candidate item can be scored
# against the single most relevant sentence rather than the whole section.
# A naive ". " split mangles decimals ("2.5"), abbreviations ("e.g.", "No.",
# "para.") and numbered citations; the splitter below protects those.
#
# A boundary candidate is sentence punctuation (with an optional closing
# quote/bracket) followed by whitespace; :func:`_is_sentence_end` then vetoes
# false boundaries.
_SENT_BOUNDARY = re.compile(r'([.!?]["\'”’)\]]?)(\s+)')

# Lowercased tokens (dotted forms kept) that should NOT end a sentence when
# followed by whitespace. Tuned for regulatory / technical text.
_ABBREVIATIONS = frozenset({
    "e.g", "i.e", "eg", "ie", "etc", "vs", "viz", "cf", "al",
    "no", "nos", "fig", "figs", "eq", "eqs", "ref", "refs",
    "sec", "secs", "art", "arts", "para", "paras", "reg", "regs",
    "ch", "chap", "pt", "pts", "pp", "vol", "vols", "ed", "eds",
    "approx", "ca", "incl", "excl", "min", "max", "avg", "std", "est",
    "mr", "mrs", "ms", "dr", "prof", "sr", "jr", "st", "rev",
    "inc", "ltd", "co", "corp", "dept", "univ", "govt", "assn",
    "u.s", "u.k", "e.u", "u.n", "no.s",
})

# When picking the most relevant sentence per metric: similarity-style scores
# are best when high; edit-distance / error-rate scores are best when low.
_LOWER_IS_BETTER = ("levenshtein_distance", "word_error_rate")

_EXCEL_UTILS = None


def _excel_utils():
    """Load scripts/excel_file_utils.py (repo root) once, by file path."""
    global _EXCEL_UTILS
    if _EXCEL_UTILS is None:
        p = Path(__file__).resolve().parents[2] / "scripts" / "excel_file_utils.py"
        spec = importlib.util.spec_from_file_location("excel_file_utils", str(p))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        _EXCEL_UTILS = mod
    return _EXCEL_UTILS


# ------------------------------------------------------------- grounding -- #

def split_items(cell):
    """Split one extracted cell into its items with the same dynamic
    separator detection as save_unique_elements_to_new_sheet (extended
    with ';', which the AI Review tab uses to join list answers)."""
    if cell is None or cell != cell:  # None or NaN (empty Excel cell)
        return []
    text = str(cell).strip()
    if not text:
        return []
    sep_counts = {s: text.count(s) for s in _SEPARATORS}
    best = max(sep_counts, key=sep_counts.get)
    if sep_counts[best] > 0:
        return [el.strip() for el in text.split(best) if el.strip()]
    return [text]


def is_grounded(item, reference):
    """data_evaluator's Is_Subset check: the extracted item is contained in
    the reference text (case-insensitive substring)."""
    return bool(reference) and str(item).lower() in str(reference).lower()


def _is_sentence_end(text, dot, after):
    """Decide whether the punctuation at index ``dot`` (with the next token
    starting at ``after``) is a real sentence boundary. Vetoes the common
    false positives: a digit right before the dot (decimals like ``2.5`` and
    numbered citations like ``para. 3.``), a known abbreviation or a
    single-uppercase-letter initial right before it, and a next token that
    starts lowercase (so an unlisted abbreviation followed by a lowercase word
    is not split). The trade-off is conservative — when unsure it keeps the
    text together rather than splitting mid-sentence."""
    # Decimal / numbered marker: a digit immediately precedes the dot.
    if dot > 0 and text[dot - 1].isdigit():
        return False
    # Abbreviation or single-letter initial immediately before the dot
    # (dotted acronyms like "e.g"/"U.S" are captured whole).
    m = re.search(r'([A-Za-z]+(?:\.[A-Za-z]+)*)$', text[:dot])
    if m:
        tok = m.group(1)
        if tok.lower() in _ABBREVIATIONS:
            return False
        if len(tok) == 1 and tok.isupper():
            return False
    # The next token should look like a sentence start, not a lowercase
    # continuation of an unlisted abbreviation.
    nxt = text[after] if after < len(text) else ""
    if nxt and nxt.islower():
        return False
    return True


def split_reference(reference):
    """Split a section's reference text into sentences so each candidate item
    can be scored against the single most relevant sentence instead of the
    whole section. Splits on sentence punctuation (``.``/``!``/``?``) followed
    by whitespace, but protects decimals (``2.5``), abbreviations (``e.g.``,
    ``No.``, ``para.``, ``Fig.``), numbered citations and single-letter
    initials via :func:`_is_sentence_end`. Returns the whole text as one
    sentence when no genuine boundary is found."""
    text = str(reference or "").strip()
    if not text:
        return []
    sentences, start = [], 0
    for m in _SENT_BOUNDARY.finditer(text):
        if not _is_sentence_end(text, m.start(1), m.end()):
            continue
        chunk = text[start:m.end(1)].strip()
        if chunk:
            sentences.append(chunk)
        start = m.end()
    tail = text[start:].strip()
    if tail:
        sentences.append(tail)
    return sentences or [text]


def _better(metric, new, current):
    """True if ``new`` is the more relevant score than ``current`` for
    ``metric``: lower for the edit-distance / error-rate metrics
    (``_LOWER_IS_BETTER``), higher for every similarity-style metric.
    ``current is None`` (nothing chosen yet) always accepts ``new``."""
    if current is None:
        return True
    return new < current if metric in _LOWER_IS_BETTER else new > current


def _lookup_reference(references, title):
    """Reference text for a run-sheet Section title (tolerant of stray
    whitespace/case differences introduced by the Excel round-trip)."""
    if title in references:
        return str(references[title] or "")
    key = str(title).strip().lower()
    for t, text in references.items():
        if str(t).strip().lower() == key:
            return str(text or "")
    return ""


# --------------------------------------------------------------- metrics -- #

_LOM = "unchecked"       # Lexical_Overlap_Metrics module, None if unusable
_DIST_MOD = "unchecked"  # Distance_w_Structural _Alignment module, None if unusable
_punkt_ready = None      # tri-state: None = unchecked, True/False = usable


def _lom():
    global _LOM
    if _LOM == "unchecked":
        try:
            from data_extraction.evaluation import Lexical_Overlap_Metrics as mod
            _LOM = mod
        except Exception as e:  # noqa: BLE001 - nltk/rouge_score not installed
            print(f"⚠️ ROUGE/BLEU disabled (Lexical_Overlap_Metrics unavailable: {e})")
            _LOM = None
    return _LOM


def _distance_mod():
    global _DIST_MOD
    if _DIST_MOD == "unchecked":
        try:
            # importlib because of the space in the module filename.
            _DIST_MOD = importlib.import_module(
                "data_extraction.evaluation.Distance_w_Structural _Alignment")
        except Exception as e:  # noqa: BLE001 - Levenshtein/jiwer not installed
            print(f"⚠️ Levenshtein/WER disabled (distance module unavailable: {e})")
            _DIST_MOD = None
    return _DIST_MOD


def _ensure_punkt():
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
            except Exception as e:  # noqa: BLE001
                print(f"⚠️ BLEU disabled (nltk punkt unavailable): {e}")
                _punkt_ready = False
    return _punkt_ready


def _lexical_metrics(reference, candidate, wanted):
    """Selected Jaccard/ROUGE/BLEU; each guarded so one failure records None."""
    out = {c: None for c in _LEXICAL if c in wanted}
    if "jaccard" in wanted:
        # Jaccard needs no third-party library (same algorithm as
        # Lexical_Overlap_Metrics.calculate_jaccard_similarity).
        words1 = set(str(reference).lower().split())
        words2 = set(str(candidate).lower().split())
        union = words1 | words2
        out["jaccard"] = round(len(words1 & words2) / len(union), 4) if union else 0.0

    rouge_wanted = [c for c in ("rouge1", "rouge2", "rougeL") if c in wanted]
    if rouge_wanted or "bleu" in wanted:
        lom = _lom()
        if lom is not None:
            if rouge_wanted:
                try:
                    rouge = lom.calculate_rouge_scores(reference, candidate)
                    for c, key in (("rouge1", "ROUGE-1"), ("rouge2", "ROUGE-2"),
                                   ("rougeL", "ROUGE-L")):
                        if c in wanted:
                            out[c] = round(rouge.get(key, 0.0), 4)
                except Exception as e:  # noqa: BLE001
                    print(f"⚠️ ROUGE failed: {e}")
            if "bleu" in wanted and _ensure_punkt():
                try:
                    out["bleu"] = round(lom.calculate_bleu_score(reference, candidate), 4)
                except Exception as e:  # noqa: BLE001
                    print(f"⚠️ BLEU failed: {e}")
    return out


def _distance_metrics(reference, candidate, wanted):
    """Selected Levenshtein/similarity/WER from the structural-alignment
    module; when its libraries are missing, difflib still supplies the
    similarity ratio."""
    out = {c: None for c in _DISTANCE if c in wanted}
    mod = _distance_mod()
    if mod is not None:
        try:
            res = mod.calculate_edit_distance_metrics(reference, candidate)
            if "levenshtein_distance" in wanted:
                out["levenshtein_distance"] = res["character_level"]["levenshtein_distance"]
            if "similarity_ratio" in wanted:
                out["similarity_ratio"] = round(res["character_level"]["similarity_ratio"], 4)
            if "word_error_rate" in wanted:
                out["word_error_rate"] = round(res["word_level"]["word_error_rate"], 4)
            return out
        except Exception as e:  # noqa: BLE001
            print(f"⚠️ Distance metrics failed: {e}")
    if "similarity_ratio" in wanted:
        import difflib
        out["similarity_ratio"] = round(
            difflib.SequenceMatcher(None, str(reference), str(candidate)).ratio(), 4)
    return out


# --------------------------------------------------- semantic embeddings -- #
# embedding_cosine / bertscore_* measure *meaning* overlap instead of word
# overlap. They are computed with the shared embedding backends via
# data_extraction.evaluation.embedding_metrics, which drives llm_utils (remote
# OpenAI-compatible /embeddings, or the local HuggingFace model). Unlike the
# lexical/distance metrics they need a backend + a service/model, so they are
# configured once per run (configure_embeddings, or the ``embedding`` argument
# of evaluate_run / evaluate_workbook) and the vectors are cached per run — the
# reference sentences and candidate items are embedded in batches (one request
# per batch) so the API's rate limiter / pre-request delay never fires per pair.
#
# The per-pair contract matches _lexical_metrics / _distance_metrics exactly
# (reference, candidate, wanted) -> {metric: value|None}, so _embedding_metrics
# slots straight into evaluate_row (whole cell vs whole section) and
# evaluate_row_items (each item vs each sentence, best kept). Both are guarded:
# disabled or failed embeddings record None, never aborting the run.

_EMB = "unchecked"           # embedding_metrics module, None if unusable
_EMB_CFG = {
    "enabled": False,
    "backend": "api",         # "api" (remote /embeddings) or "local" (HF model)
    "service": None,          # "BlaBla" / "DLR Ollama" for the api backend
    "model": None,            # explicit embedding model id (optional)
    "api_key": None,          # optional explicit key override
    "bertscore_granularity": None,  # None -> token for local, item for api
}
_EMB_POOL_CACHE = {}         # text -> pooled unit vector (np.ndarray)
_EMB_TOK_CACHE = {}          # text -> token matrix (np.ndarray), local only
_EMB_CACHE_CAP = 5000        # drop the caches past this many texts (bound RAM)
_EMB_WARNED = False          # warn about a backend failure only once per run


def _embedding_metrics_mod():
    """Load data_extraction.evaluation.embedding_metrics once; None if missing."""
    global _EMB
    if _EMB == "unchecked":
        try:
            from data_extraction.evaluation import embedding_metrics as mod
            _EMB = mod
        except Exception as e:  # noqa: BLE001 - numpy / module unavailable
            print(f"⚠️ Semantic-embedding metrics disabled "
                  f"(embedding_metrics unavailable: {e})")
            _EMB = None
    return _EMB


def configure_embeddings(enabled=True, backend="api", service=None, model=None,
                         api_key=None, bertscore_granularity=None,
                         llm_utils=None):
    """Configure the semantic-embedding metrics for the next evaluation and
    clear the cached vectors (model/service may have changed). Call once before
    evaluate_run / evaluate_workbook / evaluate_row — the AI Review *Evaluation*
    tab does this from its embedding widgets.

    ``llm_utils`` (the caller's already-imported module) is injected into
    embedding_metrics so it need not guess the package path. Returns the
    effective config."""
    _EMB_CFG.update(enabled=bool(enabled), backend=(backend or "api"),
                    service=service, model=model, api_key=api_key,
                    bertscore_granularity=bertscore_granularity)
    if llm_utils is not None:
        mod = _embedding_metrics_mod()
        if mod is not None and hasattr(mod, "set_llm_utils"):
            mod.set_llm_utils(llm_utils)
    _reset_embedding_cache()
    return dict(_EMB_CFG)


def _reset_embedding_cache():
    global _EMB_WARNED
    _EMB_POOL_CACHE.clear()
    _EMB_TOK_CACHE.clear()
    _EMB_WARNED = False


def _emb_wanted(wanted):
    """The embedding metrics among ``wanted`` (in output order)."""
    return [m for m in _EMBEDDING if m in set(wanted)]


def _need_tokens():
    """True when bertscore should use true token-level embeddings (local
    backend, token/auto granularity)."""
    if str(_EMB_CFG["backend"]).lower() not in ("local", "l", "hf"):
        return False
    return _EMB_CFG["bertscore_granularity"] in (None, "token")


def _warn_embeddings(msg):
    global _EMB_WARNED
    if not _EMB_WARNED:
        print(f"⚠️ Embedding metric: {msg}")
        _EMB_WARNED = True


def _cap_cache(cache):
    if len(cache) > _EMB_CACHE_CAP:
        cache.clear()


def _pre_embed(texts, tokens=False):
    """Batch-embed any not-yet-cached texts into the per-run caches. One request
    per batch (not per pair) keeps the rate limiter/pre-delay from firing
    repeatedly. ``tokens=True`` also fills the token cache when it is needed."""
    if not _EMB_CFG["enabled"]:
        return
    mod = _embedding_metrics_mod()
    if mod is None:
        return
    seen = {str(t) for t in texts if str(t).strip()}
    todo = [t for t in seen if t not in _EMB_POOL_CACHE]
    if todo:
        try:
            vecs = mod.embed_texts(todo, backend=_EMB_CFG["backend"],
                                   service=_EMB_CFG["service"],
                                   model=_EMB_CFG["model"],
                                   api_key=_EMB_CFG["api_key"])
            for t, v in zip(todo, vecs):
                _EMB_POOL_CACHE[t] = v
        except Exception as e:  # noqa: BLE001 - guarded
            _warn_embeddings(f"pooled embedding failed ({e})")
    _cap_cache(_EMB_POOL_CACHE)
    if tokens and _need_tokens():
        todo_t = [t for t in seen if t not in _EMB_TOK_CACHE]
        if todo_t:
            try:
                mats = mod.token_embed_local(todo_t)
                for t, m in zip(todo_t, mats):
                    _EMB_TOK_CACHE[t] = m
            except Exception as e:  # noqa: BLE001 - guarded
                _warn_embeddings(f"token embedding failed ({e})")
        _cap_cache(_EMB_TOK_CACHE)


def _pool_vec(text):
    t = str(text)
    v = _EMB_POOL_CACHE.get(t)
    if v is None:
        _pre_embed([t])                 # on-demand fallback (pre-embed avoids it)
        v = _EMB_POOL_CACHE.get(t)
    return v


def _tok_mat(text):
    t = str(text)
    m = _EMB_TOK_CACHE.get(t)
    if m is None:
        _pre_embed([t], tokens=True)
        m = _EMB_TOK_CACHE.get(t)
    return m


def _embedding_metrics(reference, candidate, wanted):
    """Selected semantic-embedding metrics for one (reference, candidate) pair,
    read from the per-run vector caches. Same contract/guarding as
    _lexical_metrics / _distance_metrics.

    * ``embedding_cosine`` — cosine of the two pooled unit vectors.
    * ``bertscore_p/r/f1`` — greedy token-level BERTScore on the local backend;
      on the remote backend (pooled vectors only) they fall back to the pooled
      cosine, i.e. equal ``embedding_cosine``."""
    ecols = _emb_wanted(wanted)
    out = {c: None for c in ecols}
    if not ecols or not _EMB_CFG["enabled"]:
        return out
    mod = _embedding_metrics_mod()
    if mod is None:
        return out
    try:
        want_cos = "embedding_cosine" in ecols
        bert_cols = [c for c in ("bertscore_p", "bertscore_r", "bertscore_f1")
                     if c in ecols]
        rv = _pool_vec(reference)
        cv = _pool_vec(candidate)
        pooled_cos = (None if rv is None or cv is None
                      else round(mod.cosine(cv, rv), 4))
        if want_cos:
            out["embedding_cosine"] = pooled_cos
        if bert_cols:
            p = r = f1 = None
            if _need_tokens():
                rm, cm = _tok_mat(reference), _tok_mat(candidate)
                if rm is not None and cm is not None and rm.size and cm.size:
                    p, r, f1 = mod.greedy_bertscore(cm, rm)
            if p is None:               # api backend / token path unavailable
                p = r = f1 = pooled_cos
            if "bertscore_p" in ecols:
                out["bertscore_p"] = None if p is None else round(p, 4)
            if "bertscore_r" in ecols:
                out["bertscore_r"] = None if r is None else round(r, 4)
            if "bertscore_f1" in ecols:
                out["bertscore_f1"] = None if f1 is None else round(f1, 4)
    except Exception as e:  # noqa: BLE001 - guarded: one bad pair never aborts
        _warn_embeddings(f"scoring failed ({e})")
    return out


# -------------------------------------------------------------- evaluate -- #

def _write_sheet(path, sheet, df):
    """Write/replace one sheet, creating the workbook when missing."""
    import pandas as pd

    if os.path.exists(path):
        with pd.ExcelWriter(path, mode="a", engine="openpyxl",
                            if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)
    else:
        with pd.ExcelWriter(path, mode="w", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)


def evaluate_row(row, references, data_cols, metrics=ALL_METRICS):
    """Evaluate one run-sheet row (a mapping with 'Section' + the extracted
    columns) against its reference section; returns the Eval-sheet entry.
    This is the per-section unit the AI Review tab calls right after each
    LLM result is saved, so the evaluation keeps pace with the analysis."""
    metrics = [m for m in ALL_METRICS if m in set(metrics)]
    cell_metrics = [m for m in metrics if m != "grounding"]
    grounding = "grounding" in metrics
    title = str(row.get("Section", ""))
    reference = _lookup_reference(references, title)
    entry = {"Section": title, "Reference found": bool(reference)}
    # Batch-embed the section + this row's cells once, so the per-cell cosine
    # below is a cache lookup instead of one embedding request per cell.
    if reference and _emb_wanted(cell_metrics):
        cells = [str(row.get(c)).strip() for c in data_cols
                 if row.get(c) is not None and row.get(c) == row.get(c)
                 and str(row.get(c)).strip()]
        _pre_embed([reference] + cells,
                   tokens=any(m.startswith("bertscore") for m in cell_metrics))
    row_true = row_false = 0
    for col in data_cols:
        raw = row.get(col)
        if grounding:
            items = split_items(raw)
            t = sum(1 for item in items if is_grounded(item, reference))
            f = len(items) - t
            entry[f"{col} True"] = t
            entry[f"{col} False"] = f
            row_true += t
            row_false += f
        candidate = "" if raw is None or raw != raw else str(raw).strip()
        cellm = {}
        if cell_metrics and candidate and reference:
            cellm.update(_lexical_metrics(reference, candidate, cell_metrics))
            cellm.update(_distance_metrics(reference, candidate, cell_metrics))
            cellm.update(_embedding_metrics(reference, candidate, cell_metrics))
        for mc in cell_metrics:
            entry[f"{col} {mc}"] = cellm.get(mc)
    if grounding:
        total = row_true + row_false
        entry["Row True"] = row_true
        entry["Row False"] = row_false
        entry["Coverage %"] = round(100.0 * row_true / total, 1) if total else None
    return entry


def evaluate_row_items(row, references, data_cols, metrics=ALL_METRICS):
    """Per-item counterpart of :func:`evaluate_row`: split every extracted
    cell into its items (:func:`split_items`) **and** split the row's
    reference section into sentences (:func:`split_reference`, sentence-aware:
    decimals, abbreviations and citations are protected).
    Each item is scored with every selected metric against **each** reference
    sentence, and only the **most relevant** sentence's value is kept — the
    highest for the similarity-style metrics and the lowest for the
    edit-distance / error-rate metrics (``_LOWER_IS_BETTER``), per
    :func:`_better`. Scoring against the best-matching single sentence,
    rather than the whole section, gives each item a sharper, more meaningful
    number (a two-word item is not drowned out by a 40-sentence reference).

    Returns a list of records, one per item::

        {"Section": …, "Reference found": …, "Column": …, "Item #": …,
         "Item": …,
         "grounding": True/False, "grounding__sentence": "<best sentence>",
         "jaccard": …, "jaccard__sentence": …, …}

    Each metric carries a ``<metric>__sentence`` companion naming the
    reference sentence that produced its kept value. Only the selected metrics
    appear. Guarded exactly like the cell metrics: a missing library or an
    empty reference records ``None`` (grounding stays a real True/False — an
    empty reference grounds nothing). These records feed
    :func:`write_metric_sheets`, which fans them out into one sheet per metric.
    """
    metrics = [m for m in ALL_METRICS if m in set(metrics)]
    item_metrics = [m for m in metrics if m != "grounding"]
    grounding = "grounding" in metrics
    title = str(row.get("Section", ""))
    reference = _lookup_reference(references, title)
    sentences = split_reference(reference)
    # Batch-embed every reference sentence and every candidate item of this row
    # once (one request per batch). The per-(sentence, item) scoring below then
    # reads unit vectors from the cache — the whole point of embedding the
    # divided items/sentences as lists rather than pair by pair.
    if sentences and _emb_wanted(item_metrics):
        all_items = []
        for col in data_cols:
            all_items.extend(split_items(row.get(col)))
        _pre_embed(list(sentences) + all_items,
                   tokens=any(m.startswith("bertscore") for m in item_metrics))
    records = []
    for col in data_cols:
        items = split_items(row.get(col))
        for idx, item in enumerate(items, 1):
            rec = {"Section": title, "Reference found": bool(reference),
                   "Column": col, "Item #": idx, "Item": item}
            if grounding:
                # Grounded if the item is a substring of ANY sentence; keep
                # the first sentence that contains it as the witness.
                hit = next((s for s in sentences if is_grounded(item, s)), "")
                rec["grounding"] = bool(hit)
                rec["grounding__sentence"] = hit
            if item_metrics:
                best = {m: None for m in item_metrics}
                best_sent = {m: "" for m in item_metrics}
                if item and sentences:
                    for sent in sentences:
                        vals = {}
                        vals.update(_lexical_metrics(sent, item, item_metrics))
                        vals.update(_distance_metrics(sent, item, item_metrics))
                        vals.update(_embedding_metrics(sent, item, item_metrics))
                        for m in item_metrics:
                            v = vals.get(m)
                            if v is not None and _better(m, v, best[m]):
                                best[m] = v
                                best_sent[m] = sent
                for m in item_metrics:
                    rec[m] = best[m]
                    rec[f"{m}__sentence"] = best_sent[m]
            records.append(rec)
    return records


def write_eval_sheet(file_path, run_sheet, entries, out_path=None):
    """(Re)write 'Eval <run sheet>' from evaluated entries, closed by an
    average row. Rebuilt on every call, so it is safe to call after each
    newly evaluated row (idempotent). Returns the sheet name."""
    import pandas as pd

    target = str(out_path) if out_path else str(file_path)
    out = pd.DataFrame(entries)
    # Average row — the per-run counterpart of metric_evaluator's Overview.
    avg = {"Section": "— AVERAGE —"}
    for c in out.columns:
        if c in ("Section", "Reference found"):
            continue
        vals = pd.to_numeric(out[c], errors="coerce").dropna()
        avg[c] = round(float(vals.mean()), 4) if len(vals) else None
    out = pd.concat([out, pd.DataFrame([avg])], ignore_index=True)
    eval_sheet = f"Eval {run_sheet}"[:31]
    _write_sheet(target, eval_sheet, out)
    return eval_sheet


def metric_sheet_name(metric, run_sheet):
    """Sheet name for one metric's per-item sheet (31-char Excel limit)."""
    code = METRIC_SHEET_CODES.get(metric, metric[:4].title())
    return f"{code} {run_sheet}"[:31]


def write_metric_sheets(file_path, run_sheet, item_records, metrics,
                        out_path=None):
    """(Re)write one per-item sheet **per metric** from the records
    :func:`evaluate_row_items` produced for a whole run sheet.

    Each sheet (named via :func:`metric_sheet_name`) carries one row per
    extracted item: ``Section | Reference found | Column | Item # | Item``
    plus that metric's value for the item and the ``Best Ref Sentence`` that
    produced it (the most relevant reference sentence for that item under
    that metric), closed by an average row — ``n_true/n (pct%)`` coverage for
    grounding, the numeric mean for every other metric. Rebuilt on every call
    (idempotent, mirroring :func:`write_eval_sheet`). Returns
    ``{metric: sheet name}``.
    """
    import pandas as pd

    target = str(out_path) if out_path else str(file_path)
    metrics = [m for m in ALL_METRICS if m in set(metrics)]
    base_cols = ["Section", "Reference found", "Column", "Item #", "Item"]
    written = {}
    for metric in metrics:
        rows = [{**{c: rec.get(c) for c in base_cols},
                 metric: rec.get(metric),
                 "Best Ref Sentence": rec.get(f"{metric}__sentence", "")}
                for rec in item_records if metric in rec]
        df = pd.DataFrame(rows, columns=base_cols + [metric, "Best Ref Sentence"])
        # Average row — grounding gets a coverage summary, the numeric
        # metrics the plain mean (None-safe on both paths).
        avg = {c: pd.NA for c in base_cols}
        avg["Section"] = "— AVERAGE —"
        avg["Best Ref Sentence"] = pd.NA
        if metric == "grounding":
            flags = [bool(r[metric]) for r in rows if r.get(metric) is not None]
            n_true, n = sum(flags), len(flags)
            pct = round(100.0 * n_true / n, 1) if n else None
            avg[metric] = f"{n_true}/{n} ({pct}%)" if n else "0/0"
            df[metric] = df[metric].map(
                lambda v: str(bool(v)) if v is not None and v == v else None)
        else:
            vals = pd.to_numeric(df[metric], errors="coerce").dropna()
            avg[metric] = round(float(vals.mean()), 4) if len(vals) else None
        df = pd.concat([df, pd.DataFrame([avg])], ignore_index=True)
        sheet = metric_sheet_name(metric, run_sheet)
        _write_sheet(target, sheet, df)
        written[metric] = sheet
    return written


def _preview(text, limit=180):
    text = " ".join(str(text if text is not None else "").split())
    return text if len(text) <= limit else text[:limit] + " …"


def evaluate_run(file_path, run_sheet, references, metrics=ALL_METRICS,
                 uniq_columns=None, out_path=None, log=None, embedding=None):
    """
    Evaluate one column-analysis snapshot sheet against its reference
    sections. ``log`` is an optional ``log(message)`` callback that receives
    a progress line per step (which row is evaluated, its reference, the
    evaluated texts and the results).

    ``references`` is ``{section title: section text}`` (the text the LLM
    analyzed). ``metrics`` selects the evaluations (see :data:`ALL_METRICS`).
    The overview is written as ``Eval <run sheet>`` (unchanged shape), and
    every selected metric additionally gets its own per-item sheet via
    :func:`write_metric_sheets` (one row per extracted item: section,
    column, item content, metric value) — into the analysis workbook
    itself, or into ``out_path`` when given (a new/separate workbook, which
    then also receives a copy of the run sheet so the unique-element pass and
    the reader have the data next to the evaluation). When ``uniq_columns``
    is given the unique generated values of those columns are identified via
    save_unique_elements_to_new_sheet (``Uniq <run sheet>``) and grounded
    against the combined reference text into ``Uniq Eval <run sheet>``.
    Re-runs replace the sheets (idempotent). Returns a summary dict.
    """
    import pandas as pd

    file_path = str(file_path)
    metrics = [m for m in ALL_METRICS if m in set(metrics)]
    if not metrics:
        raise ValueError("No evaluations selected.")
    grounding = "grounding" in metrics
    # Apply per-run embedding config (if passed) and start with a fresh vector
    # cache so this run re-embeds under the current backend/service/model.
    if embedding is not None:
        configure_embeddings(**embedding)
    else:
        _reset_embedding_cache()
    if _emb_wanted(metrics) and not _EMB_CFG["enabled"]:
        _warn_embeddings("an embedding metric is selected but embeddings are "
                         "not configured/enabled — recording None for it "
                         "(call configure_embeddings first).")

    df = pd.read_excel(file_path, sheet_name=run_sheet, engine="openpyxl")
    if "Section" not in df.columns:
        raise ValueError(f"Sheet '{run_sheet}' has no 'Section' column — "
                         "not a column-analysis run sheet.")
    data_cols = [c for c in df.columns if c != "Section"]

    target = str(out_path) if out_path else file_path
    if os.path.abspath(target) != os.path.abspath(file_path):
        # Separate evaluation workbook: carry the analyzed data along.
        _write_sheet(target, run_sheet, df)

    if log:
        log(f"⚖ Evaluating sheet '{run_sheet}': {len(df)} row(s), "
            f"columns [{', '.join(data_cols)}], metrics [{', '.join(metrics)}]")
    entries, item_records = [], []
    for i, rec in enumerate(df.to_dict("records"), 1):
        if log:
            title = str(rec.get("Section", ""))
            reference = _lookup_reference(references, title)
            log(f"▶ [{i}/{len(df)}] '{title}' — reference "
                f"({len(reference)} chars): {_preview(reference)}"
                if reference else
                f"▶ [{i}/{len(df)}] '{title}' — no matching reference section!")
            for col in data_cols:
                log(f"   evaluated text [{col}]: {_preview(rec.get(col))}")
        entry = evaluate_row(rec, references, data_cols, metrics)
        entries.append(entry)
        item_records.extend(
            evaluate_row_items(rec, references, data_cols, metrics))
        if log:
            log("   result: " + _preview(
                "; ".join(f"{k}={v}" for k, v in entry.items()
                          if k != "Section"), 300))
    eval_sheet = write_eval_sheet(target, run_sheet, entries)
    if log:
        log(f"→ wrote '{eval_sheet}'")
    metric_sheets = write_metric_sheets(target, run_sheet, item_records,
                                        metrics)
    if log and metric_sheets:
        log(f"→ per-item metric sheets ({len(item_records)} item(s)): "
            + ", ".join(f"{m} → '{s}'" for m, s in metric_sheets.items()))

    total_true = sum(e.get("Row True") or 0 for e in entries)
    total_false = sum(e.get("Row False") or 0 for e in entries)
    grand = total_true + total_false
    summary = {
        "out_path": target,
        "eval_sheet": eval_sheet,
        "metric_sheets": metric_sheets,
        "items": len(item_records),
        "rows": len(entries),
        "metrics": metrics,
        "true": total_true if grounding else None,
        "false": total_false if grounding else None,
        "coverage": (round(100.0 * total_true / grand, 1)
                     if grounding and grand else None),
    }
    if uniq_columns:
        uniq_summary = evaluate_uniques(
            target, run_sheet, references,
            [c for c in uniq_columns if c in data_cols])
        summary.update(uniq_summary)
        if log and uniq_summary.get("uniq_eval_sheet"):
            cov = uniq_summary.get("uniq_coverage") or {}
            log(f"→ unique values evaluated → '{uniq_summary['uniq_eval_sheet']}'"
                + (" (" + ", ".join(f"{c}: {p}%" for c, p in cov.items()
                                    if p is not None) + ")" if cov else ""))
    return summary


def evaluate_uniques(file_path, run_sheet, references, columns):
    """Identify the unique generated values via save_unique_elements_to_new_sheet
    (refreshing ``Uniq <run sheet>``), then ground every unique element against
    the combined reference text into ``Uniq Eval <run sheet>``."""
    import pandas as pd

    if not columns:
        return {}
    uniq_sheet = f"Uniq {run_sheet}"[:31]
    ok = _excel_utils().save_unique_elements_to_new_sheet(
        file_path, columns, new_sheet_name=uniq_sheet, source_sheet=run_sheet)
    if not ok:
        return {"uniq_sheet": None}

    combined = "\n".join(str(v or "") for v in references.values()).lower()
    df = pd.read_excel(file_path, sheet_name=uniq_sheet, engine="openpyxl")
    blocks, uniq_coverage = [], {}
    for col in columns:
        ucol, ccol = f"Unique_{col}", f"Count_{col}"
        if ucol not in df.columns:
            continue
        sub = df[[ucol, ccol]].dropna(subset=[ucol])
        grounded = [str(el).lower() in combined for el in sub[ucol]]
        n_true, n = sum(grounded), len(grounded)
        pct = round(100.0 * n_true / n, 1) if n else None
        # Element | count | grounded, closed by a per-column coverage row.
        blocks.append(pd.DataFrame({
            ucol: list(sub[ucol]) + ["— COVERAGE —"],
            ccol: list(sub[ccol]) + [pd.NA],
            f"Grounded_{col}": [str(g) for g in grounded]
                               + [f"{n_true}/{n} ({pct}%)" if n else "0/0"],
        }))
        uniq_coverage[col] = pct
    if not blocks:
        return {"uniq_sheet": uniq_sheet}

    target_df = pd.concat(blocks, axis=1)
    uniq_eval_sheet = f"Uniq Eval {run_sheet}"[:31]
    _write_sheet(file_path, uniq_eval_sheet, target_df)
    return {"uniq_sheet": uniq_sheet, "uniq_eval_sheet": uniq_eval_sheet,
            "uniq_coverage": uniq_coverage}


def evaluate_workbook(file_path, references, metrics=ALL_METRICS,
                      run_sheets=None, uniq_columns="all", out_path=None,
                      log=None, embedding=None):
    """
    Evaluate every ``Run N …`` snapshot sheet of a column-analysis workbook
    (or just ``run_sheets``). ``uniq_columns="all"`` evaluates the uniques of
    every extracted column; pass a list to restrict, or None to skip.
    ``out_path`` redirects all result sheets into a separate workbook;
    ``log`` is forwarded to :func:`evaluate_run` for per-row progress lines.
    ``embedding`` is an optional config dict for the semantic-embedding metrics
    (see :func:`configure_embeddings`); it is applied once here so each sheet
    reuses the same backend/service. Returns ``{run sheet: summary dict}``.
    """
    import pandas as pd

    if embedding is not None:
        configure_embeddings(**embedding)
    file_path = str(file_path)
    with pd.ExcelFile(file_path, engine="openpyxl") as xls:
        all_runs = [s for s in xls.sheet_names if s.startswith("Run ")]
    sheets = run_sheets if run_sheets else all_runs
    results = {}
    for sheet in sheets:
        cols = uniq_columns
        if uniq_columns == "all":
            head = pd.read_excel(file_path, sheet_name=sheet, nrows=0,
                                 engine="openpyxl")
            cols = [c for c in head.columns if c != "Section"]
        # embedding already applied above -> pass None so evaluate_run only
        # refreshes its per-run cache (keeps the shared config in place).
        results[sheet] = evaluate_run(file_path, sheet, references,
                                      metrics=metrics, uniq_columns=cols,
                                      out_path=out_path, log=log,
                                      embedding=None)
    return results


def _easa_sections(nodes, out):
    """Walk an EASA rules hierarchy: title/text per node, depth first (same
    title rule as the EASA JSON review UI, so run-sheet Sections match)."""
    for node in nodes:
        if not isinstance(node, dict):
            continue
        attrs = node.get("attributes", {}) or {}
        title = (attrs.get("source-title") or attrs.get("title")
                 or node.get("element_type", "node"))
        out.setdefault(str(title), node.get("text_content", "") or "")
        _easa_sections(node.get("children") or [], out)


def references_from_json(path):
    """{section title: text} from the sections JSON an analysis was made
    from, auto-detecting the shape: a chunks cache / Processed_chunks.json
    (normalized exactly like the chunk AI Review UI) or an EASA structured
    extraction JSON (rules hierarchy). The titles match the 'Section' column
    the run sheets carry, so each row finds its reference automatically."""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # EASA structured JSON: {rules_hierarchy: [...]} or a bare node list.
    hierarchy = None
    if isinstance(data, dict) and data.get("rules_hierarchy"):
        hierarchy = data["rules_hierarchy"]
    elif isinstance(data, list) and data and isinstance(data[0], dict) \
            and ("text_content" in data[0] or "children" in data[0]):
        hierarchy = data
    if hierarchy:
        out = {}
        _easa_sections(hierarchy, out)
        return out

    from data_extraction.chunking.ai_review_ui import sections_from_payload
    return dict(sections_from_payload(data))


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Evaluate AI Review column-analysis run sheets against "
                    "the sections JSON they were generated from.")
    parser.add_argument("workbook", help="column-analysis .xlsx workbook")
    parser.add_argument("sections_json",
                        help="chunks cache or Processed_chunks.json with the "
                             "reference section texts")
    parser.add_argument("--sheet", action="append",
                        help="run sheet to evaluate (repeatable; default: "
                             "all 'Run …' sheets)")
    parser.add_argument("--metrics", nargs="*", default=list(ALL_METRICS),
                        choices=ALL_METRICS,
                        help="evaluations to run (default: all)")
    parser.add_argument("--out", help="write the evaluation sheets into this "
                                      "separate .xlsx instead of the workbook")
    parser.add_argument("--no-uniques", action="store_true",
                        help="skip the unique-element evaluation")
    parser.add_argument("--embed-backend", choices=("api", "local"),
                        default="api",
                        help="backend for the embedding_cosine / bertscore "
                             "metrics (default: api)")
    parser.add_argument("--embed-service", choices=("BlaBla", "DLR Ollama"),
                        default="DLR Ollama",
                        help="remote service for --embed-backend api "
                             "(default: DLR Ollama)")
    parser.add_argument("--embed-model",
                        help="explicit embedding model id (optional; otherwise "
                             "the service's default embedding model)")
    args = parser.parse_args(argv)

    references = references_from_json(args.sections_json)
    if not references:
        sys.exit(f"No sections found in {args.sections_json}.")
    embedding = None
    if any(m in _EMBEDDING for m in args.metrics):
        embedding = {"enabled": True, "backend": args.embed_backend,
                     "service": args.embed_service, "model": args.embed_model}
    results = evaluate_workbook(
        args.workbook, references, metrics=args.metrics, run_sheets=args.sheet,
        uniq_columns=None if args.no_uniques else "all", out_path=args.out,
        embedding=embedding)
    if not results:
        sys.exit(f"No 'Run …' sheets found in {args.workbook}.")
    for sheet, s in results.items():
        line = f"✅ {sheet}: {s['rows']} row(s)"
        if s.get("coverage") is not None:
            line += f", grounded {s['true']}/{s['true'] + s['false']} ({s['coverage']}%)"
        line += f" → '{s['eval_sheet']}' in {s['out_path']}"
        if s.get("metric_sheets"):
            line += (f", {s.get('items', 0)} item(s) → "
                     + ", ".join(f"'{v}'" for v in s["metric_sheets"].values()))
        if s.get("uniq_eval_sheet"):
            line += f", uniques → '{s['uniq_eval_sheet']}'"
        print(line)


if __name__ == "__main__":
    main()