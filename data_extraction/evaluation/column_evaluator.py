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

Columns that hold cross-references (name contains ``reference`` — see
``_REFERENCE_COLUMN_HINTS``) get one extra pass: each unique value is matched
against the **section titles** (``match_sections``). Values that start with a
rule identifier — a 2-3 capital-letter prefix such as CS / AMC / GM / SC plus a
``dd.dd``-``dd.dddd`` number (``AMC 25.21(g)(a)``) — are matched by that
identifier alone, by regex, so qualifiers and trailing prose are ignored and
``AMC 25.21`` never lands on ``AMC 25.219``; any other value is stripped of its
``(…)`` qualifiers and matched as a case-insensitive substring in either
direction. The ``Uniq <run sheet>`` sheet keeps its
``Unique_<col>``/``Count_<col>`` pair as-is and gains ``Section_<col>``,
``Section_Count_<col>`` and ``Section_Matches_<col>`` beside it — the section
tally lives next to the unique elements, most-referenced first — while every
individual match is written to a ``Ref Map <run sheet>`` audit sheet.

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
from functools import lru_cache
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

# Extracted columns whose NAME contains one of these hints hold cross-references
# to other sections/rules, so their unique values are additionally matched
# against the reference *section titles* (see :func:`evaluate_uniques`).
# Matching is case-insensitive; extend this tuple to cover more column names.
_REFERENCE_COLUMN_HINTS = ("reference",)

# A substring match shorter than this many characters is ignored when matching a
# unique reference value against a section title: one- or two-character needles
# ("A", "3") hit almost every title and would only add noise. Raise it for
# stricter matching, lower it if your section titles are genuinely that short.
_REF_MATCH_MIN_CHARS = 3

# Sub-paragraph qualifiers in a reference value are dropped before it is matched
# against the section titles: "AMC 25.21(g)(a)" points at the section titled
# "AMC 25.21 …", which carries no "(g)(a)" in its title, so the parentheses have
# to go or the value would never match. Applied repeatedly, so nested groups
# ("(a(b))") are removed too.
_PARENTHETICAL = re.compile(r"\([^()]*\)")

# Rule identifiers ("CS 25.1309", "AMC 25.21", "GM 26.30") are matched by their
# identifier alone rather than as raw substrings. The prefix is no longer a
# hard-coded CS/AMC/GM list but any 2-3 **capital letters** ("CS", "AMC", "GM",
# "SC", "CRI", …), followed by whitespace, a 2-digit book/part number, a dot
# and a 2-4 digit rule number. Anchored at the
# start, so only values that *begin* with such an identifier take this path;
# anything else (and values that do not fit the digit format, e.g. "GM 21.A.14",
# whose "A" is not a digit, or "CS 1.5", whose parts are too short) falls back
# to substring matching. The prefix must genuinely be upper-case — no IGNORECASE
# here, or ordinary lowercase words ("and 25.21 …") would be taken for rule
# prefixes. Trailing qualifiers and prose after the identifier are ignored.
_RULE_ID = re.compile(r"^\s*([A-Z]{2,3})\s+(\d{2})\.(\d{2,4})(?!\d)")

# Bare rule identifiers — values that carry no capital-letter prefix but *start*
# with a rule number: 2 digits, a dot, then everything up to the first '(' or
# whitespace ("23.1001(a) through (f)" -> "23.1001", "23.1011 Oil system
# General" -> "23.1011", "21.A.3 Failures …" -> "21.A.3"). Tried only when
# _RULE_ID above did not recognise the value, and — unlike the prefixed path —
# NOT authoritative: a bare number is looser evidence, so when it matches no
# section title the value falls through to the plain substring hunt instead of
# being declared unmatched. Trailing punctuation on the captured number is
# stripped so "23.1011," or "23.1011." do not poison the regex needle.
_BARE_RULE_ID = re.compile(r"^\s*(\d{2}\.[^\s(]+)")

# Detects whether the section titles themselves are written with rule prefixes
# ("AMC 25.1309 System design") or bare ("25.1309 System design"). Same shape as
# _RULE_ID (2-3 capitals, 2 digits, dot, 2 digits …) but searched anywhere in
# the title, and likewise case-sensitive on the prefix.
_RULE_PREFIX_IN_TITLE = re.compile(r"\b([A-Z]{2,3})\s+\d{2}\.\d{2}")

# When the full identifier ("AMC 25.21") matches no section title, retry with the
# rule number alone ("25.21") — but only when the titles carry no rule
# prefix at all (e.g. "25.21 Proof of compliance"). Where the titles *are*
# prefixed, a missing prefix match is a real miss, so "CS 25.1309" must not land
# on an "AMC 25.1309 …" title (different book, same rule number). Set to False to
# switch the fallback off entirely and always require the prefix to match.
_RULE_MATCH_NUMBER_FALLBACK = True

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


@lru_cache(maxsize=4096)
def _squash_ws(text):
    """Lower-cased with **every** whitespace character removed.

    PDF extraction routinely drops or inserts spaces ("AMC2&3to CS-23",
    "inaremark onthe specific lineofthat"), so the reference text and the
    LLM's correctly-spaced item can differ by nothing but a space and still
    fail a literal comparison. Collapsing runs of whitespace is not enough —
    the space has to go entirely — hence ``"".join(text.split())``, which also
    covers non-breaking spaces, tabs and newlines. Cached because the same long
    reference is squashed once per item/sentence pair.
    """
    return "".join(str(text).split()).lower()


def is_grounded(item, reference):
    """data_evaluator's Is_Subset check: the extracted item is contained in
    the reference text (case-insensitive substring).

    The comparison is also **whitespace-insensitive**: a literal match is tried
    first, and only if that fails are both sides re-checked with all whitespace
    removed (:func:`_squash_ws`). That keeps items grounded when the reference
    text carries the extractor's missing/extra spaces — matching an item that
    genuinely appears in the section is the point, and a space is not a
    difference in content. It cannot un-ground anything: whatever matched
    literally still matches squashed.
    """
    if not reference:
        return False
    if str(item).lower() in str(reference).lower():
        return True
    squashed_item = _squash_ws(item)
    return bool(squashed_item) and squashed_item in _squash_ws(reference)


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
_EMB_FAILED = set()          # texts the backend could not embed (never retried)
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
    _EMB_FAILED.clear()
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
    repeatedly. ``tokens=True`` also fills the token cache when it is needed.

    embedding_metrics retries a failing call in smaller batches and then string
    by string, so a partial result is normal: whatever came back is cached, and
    anything that could not be embedded is remembered in ``_EMB_FAILED`` so the
    per-pair lookups below neither retry it nor score it."""
    if not _EMB_CFG["enabled"]:
        return
    mod = _embedding_metrics_mod()
    if mod is None:
        return
    seen = {str(t) for t in texts if str(t).strip()}
    todo = [t for t in seen
            if t not in _EMB_POOL_CACHE and t not in _EMB_FAILED]
    if todo:
        try:
            got = mod.embed_texts_map(todo, backend=_EMB_CFG["backend"],
                                      service=_EMB_CFG["service"],
                                      model=_EMB_CFG["model"],
                                      api_key=_EMB_CFG["api_key"])
            _EMB_POOL_CACHE.update(got)
            missing = [t for t in todo if t not in got]
            if missing:
                _EMB_FAILED.update(missing)
                _warn_embeddings(f"{len(missing)} of {len(todo)} string(s) could "
                                 f"not be embedded even individually — their "
                                 f"metrics stay blank")
        except Exception as e:  # noqa: BLE001 - guarded
            _EMB_FAILED.update(todo)
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
    if v is None and t not in _EMB_FAILED:
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
    * ``bertscore_p/r/f1`` — greedy token-level BERTScore, **local backend only**.
      The remote /embeddings endpoint returns a single pooled vector per string,
      and this dispatch scores one sentence against one item, so the cosine
      matrix would be 1x1 and P, R and F1 would each collapse to the pooled
      cosine — three columns duplicating ``embedding_cosine`` and telling you
      nothing. They are therefore left None ("not measured") rather than filled
      with a look-alike number. Same when the token embeddings are unavailable
      for any other reason (e.g. they failed and the ladder could not recover)."""
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
        if bert_cols and _need_tokens():
            rm, cm = _tok_mat(reference), _tok_mat(candidate)
            if rm is not None and cm is not None and rm.size and cm.size:
                p, r, f1 = mod.greedy_bertscore(cm, rm)
                for col, val in (("bertscore_p", p), ("bertscore_r", r),
                                 ("bertscore_f1", f1)):
                    if col in ecols:
                        out[col] = None if val is None else round(val, 4)
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
    if (_EMB_CFG["enabled"] and not _need_tokens()
            and any(m.startswith("bertscore") for m in metrics)):
        msg = ("BERTScore needs token-level embeddings, which only the local "
               "backend provides — the remote /embeddings endpoint returns one "
               "pooled vector per string, so P/R/F1 could only echo "
               "embedding_cosine. Leaving bertscore_* blank; use "
               "embedding_cosine here, or switch the backend to 'local'.")
        print(f"⚠️ {msg}")
        if log:
            log(f"⚠️ {msg}")

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
        if log and uniq_summary.get("ref_map_sheet"):
            secs = uniq_summary.get("ref_sections") or {}
            log(f"→ reference values matched to section titles → "
                f"'{uniq_summary['ref_map_sheet']}'"
                + (" (" + ", ".join(f"{c}: {len(t)} section(s)"
                                    for c, t in secs.items()) + ")"
                   if secs else " (no section matched)"))
    return summary


def is_reference_column(col):
    """True when an extracted column's NAME marks it as holding cross-references
    to other sections/rules (``_REFERENCE_COLUMN_HINTS``, case-insensitive), so
    its unique values are also matched against the reference section titles."""
    name = str(col).lower()
    return any(hint in name for hint in _REFERENCE_COLUMN_HINTS)


def _count_value(v):
    """The ``Count_<col>`` cell as an int; 0 for blank/NaN/non-numeric."""
    try:
        if v is None or v != v:  # None or NaN
            return 0
        return int(float(v))
    except (TypeError, ValueError):
        return 0


def strip_parentheticals(text):
    """Drop every ``(…)`` qualifier from a reference value and tidy what is left,
    so the bare rule identifier remains::

        "AMC 25.21(g)(a)"   -> "AMC 25.21"
        "21.A.14(b)"        -> "21.A.14"
        "see 21.A.3A, point (b)" -> "see 21.A.3A"

    Sub-paragraph qualifiers live in the *value*, never in the section title, so
    they must go before matching or the value could not match its own section.
    Nested groups are removed too (the substitution runs until it is stable), and
    the leftover whitespace / orphaned separators are cleaned up. Returns ``""``
    only when the value is nothing but parentheses.
    """
    out = str(text or "")
    prev = None
    while prev != out:
        prev = out
        out = _PARENTHETICAL.sub(" ", out)
    out = re.sub(r"\s+", " ", out).strip()
    return out.strip(" ,;:-")


def rule_identifier(text):
    """The rule identifier a reference value **starts with**, or ``None`` when
    it does not begin with one in that format (``_RULE_ID``: 2-3 capital
    letters, whitespace, a 2-digit book/part number, a dot, 2-4 digits)::

        "AMC 25.21(g)(a)"        -> ("AMC", "25.21")
        "CS 25.1309(a) blah"     -> ("CS",  "25.1309")
        "GM 26.30"               -> ("GM",  "26.30")
        "SC 23.2005 something"   -> ("SC",  "23.2005")
        "GM 21.A.14"             -> None   # "A" is not a digit
        "CS 1.5"                 -> None   # parts too short (need dd.dd)
        "cs 25.1309"             -> None   # prefix must be capital letters
        "21.A.3A"                -> None   # no capital-letter prefix

    Everything after the identifier — sub-paragraph qualifiers, prose — is
    ignored, so the identifier is what gets matched against the section titles.
    Returns ``(prefix, number)``.
    """
    m = _RULE_ID.match(str(text or ""))
    if not m:
        return None
    return m.group(1), f"{m.group(2)}.{m.group(3)}"


def bare_rule_identifier(text):
    """The bare rule number a reference value **starts with** — 2 digits, a
    dot, then everything up to the first ``(`` or whitespace — or ``None`` when
    the value does not begin that way (``_BARE_RULE_ID``)::

        "23.1001(a) through (f)"     -> "23.1001"
        "23.1011"                    -> "23.1011"
        "23.1011 Oil system General" -> "23.1011"
        "21.A.3 Failures …"          -> "21.A.3"
        "23.1011, see also …"        -> "23.1011"  # trailing ',' stripped
        "2.5 percent"                -> None       # needs 2 digits before dot
        "see 23.1011"                -> None       # must start with the number

    Trailing punctuation on the captured number is stripped; a capture that is
    left with nothing after the dot (``"23."``) is rejected. Meant to run
    *after* :func:`rule_identifier` — prefixed values ("AMC 25.21") never reach
    this because they do not start with a digit.
    """
    m = _BARE_RULE_ID.match(str(text or ""))
    if not m:
        return None
    number = m.group(1).rstrip(".,;:-")
    # "23." (nothing left after the dot once punctuation is gone) is no number.
    if len(number) <= 3:
        return None
    return number


def _rule_regex(needle):
    """Regex matching ``needle`` ("AMC 25.21" / "25.21") inside a section title:
    whitespace between the parts is flexible, and the trailing ``(?!\\d)`` guard
    stops "AMC 25.21" from also matching "AMC 25.219"."""
    parts = [re.escape(p) for p in str(needle).split()]
    return re.compile(r"\b" + r"\s+".join(parts) + r"(?!\d)", re.IGNORECASE)


def _titles_use_rule_prefix(references):
    """True when the section titles themselves carry capital-letter rule
    prefixes ("AMC 25.1309 System design", "SC 23.2005 …") rather than bare
    rule numbers."""
    return any(_RULE_PREFIX_IN_TITLE.search(str(t or "")) for t in references)


def match_rule_sections(prefix, number, references):
    """Match a rule identifier (2-3 capital-letter prefix + number) against the
    section titles **by regex** (:func:`_rule_regex`) rather than by raw
    substring, so "AMC 25.21" finds "AMC 25.21 Proof of compliance" but not
    "AMC 25.219 …".

    The full identifier ("AMC 25.21") is tried first. Only if no title carries it
    — and only when the titles are written without rule prefixes at all
    (:func:`_titles_use_rule_prefix`) and ``_RULE_MATCH_NUMBER_FALLBACK`` is on —
    is the rule number alone ("25.21") retried. That keeps "CS 25.1309" off an
    "AMC 25.1309 …" title (same rule number, different book) while still matching
    bare titles like "25.1309 System design". Returns the same
    ``[(section title, how, needle), ...]`` shape as :func:`match_sections`.
    """
    full = f"{prefix} {number}"
    rx = _rule_regex(full)
    hits = [(title, f"rule id '{full}' in section title", full)
            for title in references if rx.search(str(title or ""))]
    if hits or not _RULE_MATCH_NUMBER_FALLBACK:
        return hits
    if _titles_use_rule_prefix(references):
        # The titles do use prefixes, so this rule genuinely is not among them.
        return []
    rx = _rule_regex(number)
    return [(title, f"rule no. '{number}' in section title", number)
            for title in references if rx.search(str(title or ""))]


def match_bare_rule_sections(number, references):
    """Match a bare rule number (:func:`bare_rule_identifier`, e.g. ``23.1001``)
    against the section titles **by regex** (:func:`_rule_regex`), exactly like
    the prefixed path: flexible whitespace, and the trailing digit guard keeps
    ``23.100`` off ``23.1001 …`` titles. Works whether the titles are bare
    (``"23.1001 Fuel system"``) or prefixed (``"CS 23.1001 Fuel system"``) —
    the number matches inside either. Returns the same
    ``[(section title, how, needle), ...]`` shape as :func:`match_sections`;
    the caller falls back to substring matching when this returns ``[]``.
    """
    rx = _rule_regex(number)
    return [(title, f"bare rule no. '{number}' in section title", number)
            for title in references if rx.search(str(title or ""))]


def match_sections(element, references):
    """Match one unique reference value against the reference **section titles**.

    Three paths, tried in order, picked from the value itself:

    * **Rule-identifier values** — a 2-3 capital-letter prefix (CS, AMC, GM,
      SC, …) plus a ``dd.dd``-``dd.dddd`` number (:func:`rule_identifier`
      recognises the identifier the value starts with, e.g.
      ``"AMC 25.21(g)(a)"`` -> ``AMC 25.21``) — are matched by
      :func:`match_rule_sections`, i.e. by regex on
      that identifier alone. Sub-paragraph qualifiers and any trailing prose are
      ignored, and the digit guard keeps ``AMC 25.21`` off ``AMC 25.219``. This
      path is authoritative: when the identifier matches no title the value is
      reported as unmatched rather than falling back to a looser substring hunt.
    * **Bare rule numbers** — no prefix, but the value starts with 2 digits and
      a dot (:func:`bare_rule_identifier` takes everything up to the first
      ``(`` or whitespace: ``"23.1001(a) through (f)"`` -> ``23.1001``,
      ``"23.1011 Oil system General"`` -> ``23.1011``) — are matched by
      :func:`match_bare_rule_sections`, the same regex machinery with the same
      digit guard. Unlike the prefixed path this one is *not* authoritative:
      when the number matches no title, the value drops down to the substring
      path below.
    * **Everything else** keeps the substring behaviour: the value is stripped of
      its ``(…)`` qualifiers (:func:`strip_parentheticals`) and matched
      case-insensitively in either direction — the value may name a section in
      short form (``"21.A.3"`` inside the title ``"21.A.3 Failures, malfunctions
      and defects"``) or quote the title in full inside a longer phrase
      (``"see 21.A.3 Failures …, point (b)"``). The needle (the shorter,
      contained side) must be at least ``_REF_MATCH_MIN_CHARS`` characters, which
      vetoes trivial one/two-character hits. The stripped form is tried first;
      the raw text is then tried as a fallback, because dropping the qualifiers
      shortens the value and could otherwise lose a match in the "title inside
      value" direction. The first direction that hits wins per title.

    Returns ``[(section title, how, needle), ...]`` — possibly several, since one
    value can legitimately name more than one section — where ``how`` names the
    rule/direction that matched and ``needle`` is the form of the value that
    actually matched. Returns ``[]`` for a blank value or when nothing matches.
    """
    el_orig = str(element or "").strip()
    el_raw = el_orig.lower()
    if not el_raw:
        return []
    # Rule-identifier values: match the identifier by regex, not by substring.
    # Checked against the original-case text — the prefix must be capitals, so
    # the lower-cased form used by the substring path would never match.
    rule = rule_identifier(el_orig)
    if rule:
        return match_rule_sections(rule[0], rule[1], references)
    # Bare rule numbers ("23.1001(a) through (f)" -> "23.1001"): regex match
    # like the prefixed path, but NOT authoritative — no title hit means fall
    # through to the substring hunt below rather than report a miss.
    bare = bare_rule_identifier(el_orig)
    if bare:
        hits = match_bare_rule_sections(bare, references)
        if hits:
            return hits
    el_norm = strip_parentheticals(el_raw)
    # Qualifier-free form first (the intent); raw text as a fallback so no match
    # that worked before the stripping is lost.
    needles = [n for n in dict.fromkeys((el_norm, el_raw)) if n]
    hits = []
    for title in references:
        tl = str(title or "").strip().lower()
        if not tl:
            continue
        for cand in needles:
            if len(cand) >= _REF_MATCH_MIN_CHARS and cand in tl:
                hits.append((title, "value in section title", cand))
                break
            if len(tl) >= _REF_MATCH_MIN_CHARS and tl in cand:
                hits.append((title, "section title in value", cand))
                break
    return hits


def _reference_section_counts(df, col, references):
    """Match every unique value of one reference column against the section
    titles. Returns ``(counts, map_rows)`` where ``counts`` is
    ``{section title: [occurrences, matched values]}`` — ``occurrences`` sums the
    values' ``Count_<col>`` (how often the reference was actually generated),
    ``matched values`` counts the distinct unique values that hit that section —
    and ``map_rows`` are the per-match audit records for the Ref Map sheet
    (including the ``(…)``-stripped form each value was matched by)."""
    ucol, ccol = f"Unique_{col}", f"Count_{col}"
    sub = df[[ucol, ccol]].dropna(subset=[ucol]) if ccol in df.columns \
        else df[[ucol]].dropna(subset=[ucol]).assign(**{ccol: 0})
    counts, map_rows = {}, []
    for element, raw_count in zip(sub[ucol], sub[ccol]):
        n = _count_value(raw_count)
        hits = match_sections(element, references)
        if not hits:
            rule = rule_identifier(str(element).strip())
            searched = (f"{rule[0]} {rule[1]}" if rule
                        else strip_parentheticals(element))
            map_rows.append({
                "Column": col, "Unique value": element, "Count": n,
                "Matched using": searched,
                "Matched section": "— no match —", "Match type": "",
            })
            continue
        for title, how, needle in hits:
            acc = counts.setdefault(title, [0, 0])
            acc[0] += n
            acc[1] += 1
            map_rows.append({
                "Column": col, "Unique value": element, "Count": n,
                "Matched using": needle,
                "Matched section": title, "Match type": how,
            })
    return counts, map_rows


def ref_map_sheet_name(base):
    """Name of the reference-match audit sheet for ``base`` (a run sheet or a
    unique-elements sheet). A leading ``Uniq `` is dropped so ``Uniq Sheet1``
    gives ``Ref Map Sheet1`` rather than ``Ref Map Uniq Sheet1``."""
    b = str(base)
    if b.lower().startswith("uniq "):
        b = b[len("Uniq "):]
    return f"Ref Map {b}"[:31]


def add_reference_sections(file_path, uniq_sheet, references, columns,
                           map_sheet=None):
    """Match the unique values of every **reference column** in ``uniq_sheet``
    (name contains a ``_REFERENCE_COLUMN_HINTS`` hint) against the reference
    **section titles**. Each value is matched by its ``(…)``-free form
    (``"AMC 25.21(g)(a)"`` -> ``"AMC 25.21"``, see :func:`match_sections`), and
    the result is recorded in two places:

    * ``uniq_sheet`` itself keeps its ``Unique_<col>`` / ``Count_<col>`` pairs
      untouched and gains, beside them, ``Section_<col>`` (matched title),
      ``Section_Count_<col>`` (how often that section was referenced — the sum
      of the matching values' counts) and ``Section_Matches_<col>`` (how many
      distinct unique values matched it), most-referenced first. The section
      tally therefore lives in the same sheet as the unique elements.
    * ``map_sheet`` (default :func:`ref_map_sheet_name`) receives one row per
      individual match — which value matched which section and in which
      direction — including the values that matched nothing.

    Shared by :func:`evaluate_uniques` (the run-sheet pipeline) and the AI
    Review *Unique elements* tab (any workbook/sheet), so both behave alike.
    Rebuilt on every call. Returns ``{"ref_map_sheet": …, "ref_sections":
    {col: {section: occurrences}}}``.
    """
    import pandas as pd

    df = pd.read_excel(file_path, sheet_name=uniq_sheet, engine="openpyxl")
    ref_blocks, ref_map_rows, ref_sections = [], [], {}
    for col in columns:
        if not is_reference_column(col) or f"Unique_{col}" not in df.columns:
            continue
        counts, map_rows = _reference_section_counts(df, col, references)
        ref_map_rows.extend(map_rows)
        if not counts:
            continue
        # Most-referenced section first, ties alphabetical.
        ordered = sorted(counts.items(), key=lambda kv: (-kv[1][0], str(kv[0])))
        ref_blocks.append(pd.DataFrame({
            f"Section_{col}": [t for t, _ in ordered],
            f"Section_Count_{col}": [v[0] for _, v in ordered],
            f"Section_Matches_{col}": [v[1] for _, v in ordered],
        }))
        ref_sections[col] = {t: v[0] for t, v in ordered}

    ref_map_sheet = None
    if ref_blocks:
        # Side by side with the existing Unique_/Count_ columns (the sheet is
        # already a ragged, column-per-block layout; concat pads with NaN).
        _write_sheet(file_path, uniq_sheet, pd.concat([df] + ref_blocks, axis=1))
    if ref_map_rows:
        ref_map_sheet = (map_sheet or ref_map_sheet_name(uniq_sheet))[:31]
        _write_sheet(file_path, ref_map_sheet, pd.DataFrame(
            ref_map_rows,
            columns=["Column", "Unique value", "Count", "Matched using",
                     "Matched section", "Match type"]))
    return {"ref_map_sheet": ref_map_sheet, "ref_sections": ref_sections}


def evaluate_uniques(file_path, run_sheet, references, columns):
    """Identify the unique generated values via save_unique_elements_to_new_sheet
    (refreshing ``Uniq <run sheet>``), then ground every unique element against
    the combined reference text into ``Uniq Eval <run sheet>``.

    **Reference columns** (name contains a ``_REFERENCE_COLUMN_HINTS`` hint,
    e.g. ``References``) additionally get their unique values matched against the
    reference **section titles** (:func:`match_sections`, substring match in
    either direction). For those columns the ``Uniq <run sheet>`` sheet keeps its
    existing ``Unique_<col>`` / ``Count_<col>`` pair untouched and simply gains
    three more columns beside them:

        ``Section_<col>``          the matched section title
        ``Section_Count_<col>``    how often that section was referenced
                                   (the sum of the matching values' counts)
        ``Section_Matches_<col>``  how many distinct unique values matched it

    so the section tally lives in the same sheet as the unique elements, sorted
    most-referenced first. Every individual match is also written to a separate
    ``Ref Map <run sheet>`` audit sheet (which value matched which section, and
    in which direction), including the values that matched nothing.
    """
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

    # Reference columns: unique values -> section titles, added to the Uniq
    # sheet itself plus the Ref Map audit sheet (shared with the Unique
    # elements tab). Rebuilt on every call, like every sheet here: the Uniq
    # sheet is regenerated above, so the Section_* columns are re-derived
    # rather than accumulated.
    ref = add_reference_sections(file_path, uniq_sheet, references, columns,
                                 map_sheet=ref_map_sheet_name(run_sheet))
    ref_map_sheet = ref["ref_map_sheet"]
    ref_sections = ref["ref_sections"]

    blocks, uniq_coverage = [], {}
    for col in columns:
        ucol, ccol = f"Unique_{col}", f"Count_{col}"
        if ucol not in df.columns:
            continue
        sub = df[[ucol, ccol]].dropna(subset=[ucol])
        grounded = [is_grounded(el, combined) for el in sub[ucol]]
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
        return {"uniq_sheet": uniq_sheet, "ref_map_sheet": ref_map_sheet,
                "ref_sections": ref_sections}

    target_df = pd.concat(blocks, axis=1)
    uniq_eval_sheet = f"Uniq Eval {run_sheet}"[:31]
    _write_sheet(file_path, uniq_eval_sheet, target_df)
    return {"uniq_sheet": uniq_sheet, "uniq_eval_sheet": uniq_eval_sheet,
            "uniq_coverage": uniq_coverage, "ref_map_sheet": ref_map_sheet,
            "ref_sections": ref_sections}


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