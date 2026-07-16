"""
data_extraction/evaluation/embedding_metrics.py

Semantic embedding metrics for the column-evaluation pipeline.

These complement the existing lexical / surface metrics in column_evaluator
(Jaccard, ROUGE, BLEU, Levenshtein, similarity ratio, WER). Where those measure
*word overlap*, these measure *meaning overlap*: the already-divided reference
sentences and candidate items are embedded with the shared backends in
llm_utils and compared in vector space, so a correct-but-reworded extraction
still scores well.

Two families are exposed, both computed from the SAME embedding batch so a
single embedding call serves both (important because llm_utils.get_embedding
is rate-limited and sleeps between calls):

  * Embedding cosine similarity  -> semantic_cosine()
        best-match cosine of each candidate item to its closest reference
        sentence (precision-oriented, mirrors the "best reference sentence"
        tracking already used in the evaluator), plus the reverse (recall),
        their F1, and a centroid cosine.

  * BERTScore (greedy P / R / F1) -> bertscore()
        greedy token/item matching over a cosine matrix. On the local backend
        this uses true per-TOKEN contextual embeddings (real BERTScore); on the
        remote backend, whose /embeddings endpoint only returns one pooled
        vector per string, it falls back to item-level greedy matching and
        flags that in the result.

Conventions followed (per the surrounding codebase):
  * Backends are reached ONLY through llm_utils - no new HTTP / model code.
  * Inputs are the lists the caller already produced with column_evaluator's
    canonical splitters; this module never re-splits text.
  * Metrics are "guarded": an empty side, a failed embedding call or a missing
    backend yields None instead of raising, so one bad row never aborts a batch.
  * numpy / llm_utils are imported lazily inside functions (no heavy import at
    module load), matching the lazy-import style used elsewhere.
"""

from __future__ import annotations

from typing import Dict, List, Optional, Sequence

# ---------------------------------------------------------------------------
# Backend / service resolution
# ---------------------------------------------------------------------------
# Candidate import paths for llm_utils. The first that imports wins and is
# cached. Trim / reorder this list to match your package layout (llm_utils
# lives next to review_panel.py). A single explicit entry is fine, e.g.
#   _LLM_UTILS_CANDIDATES = ("data_extraction.easa.llm_utils",)
_LLM_UTILS_CANDIDATES = (
    "data_extraction.easa.llm_utils",
    "data_extraction.review.llm_utils",
    "data_extraction.llm_utils",
    "llm_utils",
)

_LLM_UTILS = None
_INJECTED_LLM_UTILS = None

# Default remote service used when backend="api" and no service is given.
DEFAULT_REMOTE_SERVICE = "DLR Ollama"


def set_llm_utils(module):
    """Inject the already-imported llm_utils module (e.g. the caller's
    ``from . import llm_utils``). Preferred over path guessing: it makes this
    module work regardless of where llm_utils lives in the package tree."""
    global _INJECTED_LLM_UTILS, _LLM_UTILS
    _INJECTED_LLM_UTILS = module
    _LLM_UTILS = module


def _llm_utils():
    """Return the llm_utils module: the injected one if set, else the first of
    the candidate import paths that imports (cached)."""
    global _LLM_UTILS
    if _INJECTED_LLM_UTILS is not None:
        return _INJECTED_LLM_UTILS
    if _LLM_UTILS is not None:
        return _LLM_UTILS
    import importlib
    last_exc = None
    for name in _LLM_UTILS_CANDIDATES:
        try:
            _LLM_UTILS = importlib.import_module(name)
            return _LLM_UTILS
        except Exception as exc:  # noqa: BLE001 - keep trying the next candidate
            last_exc = exc
    raise ImportError(
        "Could not import llm_utils from any of "
        f"{_LLM_UTILS_CANDIDATES}. Set _LLM_UTILS_CANDIDATES in "
        f"embedding_metrics.py to the correct path. Last error: {last_exc}"
    )


def _np():
    import numpy as np  # local import: numpy stays out of module import time
    return np


# ---------------------------------------------------------------------------
# Low-level embedding helpers (pooled, one vector per input string)
# ---------------------------------------------------------------------------
# Failure ladder for embedding calls
# ---------------------------------------------------------------------------
# A local embedding call that dies almost always dies of a CUDA OOM: the whole
# pool of strings is padded to the longest one and pushed through an 8B model in
# a single forward pass, so one long sentence can blow the batch. A remote call
# dies for the mirror-image reason (payload too large / gateway timeout).
# Either way the fix is the same: send fewer strings per call. So a failing call
# is retried down a ladder --
#     1. the whole pool in one call        (fast path, what succeeds normally)
#     2. batches of EMBED_FALLBACK_BATCH   (default 10)
#     3. one string at a time              (last resort)
# -- keeping whatever each rung manages to embed. Strings that fail even alone
# are reported and simply left un-embedded: their metrics stay blank (None)
# rather than being faked with a zero vector, which would score as 0.0
# similarity and read as "totally dissimilar" instead of "not measured".
EMBED_FALLBACK_BATCH = 10

# Set False to silence the per-rung retry messages.
EMBED_VERBOSE = True


def _log(msg):
    if EMBED_VERBOSE:
        print(f"⚠️ embedding_metrics: {msg}")


def _cuda_recover():
    """Release cached CUDA blocks between rungs. After an OOM the allocator is
    still holding the cache that caused it, so a retry has no chance until it is
    emptied. No-op when torch/CUDA are absent (e.g. the API backend)."""
    try:
        import torch
        if torch.cuda.is_available():
            torch.cuda.empty_cache()
            torch.cuda.synchronize()
    except Exception:  # noqa: BLE001 - best effort only
        pass


def _ladder(items, call_fn, batch=None, what="embedding"):
    """Run ``call_fn(chunk)`` (returning a sequence aligned with ``chunk``) down
    the whole-pool -> batches -> individual ladder described above.

    Returns ``(results, failed)`` where ``results`` is a list aligned with
    ``items``, holding ``None`` for anything that could not be embedded, and
    ``failed`` is how many of those there are.
    """
    n = len(items)
    out = [None] * n
    batch = max(1, int(batch or EMBED_FALLBACK_BATCH))

    # Rung 1: the whole pool in one call.
    try:
        res = call_fn(items)
        return [res[i] for i in range(n)], 0
    except Exception as exc:  # noqa: BLE001
        nxt = (f"retrying in batches of {batch}" if n > batch
               else "retrying one string at a time")
        _log(f"{what}: pool of {n} string(s) failed ({exc}); {nxt}")
    _cuda_recover()

    # Rung 2: batches of `batch` (skipped when the pool is already that small --
    # rung 1 has just tried exactly that call).
    if n > batch:
        for start in range(0, n, batch):
            chunk = items[start:start + batch]
            try:
                res = call_fn(chunk)
                for k in range(len(chunk)):
                    out[start + k] = res[k]
            except Exception as exc:  # noqa: BLE001 - next batch may still work
                _log(f"{what}: batch {start + 1}-{start + len(chunk)} failed "
                     f"({exc})")
                _cuda_recover()
        if all(o is not None for o in out):
            return out, 0
        _log(f"{what}: retrying the remaining string(s) one at a time")

    # Rung 3: one string at a time, for whatever is still missing.
    for i, item in enumerate(items):
        if out[i] is not None:
            continue
        try:
            res = call_fn([item])
            out[i] = res[0]
        except Exception as exc:  # noqa: BLE001 - guarded: skip this string
            _log(f"{what}: string #{i + 1} failed on its own ({exc})")
            _cuda_recover()

    failed = sum(1 for o in out if o is None)
    if failed:
        _log(f"{what}: {failed} of {n} string(s) could not be embedded — "
             f"their metrics stay blank")
    return out, failed


# ---------------------------------------------------------------------------
def _l2_normalize(mat):
    """Row-wise L2 normalise an (N, dim) array; zero rows are left as zeros."""
    np = _np()
    mat = np.asarray(mat, dtype=np.float32)
    if mat.ndim == 1:
        mat = mat.reshape(1, -1)
    norms = np.linalg.norm(mat, axis=1, keepdims=True)
    norms[norms == 0] = 1.0
    return mat / norms


def embed_texts(
    texts: Sequence[str],
    backend: str = "api",
    service: Optional[str] = None,
    model: Optional[str] = None,
    api_key: Optional[str] = None,
):
    """
    Embed a list of strings into an L2-normalised (N, dim) numpy array, using
    the shared llm_utils backends.

    backend:
        "api"   -> llm_utils.get_embedding(list, service, model, api_key)
                   (remote OpenAI-compatible /embeddings; the whole list is sent
                   in ONE request).
        "local" -> llm_utils.vectorize_strings_local(list)
                   (local HuggingFace model; already L2-normalised).

    A failing call is retried down the ladder in :func:`_ladder` (whole pool ->
    batches of EMBED_FALLBACK_BATCH -> one string at a time), so a CUDA OOM on a
    big pool degrades into smaller passes instead of losing the row.

    Returns an empty (0, 0) array for empty input. Rows are aligned with the
    non-blank inputs, so this raises RuntimeError if any string could not be
    embedded even alone — the caller is expected to guard it (see
    :func:`semantic_scores`). Use :func:`embed_texts_map` when partial results
    are usable.
    """
    np = _np()
    items = [str(t) for t in (texts or []) if str(t).strip() != ""]
    if not items:
        return np.zeros((0, 0), dtype=np.float32)

    vectors, failed = _ladder(
        items, lambda chunk: _embed_call(chunk, backend, service, model, api_key),
        what="pooled embedding")
    if failed:
        raise RuntimeError(
            f"{failed} of {len(items)} string(s) could not be embedded "
            f"(backend={backend}); see the messages above.")
    return np.stack(vectors).astype(np.float32)


def embed_texts_map(
    texts: Sequence[str],
    backend: str = "api",
    service: Optional[str] = None,
    model: Optional[str] = None,
    api_key: Optional[str] = None,
):
    """``{text: unit vector}`` for every string that could be embedded, using the
    same ladder as :func:`embed_texts` but **keeping partial results**: a string
    that fails even on its own is simply absent from the mapping, so the caller
    can score everything else and leave that one blank.

    Duplicates are embedded once. This is what the per-run cache in
    column_evaluator fills itself from.
    """
    items = list(dict.fromkeys(
        str(t) for t in (texts or []) if str(t).strip() != ""))
    if not items:
        return {}
    vectors, _failed = _ladder(
        items, lambda chunk: _embed_call(chunk, backend, service, model, api_key),
        what="pooled embedding")
    return {t: v for t, v in zip(items, vectors) if v is not None}


def _embed_call(items, backend, service, model, api_key):
    """One raw embedding call for ``items`` — no retry, no guarding. The unit of
    work the ladder retries at ever-smaller sizes."""
    lu = _llm_utils()
    b = (backend or "api").lower()

    if b in ("local", "l", "hf"):
        vectors = lu.vectorize_strings_local(items)  # (N, dim), normalised
        return _l2_normalize(vectors)

    # Remote API path -----------------------------------------------------
    svc = service or DEFAULT_REMOTE_SERVICE
    result = lu.get_embedding(items, service=svc, model=model, api_key=api_key)
    vectors = (result or {}).get("embeddings") or []
    if len(vectors) != len(items):
        # Defensive: keep vector<->text alignment strict.
        raise ValueError(
            f"Embedding count mismatch from {svc}: got {len(vectors)} "
            f"vector(s) for {len(items)} input(s)."
        )
    return _l2_normalize(vectors)


def _cosine_matrix(a, b):
    """Cosine-similarity matrix between two L2-normalised sets: (a rows x b rows)."""
    np = _np()
    if a.size == 0 or b.size == 0:
        return np.zeros((a.shape[0], b.shape[0]), dtype=np.float32)
    return np.clip(a @ b.T, -1.0, 1.0)


def _greedy_prf(cos):
    """
    Greedy precision / recall / F1 over a cosine matrix `cos` of shape
    (n_candidate, n_reference):

        precision = mean over candidate rows of max-over-references
                    ("each generated item is supported by some reference")
        recall    = mean over reference cols of max-over-candidates
                    ("each reference sentence is covered by some item")
        f1        = harmonic mean of the two.

    Returns (precision, recall, f1) as plain floats, or (None, None, None)
    when either side is empty.
    """
    if cos.size == 0 or cos.shape[0] == 0 or cos.shape[1] == 0:
        return None, None, None
    precision = float(cos.max(axis=1).mean())   # candidate -> best reference
    recall = float(cos.max(axis=0).mean())      # reference -> best candidate
    denom = precision + recall
    f1 = float(2.0 * precision * recall / denom) if denom > 0 else 0.0
    return precision, recall, f1


def _round(x, ndigits=4):
    return None if x is None else round(float(x), ndigits)


# --- public wrappers over the primitives (used by column_evaluator) --------- #
def cosine(a, b):
    """Cosine similarity between two 1-D vectors (each L2-normalised first)."""
    np = _np()
    a = _l2_normalize(np.asarray(a, dtype=np.float32).reshape(1, -1))
    b = _l2_normalize(np.asarray(b, dtype=np.float32).reshape(1, -1))
    return float(np.clip(a @ b.T, -1.0, 1.0)[0, 0])


def cosine_matrix(cand, ref):
    """Public cosine matrix between two L2-normalised sets (n_cand x n_ref)."""
    return _cosine_matrix(cand, ref)


def greedy_bertscore(cand_vectors, ref_vectors):
    """Greedy BERTScore-style (precision, recall, f1) over a cosine matrix built
    from candidate vs reference vectors (rows are unit vectors). Works for token
    matrices (true token-level BERTScore) or item/sentence matrices (item-level
    approximation). Returns (None, None, None) when either side is empty."""
    return _greedy_prf(_cosine_matrix(cand_vectors, ref_vectors))


def token_embed_local(texts, max_length: int = 512):
    """Public alias for per-token local embeddings (true token-level BERTScore).
    Returns a list of L2-normalised (n_tokens_i, dim) arrays, one per input."""
    return _token_embed_local(texts, max_length=max_length)


# ---------------------------------------------------------------------------
# Public metric: embedding cosine similarity
# ---------------------------------------------------------------------------
def semantic_cosine(
    reference_items: Sequence[str],
    candidate_items: Sequence[str],
    backend: str = "api",
    service: Optional[str] = None,
    model: Optional[str] = None,
    api_key: Optional[str] = None,
    ndigits: int = 4,
) -> Dict[str, Optional[float]]:
    """
    Semantic cosine similarity between the divided reference sentences and the
    divided candidate items.

    Returns a dict of guarded floats (None when not computable):
        embedding_cosine           primary, precision-oriented:
                                   mean best-match of each candidate item to its
                                   closest reference sentence.
        embedding_cosine_recall    mean best-match of each reference to its
                                   closest candidate item.
        embedding_cosine_f1        harmonic mean of the two above.
        embedding_cosine_centroid  cosine between the mean (centroid) vectors of
                                   each side - a single robust summary number.
    """
    keys = ("embedding_cosine", "embedding_cosine_recall",
            "embedding_cosine_f1", "embedding_cosine_centroid")
    ref = embed_texts(reference_items, backend, service, model, api_key)
    cand = embed_texts(candidate_items, backend, service, model, api_key)
    if ref.shape[0] == 0 or cand.shape[0] == 0:
        return {k: None for k in keys}

    cos = _cosine_matrix(cand, ref)
    precision, recall, f1 = _greedy_prf(cos)

    np = _np()
    cand_centroid = _l2_normalize(cand.mean(axis=0, keepdims=True))
    ref_centroid = _l2_normalize(ref.mean(axis=0, keepdims=True))
    centroid = float(np.clip(cand_centroid @ ref_centroid.T, -1.0, 1.0)[0, 0])

    return {
        "embedding_cosine": _round(precision, ndigits),
        "embedding_cosine_recall": _round(recall, ndigits),
        "embedding_cosine_f1": _round(f1, ndigits),
        "embedding_cosine_centroid": _round(centroid, ndigits),
    }


# ---------------------------------------------------------------------------
# Local per-token embeddings (for TRUE token-level BERTScore)
# ---------------------------------------------------------------------------
def _token_embed_call(texts: Sequence[str], max_length: int = 512):
    """One raw per-token embedding call — the unit the ladder retries."""
    np = _np()
    import torch
    import torch.nn.functional as F

    lu = _llm_utils()
    # Reuse the cached loader if present, else the public loader.
    if hasattr(lu, "_get_embedding_model_and_tokenizer"):
        tokenizer, model = lu._get_embedding_model_and_tokenizer()
    else:
        tokenizer, model = lu.load_embedding_model_and_tokenizer(
            getattr(lu, "local_embedding_model_dir", None))

    items = [str(t) for t in texts]
    with torch.inference_mode():
        batch = tokenizer(
            items,
            padding=True,
            truncation=True,
            max_length=max_length,
            return_tensors="pt",
        )
        batch = {k: v.to(model.device) for k, v in batch.items()}
        outputs = model(**batch)
        hidden = outputs.last_hidden_state           # (B, T, dim)
        hidden = F.normalize(hidden, p=2, dim=2)     # normalise each token
        mask = batch["attention_mask"].bool()        # (B, T)

        per_text = []
        hidden = hidden.detach().cpu().numpy().astype(np.float32)
        mask_np = mask.detach().cpu().numpy()
        for i in range(hidden.shape[0]):
            valid = hidden[i][mask_np[i]]            # (n_tokens_i, dim)
            per_text.append(valid if valid.size else np.zeros((0, hidden.shape[2]),
                                                              dtype=np.float32))
    return per_text


def _token_embed_local(texts: Sequence[str], max_length: int = 512):
    """
    Per-TOKEN contextual embeddings for each string, using the SAME local model
    that llm_utils loads (reuses its cached model/tokenizer - no reload, no new
    weights). Returns a list of L2-normalised (n_tokens_i, dim) arrays, one per
    input string, with padding removed.

    This is what makes real (token-level) BERTScore possible on the local
    backend. The remote /embeddings endpoint only returns pooled vectors, so it
    cannot supply this.

    Token embeddings are far heavier than pooled ones (every token is kept, not
    just one vector per string), so this is the most OOM-prone call in the
    module and runs down the same ladder as :func:`embed_texts`. A string that
    fails even alone gets an empty (0, dim) array, which the callers already
    treat as "no token data" and fall back to the pooled cosine for.
    """
    np = _np()
    items = [str(t) for t in texts]
    if not items:
        return []
    mats, _failed = _ladder(
        items, lambda chunk: _token_embed_call(chunk, max_length=max_length),
        what="token embedding")
    return [m if m is not None else np.zeros((0, 0), dtype=np.float32)
            for m in mats]


def _stack_tokens(per_text):
    """Concatenate a list of (n_i, dim) token arrays into one (sum n_i, dim)."""
    np = _np()
    non_empty = [a for a in per_text if a.size]
    if not non_empty:
        return np.zeros((0, 0), dtype=np.float32)
    return np.concatenate(non_empty, axis=0)


# ---------------------------------------------------------------------------
# Public metric: BERTScore (greedy P / R / F1)
# ---------------------------------------------------------------------------
def bertscore(
    reference_items: Sequence[str],
    candidate_items: Sequence[str],
    backend: str = "api",
    service: Optional[str] = None,
    model: Optional[str] = None,
    api_key: Optional[str] = None,
    granularity: Optional[str] = None,
    ndigits: int = 4,
) -> Dict[str, Optional[float]]:
    """
    BERTScore-style greedy Precision / Recall / F1.

    granularity:
        "token" -> real token-level BERTScore (local backend only): embeds every
                   token of the joined reference / candidate text and greedily
                   matches token-to-token.
        "item"  -> item/sentence-level approximation over pooled vectors (works
                   on any backend; identical numbers to the embedding_cosine
                   P/R/F1, since it is the same cosine matrix).
        None    -> "token" when backend is local, else "item".

    Returns:
        bertscore_p, bertscore_r, bertscore_f1  (guarded floats, None if empty),
        plus bertscore_granularity so the sheet records which mode was used.

    Not implemented (documented deliberately): IDF weighting and baseline
    rescaling from the published BERTScore. Add IDF here if you later want the
    weighted variant.
    """
    keys = ("bertscore_p", "bertscore_r", "bertscore_f1")
    b = (backend or "api").lower()
    is_local = b in ("local", "l", "hf")
    mode = granularity or ("token" if is_local else "item")

    if mode == "token" and not is_local:
        # Remote endpoint cannot give token embeddings; degrade transparently.
        mode = "item"

    if mode == "token":
        ref_tokens = _stack_tokens(_token_embed_local(reference_items))
        cand_tokens = _stack_tokens(_token_embed_local(candidate_items))
        if ref_tokens.shape[0] == 0 or cand_tokens.shape[0] == 0:
            out = {k: None for k in keys}
            out["bertscore_granularity"] = "token"
            return out
        cos = _cosine_matrix(cand_tokens, ref_tokens)
    else:
        ref = embed_texts(reference_items, backend, service, model, api_key)
        cand = embed_texts(candidate_items, backend, service, model, api_key)
        if ref.shape[0] == 0 or cand.shape[0] == 0:
            out = {k: None for k in keys}
            out["bertscore_granularity"] = "item"
            return out
        cos = _cosine_matrix(cand, ref)

    precision, recall, f1 = _greedy_prf(cos)
    return {
        "bertscore_p": _round(precision, ndigits),
        "bertscore_r": _round(recall, ndigits),
        "bertscore_f1": _round(f1, ndigits),
        "bertscore_granularity": mode,
    }


# ---------------------------------------------------------------------------
# Combined entry point (one embedding batch for both families)
# ---------------------------------------------------------------------------
# Metric-name -> family, so column_evaluator can dispatch by the same string
# names it already uses ("jaccard", "rouge1", ...).
COSINE_METRIC_NAMES = (
    "embedding_cosine",
    "embedding_cosine_recall",
    "embedding_cosine_f1",
    "embedding_cosine_centroid",
)
BERTSCORE_METRIC_NAMES = ("bertscore_p", "bertscore_r", "bertscore_f1")
ALL_METRIC_NAMES = COSINE_METRIC_NAMES + BERTSCORE_METRIC_NAMES


def semantic_scores(
    reference_items: Sequence[str],
    candidate_items: Sequence[str],
    metrics: Optional[Sequence[str]] = None,
    backend: str = "api",
    service: Optional[str] = None,
    model: Optional[str] = None,
    api_key: Optional[str] = None,
    ndigits: int = 4,
) -> Dict[str, Optional[float]]:
    """
    Compute the requested semantic metrics for one reference/candidate pair and
    return a flat {metric_name: value} dict, ready to merge into an evaluation
    row. Only the requested families are computed. Guarded end-to-end: any
    backend failure returns None for the requested keys (with the error under
    'embedding_error') rather than raising, so a batch is never aborted.

    `metrics` uses the same string names as the rest of the evaluator; pass None
    to compute the cosine family only (the cheapest, single-vector-per-item
    family). Pass any bertscore_* name to also compute BERTScore.
    """
    wanted = list(metrics) if metrics else list(COSINE_METRIC_NAMES[:1])
    want_cos = any(m in COSINE_METRIC_NAMES for m in wanted)
    want_bert = any(m in BERTSCORE_METRIC_NAMES for m in wanted)

    out: Dict[str, Optional[float]] = {}
    try:
        if want_cos:
            out.update(semantic_cosine(reference_items, candidate_items,
                                       backend, service, model, api_key, ndigits))
        if want_bert:
            out.update(bertscore(reference_items, candidate_items,
                                 backend, service, model, api_key,
                                 granularity=None, ndigits=ndigits))
    except Exception as exc:  # noqa: BLE001 - guarded metric: never abort a batch
        for m in wanted:
            out.setdefault(m, None)
        out["embedding_error"] = str(exc)

    # Return only the requested metric names (plus any diagnostic extras).
    filtered = {k: v for k, v in out.items() if k in wanted}
    for extra in ("bertscore_granularity", "embedding_error"):
        if extra in out:
            filtered[extra] = out[extra]
    return filtered


if __name__ == "__main__":
    # Tiny smoke test against the remote backend. Requires a stored API key and
    # network access; safe to delete.
    ref = [
        "The operator shall maintain a record of each inspection.",
        "Records must be kept for at least three years.",
    ]
    cand = [
        "Inspection records are kept by the operator.",
        "Records are retained for a minimum of 3 years.",
    ]
    print("cosine:", semantic_cosine(ref, cand, backend="api"))
    print("bertscore(item):", bertscore(ref, cand, backend="api", granularity="item"))