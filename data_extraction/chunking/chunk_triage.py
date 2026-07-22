"""
data_extraction.chunking.chunk_triage
=====================================

Automatic **sorting / triage** of the chunks Docling produced, so the manual
chunk review only has to look at what is genuinely uncertain.

The design follows the alr pipeline's section processing
(``alr.data_analysis.Pdf_File_processor.process_pdf_sections`` Step D →
``_process_sections``): collect the headings Docling attached to the chunks,
work out which of them are *real* section headings, and then fold every chunk
whose heading is not real into the preceding valid section. Here that idea is
extended with a **Table of Contents** reference, which regulatory documents
almost always carry:

1. :func:`collect_headings` — the ordered, unique heading list (alr's
   ``unique_headings_list``).
2. :func:`find_toc` / :func:`parse_toc_entries` — locate a Table of Contents
   among the chunks and parse its entries. When one exists it is the
   authoritative list of sections, so no LLM call is needed at all.
3. :func:`refine_headings` — decide the valid headings, in order of
   preference: the TOC, else an LLM pass over the heading list (alr's
   ``SYSTEM_PROMPT_Heading_identifier`` approach), else deterministic rules.
4. :func:`triage_chunks` — per chunk propose ``log`` / ``skip`` / ``review``
   with the heading to use and a human-readable reason.
5. :func:`build_merged_sections` — the ``merged_headings`` payload the rest of
   the pipeline (Section Review, AI Review) already consumes.

Nothing here touches Tkinter and the LLM is injected as a callable, so the
whole engine runs (and is tested) headless. Every decision carries a
``reason`` because the review UI proposes rather than silently applies.
"""

from __future__ import annotations

import ast
import re
from collections import Counter
from difflib import SequenceMatcher

# Docling doc-item labels that are page furniture rather than content.
BOILERPLATE_LABELS = frozenset({"page_header", "page_footer"})

# Docling labels for Table-of-Contents / index entries. Docling routinely
# shreds a printed TOC into many dot-leader fragments and tags them
# ``document_index``; that whole run is navigation, not content, so it is
# always skipped (and never re-used as the section reference — printed TOCs
# come out too mangled to trust).
INDEX_LABELS = frozenset({"document_index"})

# Structural labels: these are short by nature, so the "too short to be
# content" rule must not throw them away.
HEADING_LABELS = frozenset({"title", "subtitle", "section_header"})

# Labels that carry real content (anything else is looked at more carefully).
CONTENT_LABELS = frozenset({
    "text", "paragraph", "list_item", "section_header", "title", "subtitle",
    "caption", "table", "code", "formula", "checkbox_selected",
    "checkbox_unselected",
})

# A chunk shorter than this (after normalisation) carries no usable content.
MIN_CONTENT_CHARS = 25

# A parsed TOC is only trusted as the section reference when it yields at
# least this many entries carrying page numbers. Docling frequently mangles a
# printed TOC into a handful of dot-leader fragments; below this it is
# rejected and heading refinement falls through to the LLM / rules.
MIN_TOC_ENTRIES = 4

# Identical text repeated on at least this many pages is running furniture.
BOILERPLATE_REPEAT_PAGES = 4

# SequenceMatcher ratio above which two headings are "the same heading"
# (alr uses 0.9 in merge_content_by_refined_headings).
HEADING_MATCH_RATIO = 0.88

# Headings look like these in regulatory / technical documents. A heading
# carrying explicit numbering is always kept (alr's MANDATORY RULE) — but
# "numbering" must be more than a bare number, or every page number and date
# in the document would be promoted to a section heading.

# Complete by itself: 'Part 21', 'Subpart A', 'AMC1 ORO.GEN.200', 'CS 25.301'.
_DESIGNATOR_RE = re.compile(
    r"""^\s*(part|subpart|section|chapter|annex|appendix|article
             |amc\d*|gm\d*|cs)\s+[\w.\-]+""",
    re.IGNORECASE | re.VERBOSE,
)

# Coded clause identifiers: an uppercase prefix (optionally dotted with more
# uppercase segments) then a dotted/spaced number, as regulatory paragraphs
# are labelled — 'EHPS.10 Scope', 'EHPS 480 …', 'ORO.GEN.200',
# 'CAT.OP.MPA.100'. The trailing digits keep this off plain ALL-CAPS words
# like 'GENERAL' or 'SCOPE'.
_CODE_DESIGNATOR_RE = re.compile(r"^\s*[A-Z]{2,}(\.[A-Z]+)*[.\s]\s*\d+")

# Numbering followed by actual title text: '1. Scope', '2.1 Applicability'.
_NUMBERED_TITLE_RE = re.compile(
    r"""^\s*(
        \d+(\.\d+)*\.?  |             # 1.   2.1   1.2.3.
        [IVXLC]+\.      |             # IV.
        [A-Z]\.                       # A.
    )\s+\S""",
    re.VERBOSE,
)

# A bare section marker with no title — only trusted for roman numerals and
# single letters, never for digits (those are page numbers far more often).
_MARKER_RE = re.compile(r"^\s*([IVXLC]+\.|[A-Z]\.)\s*$")

# 2021-04-01, 01/02/2021 — a date is never a heading.
_DATE_RE = re.compile(r"^\s*\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}\s*$")

# A Table of Contents chunk announces itself with one of these titles.
_TOC_TITLE_RE = re.compile(
    r"^\s*(table\s+of\s+contents|contents|inhaltsverzeichnis|index)\s*$",
    re.IGNORECASE,
)

# "1.2 Scope .......... 14"  /  "Subpart A — General    12"
_TOC_ENTRY_RE = re.compile(
    r"^\s*(?P<title>\S.*?)"          # the entry title
    r"(?:[\s.·•_-]{2,}|\s{2,}|\t+)"  # dot leaders / wide gap
    r"(?P<page>\d{1,4})\s*$"         # the page number
)

# The system prompt for the LLM heading pass. Adapted from alr's
# SYSTEM_PROMPT_Heading_identifier for regulatory documents.
SYSTEM_PROMPT_HEADING_IDENTIFIER = (
    "You are a Professional Document Structure Analyst specializing in "
    "regulatory, legal and technical publications (EASA, FAA, ICAO, ISO, "
    "RTCA and similar).\n"
    "Your task is to evaluate a list of extracted document headings and keep "
    "only the valid, meaningful section titles.\n\n"

    "### MANDATORY RULE - DO NOT REMOVE:\n"
    "- Any heading that includes section numbering or a regulatory "
    "designator (e.g., '1.', '2.1', 'Part 21', 'Subpart A', 'Section III', "
    "'Appendix A', 'AMC1 ORO.GEN.200', 'CS 25.301', 'GM1 145.A.30') MUST be "
    "kept. These are always valid.\n\n"

    "### Evaluation Criteria for 'Valid Headings':\n"
    "1. Structural Relevance: keep headings that define a logical section "
    "(e.g., 'Scope', 'Applicability', 'Definitions', 'Requirements').\n"
    "2. Publication Standards: keep standard unnumbered headings (e.g., "
    "'Foreword', 'Preamble', 'References', 'Appendix').\n"
    "3. Content-Specific Titles: keep descriptive titles naming the subject "
    "matter.\n"
    "4. Invalid: remove noise such as page numbers, running headers and "
    "footers, document/revision identifiers, dates on their own, fragments "
    "of sentences, and 'Table of Contents'.\n\n"

    "### Strict Formatting Instructions:\n"
    "- Maintain the exact original order of the headings.\n"
    "- Keep the original wording; you may normalize casing only.\n"
    "- Ensure no two headings in the output list are identical.\n"
    "- DO NOT use markdown code blocks (no ```).\n"
    "- DO NOT include variable names or commentary.\n"
    "- Output Format: your response must be ONLY a valid Python list of "
    "strings.\n"
    "- Example of desired output: ['Heading 1', 'Heading 2']\n\n"

    "### Example:\n"
    "Input: ['1. Scope', 'Page 4 of 88', '2. Applicability', "
    "'Doc. No. EASA-2021-04', 'Definitions']\n"
    "Output: ['1. Scope', '2. Applicability', 'Definitions']"
)


# --------------------------------------------------------------- helpers -- #
def chunk_heading(chunk):
    """The first heading Docling attached to a chunk, or "" when it has none.
    Handles both the cached dict shape and a live Docling chunk (mirrors
    alr's ``extract_chunk_heading``)."""
    if isinstance(chunk, dict):
        headings = chunk.get("heading") or chunk.get("meta", {}).get("headings") or []
    else:
        meta = getattr(chunk, "meta", None)
        headings = getattr(meta, "headings", []) if meta is not None else []
    if isinstance(headings, str):
        headings = [headings]
    for h in headings or []:
        if str(h).strip():
            return str(h).strip()
    return ""


def chunk_text(chunk):
    """The chunk's text, from either shape."""
    if isinstance(chunk, dict):
        return str(chunk.get("chunk_text") or chunk.get("text") or "")
    return str(getattr(chunk, "text", "") or "")


def chunk_labels(chunk):
    """The Docling doc-item labels of a chunk, lower-cased."""
    if isinstance(chunk, dict):
        labels = chunk.get("type_of_docitem") or []
        if not labels:
            labels = [d.get("label") for d
                      in chunk.get("meta", {}).get("doc_items", []) or []]
    else:
        meta = getattr(chunk, "meta", None)
        items = getattr(meta, "doc_items", []) if meta is not None else []
        labels = [getattr(i, "label", None) for i in items]
    return [str(x).strip().lower() for x in labels if x]


def chunk_pages(chunk):
    """The page numbers a chunk spans."""
    if isinstance(chunk, dict):
        pages = chunk.get("page_num")
        if pages is None:
            pages = chunk.get("meta", {}).get("page_numbers")
    else:
        meta = getattr(chunk, "meta", None)
        pages = getattr(meta, "page_numbers", None) if meta is not None else None
    if pages is None:
        return []
    return list(pages) if isinstance(pages, (list, tuple, set)) else [pages]


def normalize_heading(text):
    """Casefolded, whitespace-collapsed, punctuation-trimmed form used for
    comparing headings — so 'Part 21 ', 'PART 21' and 'Part  21.' match."""
    raw = re.sub(r"\s+", " ", str(text or "")).strip()
    return raw.strip(" .:;-–—_·").casefold()


def headings_match(a, b, ratio=HEADING_MATCH_RATIO):
    """True when two headings denote the same section: equal once normalised,
    one contained in the other, or similar above ``ratio`` (alr compares with
    SequenceMatcher the same way)."""
    na, nb = normalize_heading(a), normalize_heading(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    # A Docling heading is often the TOC entry without its page number, or
    # the TOC entry carries extra numbering — accept containment both ways.
    if len(na) >= 4 and len(nb) >= 4 and (na in nb or nb in na):
        return True
    return SequenceMatcher(None, na, nb).ratio() >= ratio


def is_numbered_heading(text):
    """True when a heading carries section numbering or a regulatory
    designator — always a valid heading (alr's MANDATORY RULE).

    A bare number ('12') or a date is NOT one: those are page numbers, and
    treating them as headings would scatter the document into noise sections.
    Digits must be followed by title text; roman numerals and single letters
    are accepted standalone because they are unambiguous section markers."""
    raw = str(text or "").strip()
    if not raw or _DATE_RE.match(raw):
        return False
    return bool(_DESIGNATOR_RE.match(raw) or _CODE_DESIGNATOR_RE.match(raw)
                or _NUMBERED_TITLE_RE.match(raw) or _MARKER_RE.match(raw))


def collect_headings(chunks):
    """The ordered, unique, non-empty headings of the chunks — the input to
    the TOC/LLM refinement (alr's ``unique_headings_list``)."""
    out, seen = [], set()
    for chunk in chunks:
        heading = chunk_heading(chunk)
        if not heading:
            continue
        key = normalize_heading(heading)
        if key and key not in seen:
            seen.add(key)
            out.append(heading)
    return out


# ------------------------------------------------------ table of contents -- #
def parse_toc_entries(text):
    """Parse Table-of-Contents lines into ``[{title, page}]``. Recognises the
    usual ``Title .......... 12`` form as well as entries separated from the
    page number by a wide gap or a tab. Lines without a page number are kept
    when they look like a numbered heading, since some TOCs wrap."""
    entries = []
    for line in str(text or "").splitlines():
        line = line.strip()
        if not line or _TOC_TITLE_RE.match(line):
            continue
        match = _TOC_ENTRY_RE.match(line)
        if match:
            title = match.group("title").strip(" .·•_-\t")
            if title:
                entries.append({"title": title,
                                "page": int(match.group("page"))})
            continue
        # No page number: only trust it when it is clearly a numbered heading.
        if is_numbered_heading(line) and len(line) <= 200:
            entries.append({"title": line.strip(" .·•_-\t"), "page": None})
    return entries


def find_toc(chunks):
    """Locate a Table of Contents among the chunks.

    Returns ``{"entries": [...], "chunk_indices": [...]}``; ``entries`` is
    empty when the document has no usable TOC. A chunk counts as TOC when its
    heading (or first line) announces a table of contents, and the entries are
    taken from that chunk plus the immediately following ones that keep
    parsing as TOC lines (long TOCs span several chunks)."""
    chunk_list = list(chunks)
    start = None
    for i, chunk in enumerate(chunk_list):
        heading = chunk_heading(chunk)
        first_line = chunk_text(chunk).strip().splitlines()[:1]
        first_line = first_line[0] if first_line else ""
        if _TOC_TITLE_RE.match(heading) or _TOC_TITLE_RE.match(first_line):
            start = i
            break
    if start is None:
        return {"entries": [], "chunk_indices": []}

    entries, indices = [], []
    for i in range(start, len(chunk_list)):
        parsed = parse_toc_entries(chunk_text(chunk_list[i]))
        if i == start:
            entries.extend(parsed)
            indices.append(i)
            continue
        # Continue only while the chunk still looks like TOC lines.
        lines = [ln for ln in chunk_text(chunk_list[i]).splitlines() if ln.strip()]
        if not lines or len(parsed) < max(2, len(lines) // 2):
            break
        entries.extend(parsed)
        indices.append(i)

    # De-duplicate, preserving order.
    unique, seen = [], set()
    for entry in entries:
        key = normalize_heading(entry["title"])
        if key and key not in seen:
            seen.add(key)
            unique.append(entry)
    return {"entries": unique, "chunk_indices": indices}


# ------------------------------------------------------ heading refinement -- #
def parse_llm_list(text):
    """Pull a Python list of strings out of an LLM reply (alr's
    ``process_llm_refined_structure``): take the outermost [...] and
    literal_eval it. Returns [] when the reply is unusable."""
    raw = str(text or "")
    match = re.search(r"(\[.*\])", raw, re.DOTALL)
    candidate = match.group(1).strip() if match else raw.strip()
    try:
        value = ast.literal_eval(candidate)
    except (ValueError, SyntaxError):
        return []
    if isinstance(value, (list, tuple)):
        return [str(v).strip() for v in value if str(v).strip()]
    return []


def refine_headings_rules(headings):
    """Deterministic fallback: keep numbered/designated headings and plausible
    title-like ones, drop obvious page furniture. Used when there is neither a
    TOC nor an LLM."""
    valid = []
    for heading in headings:
        text = str(heading).strip()
        if not text or _TOC_TITLE_RE.match(text):
            continue
        if is_numbered_heading(text):
            valid.append(text)
            continue
        stripped = text.strip(" .·•_-")
        # Pure numbers / dates / "Page 4 of 88" style furniture.
        if re.fullmatch(r"[\d\s./\-]+", stripped):
            continue
        if re.match(r"^page\s+\d+", stripped, re.IGNORECASE):
            continue
        if 3 <= len(stripped) <= 160:
            valid.append(text)
    return valid


def refine_headings(headings, toc=None, llm=None, log=None):
    """Work out the document's valid section headings.

    Preference order — TOC first (authoritative and free), then the LLM pass
    over the collected headings, then deterministic rules. Returns
    ``{"valid": [...], "source": "toc"|"llm"|"rules", "note": str}``."""
    def _log(msg):
        if log:
            log(msg)

    toc_entries = (toc or {}).get("entries") or []
    with_pages = [e for e in toc_entries if e.get("page") is not None]
    # Only trust the TOC when it parsed into a real list of entries — a couple
    # of dot-leader fragments are a mangled TOC, not a section reference.
    if len(toc_entries) >= MIN_TOC_ENTRIES and len(with_pages) >= MIN_TOC_ENTRIES:
        valid = [e["title"] for e in toc_entries]
        _log(f"Table of Contents found: {len(valid)} entries — using it as "
             "the section reference (no LLM call needed).")
        return {"valid": valid, "source": "toc",
                "note": f"{len(valid)} TOC entries"}
    if toc_entries:
        _log(f"A Table of Contents was detected but only {len(toc_entries)} "
             "usable entrie(s) parsed (Docling likely mangled it) — its pages "
             "are skipped but headings are validated another way.")

    if llm and headings:
        _log(f"No Table of Contents — asking the LLM to validate "
             f"{len(headings)} extracted heading(s).")
        try:
            reply = llm(f"List of Headings: {headings}",
                        SYSTEM_PROMPT_HEADING_IDENTIFIER)
        except Exception as exc:  # noqa: BLE001 - fall back, never crash triage
            _log(f"[WARN] LLM heading refinement failed ({exc}) — "
                 "falling back to rules.")
        else:
            valid = parse_llm_list(reply)
            # Never let the model drop a numbered heading (alr's mandatory
            # rule) — re-insert any it removed, keeping document order.
            if valid:
                kept = list(valid)
                for heading in headings:
                    if is_numbered_heading(heading) and not any(
                            headings_match(heading, v) for v in kept):
                        kept.append(heading)
                ordered = [h for h in headings
                           if any(headings_match(h, k) for k in kept)]
                # keep LLM-only rephrasings that match nothing in the input
                for k in kept:
                    if not any(headings_match(k, h) for h in ordered):
                        ordered.append(k)
                _log(f"LLM kept {len(ordered)} of {len(headings)} heading(s).")
                return {"valid": ordered, "source": "llm",
                        "note": f"{len(ordered)}/{len(headings)} kept by LLM"}
            _log("[WARN] LLM reply was not a usable list — falling back to "
                 "rules.")

    valid = refine_headings_rules(headings)
    _log(f"Rule-based heading refinement kept {len(valid)} of "
         f"{len(headings)} heading(s).")
    return {"valid": valid, "source": "rules",
            "note": f"{len(valid)}/{len(headings)} kept by rules"}


# ------------------------------------------------------------ boilerplate -- #
def find_boilerplate(chunks, min_pages=BOILERPLATE_REPEAT_PAGES):
    """Normalised texts that repeat across at least ``min_pages`` different
    pages — running headers/footers, document IDs and the like."""
    seen = {}
    for chunk in chunks:
        text = re.sub(r"\s+", " ", chunk_text(chunk)).strip().casefold()
        if not text or len(text) > 200:
            continue          # long text repeating is real duplicated content
        pages = chunk_pages(chunk)
        seen.setdefault(text, set()).update(pages or [None])
    return {text for text, pages in seen.items() if len(pages) >= min_pages}


# ----------------------------------------------------------------- triage -- #
def triage_chunks(chunks, valid_headings=None, toc=None, boilerplate=None):
    """Propose a decision for every chunk.

    Each proposal is ``{chunk_index, action, heading, reason, confidence,
    page_num, labels, text}`` where ``action`` is:

    * ``skip``   — page furniture, empty, or the Table of Contents itself;
    * ``log``    — keep, under ``heading``;
    * ``review`` — needs a human (no heading could be determined).

    A chunk whose own heading is not a valid section heading is **not**
    thrown away: its content is logged under the last valid heading, which is
    what the reviewer used to do by hand with *Use Prev Heading* and what
    alr's ``merge_content_by_refined_headings`` does automatically."""
    valid_headings = list(valid_headings or [])
    boilerplate = boilerplate if boilerplate is not None else find_boilerplate(chunks)
    toc_indices = set((toc or {}).get("chunk_indices") or [])

    proposals = []
    current_heading = ""
    for i, chunk in enumerate(chunks):
        text = chunk_text(chunk)
        norm = re.sub(r"\s+", " ", text).strip()
        labels = chunk_labels(chunk)
        heading = chunk_heading(chunk)
        index = chunk.get("chunk_index", i + 1) if isinstance(chunk, dict) else i + 1

        action = heading_out = reason = None
        confidence = "high"

        if labels and any(lbl in INDEX_LABELS for lbl in labels):
            action, heading_out = "skip", heading
            reason = "Table of Contents / index entry"
        elif i in toc_indices:
            action, heading_out = "skip", heading
            reason = "Table of Contents page (used as the section reference)"
        elif not norm:
            action, heading_out, reason = "skip", heading, "empty chunk"
        elif labels and all(lbl in BOILERPLATE_LABELS for lbl in labels):
            action, heading_out = "skip", heading
            reason = f"page furniture ({', '.join(sorted(set(labels)))})"
        elif norm.casefold() in boilerplate:
            action, heading_out = "skip", heading
            reason = "repeated on many pages (running header/footer)"
        elif (len(norm) < MIN_CONTENT_CHARS and not is_numbered_heading(norm)
                and not set(labels) & HEADING_LABELS):
            action, heading_out = "skip", heading
            reason = f"too short to be content ({len(norm)} chars)"
            confidence = "medium"
        else:
            # Content. Decide which heading it belongs under.
            match = next((v for v in valid_headings
                          if heading and headings_match(heading, v)), None)
            if match:
                current_heading = match
                action, heading_out = "log", match
                reason = ("heading matches the Table of Contents"
                          if (toc or {}).get("entries")
                          else "recognised section heading")
            elif heading and is_numbered_heading(heading):
                # Numbered headings are always valid, even if the TOC/LLM
                # list missed them (alr's mandatory rule).
                current_heading = heading
                action, heading_out = "log", heading
                reason = "numbered/designated heading — always valid"
            elif current_heading:
                action, heading_out = "log", current_heading
                reason = (f"heading '{heading}' is not a section heading — "
                          f"content folded into '{current_heading}'"
                          if heading else
                          f"no heading — continues '{current_heading}'")
                confidence = "medium"
            else:
                action, heading_out = "review", heading
                reason = "no valid heading yet (nothing precedes it)"
                confidence = "low"

        proposals.append({
            "chunk_index": index,
            "position": i,
            "action": action,
            "heading": heading_out or "",
            "original_heading": heading,
            "reason": reason,
            "confidence": confidence,
            "page_num": chunk_pages(chunk),
            "labels": labels,
            "text": text,
        })
    return proposals


def triage_summary(proposals):
    """Counts per action plus how many need a human — what the UI shows as
    'you only have to look at N of M chunks'."""
    counts = Counter(p["action"] for p in proposals)
    total = len(proposals)
    return {
        "total": total,
        "log": counts.get("log", 0),
        "skip": counts.get("skip", 0),
        "review": counts.get("review", 0),
        "low_confidence": sum(1 for p in proposals
                              if p["confidence"] == "low"),
        "auto_handled_pct": (round(100.0 * (counts.get("log", 0)
                                            + counts.get("skip", 0)) / total, 1)
                             if total else 0.0),
    }


def analyze_chunks(chunks, llm=None, log=None):
    """Run the whole triage: headings → TOC → refinement → per-chunk
    proposals. Returns ``{proposals, summary, toc, refinement, headings}``."""
    chunks = list(chunks)
    headings = collect_headings(chunks)
    toc = find_toc(chunks)
    refinement = refine_headings(headings, toc=toc, llm=llm, log=log)
    boilerplate = find_boilerplate(chunks)
    if log and boilerplate:
        log(f"{len(boilerplate)} repeated text block(s) detected as running "
            "headers/footers.")
    proposals = triage_chunks(chunks, refinement["valid"], toc=toc,
                              boilerplate=boilerplate)
    summary = triage_summary(proposals)
    if log:
        log(f"Triage: {summary['log']} to log, {summary['skip']} to skip, "
            f"{summary['review']} need review "
            f"({summary['auto_handled_pct']}% decided automatically).")
    return {"proposals": proposals, "summary": summary, "toc": toc,
            "refinement": refinement, "headings": headings}


# ------------------------------------------------------------- the output -- #
def build_merged_sections(decisions):
    """Build the ``merged_headings`` payload from accepted decisions.

    ``decisions`` are proposal dicts (possibly edited by the reviewer) with an
    ``action``; ``skip`` entries are excluded, exactly as the manual review
    does. Grouping is by **normalised** heading, which is what stops
    'Part 21' and 'Part 21 ' becoming two sections — the original spelling of
    the first occurrence is kept for display."""
    merged, order = {}, []
    for entry in decisions:
        if entry.get("action") == "skip":
            continue
        heading = entry.get("heading") or ""
        key = normalize_heading(heading)
        if key not in merged:
            merged[key] = {
                "heading": [heading] if heading else [],
                "merged_text": entry.get("text", ""),
                "chunk_indices": [entry.get("chunk_index")],
                "types_of_docitem": list(entry.get("labels") or []),
                "page_nums": list(entry.get("page_num") or []),
            }
            order.append(key)
        else:
            target = merged[key]
            target["merged_text"] += "\n\n" + entry.get("text", "")
            target["chunk_indices"].append(entry.get("chunk_index"))
            for label in entry.get("labels") or []:
                if label not in target["types_of_docitem"]:
                    target["types_of_docitem"].append(label)
            for page in entry.get("page_num") or []:
                if page not in target["page_nums"]:
                    target["page_nums"].append(page)
    return [merged[k] for k in order]


def build_output_payload(decisions):
    """The full review-output payload: ``merged_headings`` (what Section
    Review and AI Review read) plus ``raw_session_history`` in the shape the
    manual review has always written, so resuming and the existing readers
    keep working."""
    history = []
    for entry in decisions:
        history.append({
            "chunk_index": entry.get("chunk_index"),
            "status": ("skipped" if entry.get("action") == "skip"
                       else entry.get("status") or "logged (auto)"),
            "heading": [entry["heading"]] if entry.get("heading") else [],
            "chunk_text": entry.get("text", ""),
            "type_of_docitem": list(entry.get("labels") or []),
            "page_num": list(entry.get("page_num") or []),
            "triage_reason": entry.get("reason", ""),
        })
    return {"merged_headings": build_merged_sections(decisions),
            "raw_session_history": history}
