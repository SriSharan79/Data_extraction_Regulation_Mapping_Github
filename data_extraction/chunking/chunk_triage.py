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
2. :func:`resolve_toc` — the Table of Contents, taken from the first source
   that yields a usable one:

   a. the PDF's embedded outline (bookmarks) — :func:`toc_from_pdf`;
   b. the printed TOC parsed from the PDF's page text — :func:`toc_from_pdf`;
   c. the TOC reassembled from the Docling chunks — :func:`find_toc`;
   d. the TOC read from an **extracted table** sitting on the contents pages
      — :func:`toc_pages_from_chunks` + :func:`toc_from_tables`. Regulatory
      documents very often lay their contents out as a table, which Docling
      then extracts as a table rather than as text; the table is only trusted
      when it passes :func:`table_looks_like_toc` (enough rows carrying a page
      number, page numbers running forwards, titles that read like titles);
   e. the LLM asked to pull the TOC out of the contents chunks —
      :func:`toc_from_llm`.

   When a TOC exists it is the authoritative list of sections, so the LLM
   heading check below is not needed at all.
3. :func:`refine_headings` — decide the valid headings, in order of
   preference: the TOC, else an LLM pass over the heading list (alr's
   ``SYSTEM_PROMPT_Heading_identifier`` approach), else deterministic rules.
4. :func:`triage_chunks` — per chunk propose ``log`` / ``skip`` / ``review``
   with the heading to use and a human-readable reason.
5. :func:`resolve_merges` — apply the reviewer's ``merge`` decisions: a chunk
   marked ``merge`` is folded into the text of the chunk above it (the
   nearest preceding chunk that is kept in its own right), optionally with
   its own heading written as a line before the text.
6. :func:`build_merged_sections` — the ``merged_headings`` payload the rest of
   the pipeline (Section Review, AI Review) already consumes.

Nothing here touches Tkinter and the LLM is injected as a callable, so the
whole engine runs (and is tested) headless. Every decision carries a
``reason`` because the review UI proposes rather than silently applies.
"""

from __future__ import annotations

import ast
import json
import os
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

# An extracted table is only accepted as the Table of Contents when at least
# this share of its non-empty rows yields a title AND a page number. A table
# that Docling shredded (merged cells, split columns, OCR noise) falls below
# this and is reported as "did not extract cleanly" rather than used.
MIN_TOC_TABLE_ROW_RATIO = 0.6

# ... and its page numbers must run forwards: at least this share of
# consecutive pairs must be non-decreasing. A real contents list always does;
# a data table of unrelated numbers almost never does.
MIN_TOC_TABLE_ORDER_RATIO = 0.8

# How far into the document the contents pages are looked for when the chunks
# do not announce a Table of Contents themselves.
MAX_TOC_SCAN_PAGE = 25

# How much contents text is sent to the LLM in the last-resort TOC pass.
MAX_TOC_LLM_CHARS = 12000

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

# The system prompt for the last-resort TOC pass: the reader is given the raw
# text of the pages that look like the contents and has to return the entries.
# Kept separate from the heading check above because the task is different —
# this one reads a contents list, it does not judge a list of headings.
SYSTEM_PROMPT_TOC_EXTRACTOR = (
    "You are a Professional Document Structure Analyst specializing in "
    "regulatory, legal and technical publications (EASA, FAA, ICAO, ISO, "
    "RTCA and similar).\n"
    "You are given the raw extracted text of the pages that appear to hold a "
    "document's Table of Contents. The extraction is imperfect: dot leaders, "
    "column breaks and page numbers may be mangled or split across lines.\n"
    "Your task is to reconstruct the Table of Contents entries.\n\n"

    "### Rules:\n"
    "- Return ONE entry per section listed in the contents, in the order they "
    "appear.\n"
    "- Keep the section numbering / regulatory designator as part of the "
    "title (e.g. '1.2 Scope', 'Subpart A - General', 'AMC1 ORO.GEN.200').\n"
    "- 'page' is the page number printed next to the entry; use null when the "
    "entry has none. Never invent a page number.\n"
    "- Repair obvious extraction damage (a title split over two lines is ONE "
    "entry), but never invent entries that are not in the text.\n"
    "- Do NOT include the words 'Table of Contents'/'Contents' themselves, "
    "running headers/footers, document or revision identifiers, or page "
    "numbers on their own.\n"
    "- If the text is not a table of contents at all, return an empty list.\n\n"

    "### Strict Formatting Instructions:\n"
    "- DO NOT use markdown code blocks (no ```).\n"
    "- DO NOT include commentary, explanations or variable names.\n"
    "- Output Format: your response must be ONLY a valid JSON array of "
    "objects, each with exactly the keys \"title\" and \"page\".\n"
    "- Example of desired output: "
    "[{\"title\": \"1. Scope\", \"page\": 4}, "
    "{\"title\": \"2. Applicability\", \"page\": 7}]"
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


def toc_from_pdf(pdf_path, max_scan_pages=20):
    """The Table of Contents straight from the **PDF itself** — more reliable
    than reassembling it from Docling chunks (which routinely shreds a
    printed TOC into dot-leader fragments).

    Two sources, in order:

    1. the embedded outline (PDF bookmarks) via PyMuPDF — authoritative when
       present;
    2. the printed TOC parsed from the first ``max_scan_pages`` pages' text
       (a page announcing "Table of Contents"/"Contents" plus the following
       pages that keep parsing as TOC lines).

    Returns ``{"entries": [{title, page, ...}], "source":
    "pdf-bookmarks"|"pdf-text"|None}`` — entries ``[]`` when the PDF offers
    no usable TOC (the caller then falls back to the LLM heading check)."""
    if not pdf_path:
        return {"entries": [], "source": None}
    try:
        import fitz  # PyMuPDF — optional; without it chunk-level detection remains
    except Exception:  # noqa: BLE001
        return {"entries": [], "source": None}
    try:
        doc = fitz.open(str(pdf_path))
    except Exception:  # noqa: BLE001
        return {"entries": [], "source": None}
    try:
        # 1. embedded outline (bookmarks)
        try:
            outline = doc.get_toc() or []
        except Exception:  # noqa: BLE001
            outline = []
        entries = [{"title": str(title).strip(), "page": page, "level": level}
                   for level, title, page in outline if str(title).strip()]
        if len(entries) >= MIN_TOC_ENTRIES:
            return {"entries": entries, "source": "pdf-bookmarks"}

        # 2. printed TOC in the page text
        page_count = getattr(doc, "page_count", 0) or len(doc)
        start = None
        for i in range(min(page_count, max_scan_pages)):
            lines = (doc[i].get_text() or "").splitlines()
            if any(_TOC_TITLE_RE.match(line.strip()) for line in lines[:15]):
                start = i
                break
        found = []
        if start is not None:
            for i in range(start, min(page_count, start + 10)):
                parsed = parse_toc_entries(doc[i].get_text() or "")
                if i > start and len(parsed) < 3:
                    break     # the TOC ended on the previous page
                found.extend(parsed)
        unique, seen = [], set()
        for entry in found:
            key = normalize_heading(entry["title"])
            if key and key not in seen:
                seen.add(key)
                unique.append(entry)
        if len([e for e in unique if e.get("page") is not None]) >= MIN_TOC_ENTRIES:
            return {"entries": unique, "source": "pdf-text"}
        return {"entries": [], "source": None}
    finally:
        try:
            doc.close()
        except Exception:  # noqa: BLE001
            pass


# ------------------------------------------- toc from the extracted tables -- #
def toc_is_usable(entries, source=None):
    """Whether a parsed TOC is solid enough to be *the* section reference.

    Page numbers are what separates a real contents list from a handful of
    dot-leader fragments, so they are required — except for sources that are
    structural rather than parsed (the PDF's own bookmarks) or that have
    already reconstructed the list (the LLM pass), where a contents list
    without printed page numbers is still a contents list."""
    entries = entries or []
    if len(entries) < MIN_TOC_ENTRIES:
        return False
    if str(source or "") in ("pdf-bookmarks", "llm"):
        return True
    with_pages = [e for e in entries if e.get("page") is not None]
    return len(with_pages) >= MIN_TOC_ENTRIES


def toc_pages_from_chunks(chunks, toc=None, max_page=MAX_TOC_SCAN_PAGE):
    """The document pages the printed Table of Contents sits on, worked out
    from the chunks — which is what makes it possible to go and look at the
    **tables** that were extracted from exactly those pages.

    Two signals: the pages of the chunks :func:`find_toc` identified as the
    contents run, and the pages of chunks Docling labelled ``document_index``.
    The label alone is not trusted deep into a document (a back-of-book index
    carries it too), so without an announced contents run only the first
    ``max_page`` pages are considered."""
    chunk_list = list(chunks or [])
    if not chunk_list:
        return []
    toc = toc if toc is not None else find_toc(chunk_list)

    pages = set()
    for i in toc.get("chunk_indices") or []:
        if 0 <= i < len(chunk_list):
            pages.update(p for p in chunk_pages(chunk_list[i]) if p is not None)

    labelled = set()
    for chunk in chunk_list:
        labels = chunk_labels(chunk)
        if labels and any(lbl in INDEX_LABELS for lbl in labels):
            labelled.update(p for p in chunk_pages(chunk) if p is not None)
    if pages:
        # Keep only index pages that continue the announced contents run.
        low, high = min(pages), max(pages)
        pages.update(p for p in labelled if low - 1 <= p <= high + 2)
    else:
        pages.update(p for p in labelled
                     if isinstance(p, int) and p <= max_page)
    return sorted(p for p in pages if p is not None)


def _rows_from_workbook(path):
    """Rows of a saved table file (the .xlsx the extractor writes, or a CSV).
    Used when the cache entry carries no inline ``data``."""
    if not path or not os.path.exists(str(path)):
        return []
    try:
        import pandas as pd          # heavy; only needed for this fallback
    except Exception:  # noqa: BLE001
        return []
    try:
        if str(path).lower().endswith((".xlsx", ".xlsm", ".xls")):
            frame = pd.read_excel(path, header=None, dtype=str)
        else:
            frame = pd.read_csv(path, header=None, dtype=str)
    except Exception:  # noqa: BLE001
        return []
    return [["" if value is None else str(value) for value in row]
            for row in frame.values.tolist()]


def table_rows(table):
    """The rows of an extracted table as lists of strings.

    Reads the ``data`` records the extractor cached (``df.to_dict('records')``)
    and falls back to the saved workbook. The column names are offered as a
    first row too — but only when they read as a contents entry *with a page
    number*: that is how a first data line that got promoted into the header
    ('1. Scope | 4') is told apart from a genuine header line ('Section |
    Title') or from positional column indices."""
    if not isinstance(table, dict):
        return []
    rows = []
    records = table.get("data")
    if isinstance(records, list) and records:
        first = next((r for r in records if isinstance(r, dict)), None)
        if first is not None:
            header = [str(key) for key in first.keys()]
            entry = toc_entry_from_cells(header)
            if entry and entry.get("page") is not None:
                rows.append(header)
            for record in records:
                if isinstance(record, dict):
                    rows.append(["" if v is None else str(v)
                                 for v in record.values()])
        else:
            for record in records:
                if isinstance(record, (list, tuple)):
                    rows.append(["" if v is None else str(v) for v in record])
                elif record is not None:
                    rows.append([str(record)])
    if rows:
        return rows
    return _rows_from_workbook(table.get("csv_path") or table.get("path")
                               or table.get("table_path"))


def _clean_cell(value):
    """A table cell as comparable text; empties and pandas NaNs drop out."""
    text = re.sub(r"\s+", " ", str(value if value is not None else "")).strip()
    return "" if text.lower() in ("", "nan", "none", "null", "-", "—") else text


def _page_from_cell(value):
    """The page number a cell holds, or None — tolerates '14.', 'p. 14'."""
    text = _clean_cell(value)
    if not text:
        return None
    text = re.sub(r"^(page|p\.?|pg\.?|seite)\s*", "", text, flags=re.IGNORECASE)
    text = text.strip(" .·•_-")
    return int(text) if re.fullmatch(r"\d{1,4}", text) else None


def toc_entry_from_cells(cells):
    """One Table-of-Contents entry from one table row, or None.

    Handles the layouts a contents table actually comes in: ``['1.2 Scope',
    '14']``, ``['1.2', 'Scope', '14']`` and the single-cell ``'1.2 Scope
    ......... 14'``. The last cell is the page when it is a bare number;
    everything before it is the title.

    A row that fills several columns must carry a page number — that is what
    tells a real entry apart from a header line like ``['Section', 'Title']``.
    A row with only one filled cell may go without one, because that is how a
    spanning group row ('PART A — GENERAL') and a wrapped line look. Every
    title has to contain a letter, so column indices and stray numbers never
    become sections."""
    cells = [_clean_cell(c) for c in (cells or [])]
    cells = [c for c in cells if c]
    if not cells:
        return None

    if len(cells) == 1:
        match = _TOC_ENTRY_RE.match(cells[0])
        if match:
            title, page = match.group("title"), int(match.group("page"))
        elif is_numbered_heading(cells[0]) and len(cells[0]) <= 200:
            title, page = cells[0], None
        else:
            return None
    else:
        page = _page_from_cell(cells[-1])
        if page is not None:
            title = " ".join(cells[:-1])
        else:
            # The page can still be glued to the title with dot leaders.
            match = _TOC_ENTRY_RE.match(" ".join(cells))
            if not match:
                return None
            title, page = match.group("title"), int(match.group("page"))

    title = title.strip(" .·•_-\t")
    if not title or _TOC_TITLE_RE.match(title):
        return None
    if not re.search(r"[A-Za-z]", title):
        return None
    return {"title": title, "page": page}


def table_looks_like_toc(rows, min_entries=MIN_TOC_ENTRIES):
    """Did this table extract cleanly enough to be used as the contents?

    Returns ``(ok, entries, note)``. Three things have to hold, and the note
    says which one failed so the reviewer is told *why* a contents table was
    rejected instead of it silently not being used:

    1. enough rows turn into a title with a page number, and they are the
       majority of the rows — a table Docling shredded fails here;
    2. the page numbers run forwards — a contents list always does, a table
       of unrelated numbers does not;
    3. the titles read like titles rather than like measurements."""
    rows = list(rows or [])
    entries, non_empty = [], 0
    for row in rows:
        if not any(_clean_cell(c) for c in row):
            continue
        non_empty += 1
        entry = toc_entry_from_cells(row)
        if entry:
            entries.append(entry)
    with_pages = [e for e in entries if e["page"] is not None]

    if not non_empty:
        return False, entries, "the table is empty"
    if len(with_pages) < min_entries:
        return False, entries, (f"only {len(with_pages)} of {non_empty} row(s) "
                                "gave a title with a page number — the table "
                                "did not extract cleanly")
    ratio = len(with_pages) / float(non_empty)
    if ratio < MIN_TOC_TABLE_ROW_RATIO:
        return False, entries, (f"only {round(100 * ratio)}% of the rows gave a "
                                "title with a page number — the table did not "
                                "extract cleanly")
    pages = [e["page"] for e in with_pages]
    pairs = len(pages) - 1
    if pairs > 0:
        forward = sum(1 for a, b in zip(pages, pages[1:]) if b >= a)
        if forward / float(pairs) < MIN_TOC_TABLE_ORDER_RATIO:
            return False, entries, ("the page numbers do not run forwards — "
                                    "this is not a contents table")
    titled = sum(1 for e in with_pages
                 if re.search(r"[A-Za-z]", e["title"]) and len(e["title"]) >= 4)
    if titled < 0.5 * len(with_pages):
        return False, entries, ("the first column does not read like section "
                                "titles — this is not a contents table")
    return True, entries, (f"{len(with_pages)} entries from {non_empty} rows")


def toc_from_tables(tables, pages=None, max_page=MAX_TOC_SCAN_PAGE):
    """The Table of Contents read from the **extracted tables**.

    Regulatory documents very often typeset their contents as a table, and
    Docling then hands it over as a table rather than as text — which is why
    the text-level parsers above come back empty on exactly the documents
    that have the tidiest contents list. ``pages`` are the contents pages
    worked out by :func:`toc_pages_from_chunks`; without them only tables on
    the first ``max_page`` pages are looked at.

    Returns ``{"entries": [...], "source": "tables"|None, "tables": [report]}``
    where the report says, per table, whether it was used and why not."""
    report, collected = [], []
    wanted = {int(p) for p in (pages or []) if isinstance(p, int)}
    for i, table in enumerate(tables or []):
        if not isinstance(table, dict):
            continue
        raw_page = table.get("page_no")
        try:
            page_no = int(raw_page)
        except (TypeError, ValueError):
            page_no = None
        if wanted:
            if page_no not in wanted:
                continue
        elif page_no is None or page_no > max_page:
            continue

        rows = table_rows(table)
        if not rows:
            report.append({"table_index": table.get("table_index", i),
                           "page_no": raw_page, "used": False, "entries": 0,
                           "note": "the table holds no readable rows"})
            continue
        ok, entries, note = table_looks_like_toc(rows)
        report.append({"table_index": table.get("table_index", i),
                       "page_no": raw_page, "used": ok,
                       "entries": sum(1 for e in entries
                                      if e["page"] is not None),
                       "note": note,
                       "path": table.get("csv_path")})
        if ok:
            collected.extend(entries)

    unique, seen = [], set()
    for entry in collected:
        key = normalize_heading(entry["title"])
        if key and key not in seen:
            seen.add(key)
            unique.append(entry)
    return {"entries": unique, "source": "tables" if unique else None,
            "tables": report}


# ---------------------------------------------------------- toc from an llm -- #
def _as_page(value):
    """A TOC page number from whatever the LLM put in the field."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, int):
        return value if 0 < value < 10000 else None
    text = str(value).strip()
    return int(text) if re.fullmatch(r"\d{1,4}", text) else None


def parse_llm_toc(text):
    """Pull TOC entries out of an LLM reply: a JSON array of
    ``{"title", "page"}`` objects (what the prompt asks for), a Python list
    literal, or a plain list of contents lines — all three are accepted so a
    model that ignores the format instruction is still usable."""
    raw = str(text or "")
    match = re.search(r"(\[.*\])", raw, re.DOTALL)
    candidate = match.group(1).strip() if match else raw.strip()
    value = None
    try:
        value = json.loads(candidate)
    except Exception:  # noqa: BLE001
        try:
            value = ast.literal_eval(candidate)
        except (ValueError, SyntaxError):
            value = None
    if not isinstance(value, (list, tuple)):
        # Not a list at all: fall back to reading the reply as contents lines.
        return parse_toc_entries(raw)

    entries = []
    for item in value:
        if isinstance(item, dict):
            title = ""
            for key in ("title", "heading", "section", "name", "entry"):
                if str(item.get(key) or "").strip():
                    title = str(item[key]).strip()
                    break
            if not title:
                continue
            page = None
            for key in ("page", "page_no", "page_number", "pageNo"):
                if key in item:
                    page = _as_page(item.get(key))
                    if page is not None:
                        break
            title = title.strip(" .·•_-\t")
            if title and not _TOC_TITLE_RE.match(title):
                entries.append({"title": title, "page": page})
        elif isinstance(item, str) and item.strip():
            parsed = parse_toc_entries(item)
            if parsed:
                entries.extend(parsed)
            else:
                title = item.strip().strip(" .·•_-\t")
                if title and not _TOC_TITLE_RE.match(title):
                    entries.append({"title": title, "page": None})
    return entries


def toc_chunk_text(chunks, toc=None, max_chars=MAX_TOC_LLM_CHARS):
    """The raw text of the chunks that look like the contents — what the LLM
    is asked to reconstruct the TOC from.

    Everything sitting on a contents page is taken, not only the chunks
    :func:`find_toc` managed to string together: the run detection stops at
    the first chunk that no longer parses as TOC lines, which on a badly
    extracted contents page is the second chunk — exactly the case this stage
    exists for. Falls back to the opening chunks of the document, since a
    contents list is always at the front."""
    chunk_list = list(chunks or [])
    if not chunk_list:
        return ""
    toc = toc if toc is not None else find_toc(chunk_list)

    picked = set(i for i in (toc.get("chunk_indices") or [])
                 if 0 <= i < len(chunk_list))
    pages = set(toc_pages_from_chunks(chunk_list, toc=toc))
    for i, chunk in enumerate(chunk_list):
        if any(lbl in INDEX_LABELS for lbl in chunk_labels(chunk)):
            picked.add(i)
        elif pages and pages.intersection(chunk_pages(chunk)):
            picked.add(i)
    if not picked:
        picked = set(range(min(20, len(chunk_list))))

    out, size = [], 0
    for i in sorted(picked):
        text = chunk_text(chunk_list[i]).strip()
        if not text:
            continue
        out.append(text)
        size += len(text) + 2
        if size >= max_chars:
            break
    return "\n\n".join(out)[:max_chars]


def toc_from_llm(chunks, llm, toc=None, log=None):
    """Last resort: hand the contents text to the LLM and let it reconstruct
    the Table of Contents. Never raises — a failing model just returns no
    entries and the caller carries on with the heading-based refinement."""
    if not llm:
        return {"entries": [], "source": None,
                "note": "no LLM available for the TOC pass"}
    text = toc_chunk_text(chunks, toc=toc)
    if not text.strip():
        return {"entries": [], "source": None,
                "note": "no contents text to send"}
    if log:
        log(f"No Table of Contents from the PDF, the chunks or the extracted "
            f"tables — asking the LLM to read it out of {len(text)} "
            "characters of contents text.")
    try:
        reply = llm(f"Contents text:\n{text}", SYSTEM_PROMPT_TOC_EXTRACTOR)
    except Exception as exc:  # noqa: BLE001 - never break the triage
        if log:
            log(f"[WARN] LLM TOC extraction failed ({exc}).")
        return {"entries": [], "source": None, "note": f"LLM failed: {exc}"}
    entries = parse_llm_toc(reply)
    unique, seen = [], set()
    for entry in entries:
        key = normalize_heading(entry["title"])
        if key and key not in seen:
            seen.add(key)
            unique.append(entry)
    return {"entries": unique, "source": "llm" if unique else None,
            "note": f"{len(unique)} entries reconstructed by the LLM"}


# ------------------------------------------------------------ toc cascade -- #
TOC_SOURCE_LABEL = {
    "pdf-bookmarks": "the PDF's embedded outline (bookmarks)",
    "pdf-text": "the printed TOC parsed from the PDF page text",
    "chunks": "the TOC reassembled from the Docling chunks",
    "tables": "an extracted table on the contents pages",
    "llm": "the LLM reading the contents pages",
}


def resolve_toc(pdf_path=None, chunks=None, tables=None, llm=None, log=None,
                use_llm=True):
    """The document's Table of Contents, from the first source that yields a
    usable one — this is the single place the cascade is defined, so the
    triage and the *View TOC* button always agree.

    Order: the PDF's bookmarks, the printed TOC in the PDF text, the TOC
    reassembled from the chunks, an extracted **table** sitting on the
    contents pages, and finally the LLM asked to read the contents text.

    Returns the usual TOC dict — ``entries``, ``source`` and the
    ``chunk_indices`` of the chunks that ARE the printed contents (so they
    can be skipped) — plus ``note``, ``attempts`` (what every stage produced,
    for the log and the TOC viewer) and ``tables`` (per-table report of the
    table stage). When nothing is usable the richest partial result is
    returned, so the caller can still say 'a TOC was detected but it did not
    parse'."""
    chunk_list = list(chunks or [])
    chunk_toc = (find_toc(chunk_list) if chunk_list
                 else {"entries": [], "chunk_indices": []})
    chunk_indices = chunk_toc.get("chunk_indices") or []
    attempts, candidates = [], []

    def _stage(source, entries, note="", skipped=False):
        """Record one stage and say whether it settled the question."""
        entries = list(entries or [])
        usable = bool(entries) and toc_is_usable(entries, source)
        attempts.append({
            "source": source,
            "label": TOC_SOURCE_LABEL.get(source, source),
            "entries": len(entries),
            "with_pages": sum(1 for e in entries
                              if e.get("page") is not None),
            "usable": usable,
            "skipped": bool(skipped),
            "note": note,
        })
        if entries:
            candidates.append((source, entries))
        return usable

    def _log(msg):
        if log:
            log(msg)

    # a/b — the PDF itself.
    pdf_toc = toc_from_pdf(pdf_path) if pdf_path else {"entries": [],
                                                       "source": None}
    pdf_source = pdf_toc.get("source") or "pdf-text"
    if _stage(pdf_source, pdf_toc.get("entries"),
              "" if pdf_path else "no PDF given", skipped=not pdf_path):
        _log(f"Table of Contents read from {TOC_SOURCE_LABEL.get(pdf_source, pdf_source)} "
             f"({len(pdf_toc['entries'])} entries) — chunk headings are "
             "verified against it.")
        return {"entries": pdf_toc["entries"], "chunk_indices": chunk_indices,
                "source": pdf_source, "note": "read from the PDF",
                "attempts": attempts, "tables": []}

    # c — the chunks.
    if _stage("chunks", chunk_toc.get("entries"),
              "" if chunk_list else "no chunks given",
              skipped=not chunk_list):
        _log(f"Table of Contents reassembled from the chunks "
             f"({len(chunk_toc['entries'])} entries).")
        return {"entries": chunk_toc["entries"],
                "chunk_indices": chunk_indices, "source": "chunks",
                "note": "reassembled from the chunks", "attempts": attempts,
                "tables": []}

    # d — an extracted table on the contents pages.
    toc_pages = toc_pages_from_chunks(chunk_list, toc=chunk_toc)
    table_toc = {"entries": [], "tables": []}
    if tables:
        table_toc = toc_from_tables(tables, pages=toc_pages)
        note = (f"contents pages {toc_pages}" if toc_pages
                else f"no contents page identified — scanned tables on the "
                     f"first {MAX_TOC_SCAN_PAGE} pages")
        used = [t for t in table_toc["tables"] if t.get("used")]
        rejected = [t for t in table_toc["tables"] if not t.get("used")]
        if rejected:
            note += "; rejected: " + "; ".join(
                f"table {t['table_index']} on page {t['page_no']} "
                f"({t['note']})" for t in rejected[:4])
        settled = _stage("tables", table_toc["entries"], note)
        for t in used:
            _log(f"Contents table accepted: table {t['table_index']} on page "
                 f"{t['page_no']} — {t['note']}.")
        for t in rejected:
            _log(f"Contents table rejected: table {t['table_index']} on page "
                 f"{t['page_no']} — {t['note']}.")
    else:
        settled = _stage("tables", [], "no extracted tables available",
                         skipped=True)
    if settled:
        _log(f"Table of Contents read from an extracted table on the contents "
             f"pages ({len(table_toc['entries'])} entries).")
        return {"entries": table_toc["entries"],
                "chunk_indices": chunk_indices, "source": "tables",
                "note": (f"from the table(s) on page(s) "
                         f"{toc_pages or 'scanned at the front'}"),
                "attempts": attempts, "tables": table_toc["tables"]}

    # e — the LLM.
    llm_toc = {"entries": [], "note": ""}
    if use_llm and llm and chunk_list:
        llm_toc = toc_from_llm(chunk_list, llm, toc=chunk_toc, log=log)
        settled = _stage("llm", llm_toc.get("entries"),
                         llm_toc.get("note", ""))
    else:
        settled = _stage("llm", [],
                         "no LLM selected" if not llm else "no chunks given",
                         skipped=True)
    if settled:
        _log(f"Table of Contents reconstructed by the LLM "
             f"({len(llm_toc['entries'])} entries).")
        return {"entries": llm_toc["entries"], "chunk_indices": chunk_indices,
                "source": "llm", "note": "reconstructed by the LLM",
                "attempts": attempts, "tables": table_toc.get("tables") or []}

    # Nothing usable: hand back the richest partial so the caller can report
    # 'a TOC was detected but it did not parse' and fall back to headings.
    best_source, best_entries = "", []
    for source, entries in candidates:
        if len(entries) > len(best_entries):
            best_source, best_entries = source, entries
    _log("No usable Table of Contents from the PDF, the chunks, the extracted "
         "tables or the LLM — the heading list is validated instead.")
    return {"entries": best_entries, "chunk_indices": chunk_indices,
            "source": best_source or None,
            "note": "no usable Table of Contents", "attempts": attempts,
            "tables": table_toc.get("tables") or []}


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
    toc_source = (toc or {}).get("source")
    # Only trust the TOC when it parsed into a real list of entries — a couple
    # of dot-leader fragments are a mangled TOC, not a section reference.
    if toc_is_usable(toc_entries, toc_source):
        valid = [e["title"] for e in toc_entries]
        _log(f"Table of Contents found: {len(valid)} entries — using it as "
             "the section reference (no LLM heading check needed).")
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

    A fourth action, ``merge``, exists but is never *proposed* here: it is
    set by the reviewer to fold a chunk's text into the chunk above it (see
    :func:`resolve_merges`), because only a human can tell that a section was
    split across two chunks.

    A chunk whose own heading is not a valid section heading is **not**
    thrown away: its content is logged under the last valid heading, which is
    what the reviewer used to do by hand with *Use Prev Heading* and what
    alr's ``merge_content_by_refined_headings`` does automatically."""
    valid_headings = list(valid_headings or [])
    boilerplate = boilerplate if boilerplate is not None else find_boilerplate(chunks)
    toc_indices = set((toc or {}).get("chunk_indices") or [])
    toc_exists = bool((toc or {}).get("entries"))

    proposals = []
    current_heading = ""
    current_from_toc = None    # did the current section's heading match the TOC?
    for i, chunk in enumerate(chunks):
        text = chunk_text(chunk)
        norm = re.sub(r"\s+", " ", text).strip()
        labels = chunk_labels(chunk)
        heading = chunk_heading(chunk)
        index = chunk.get("chunk_index", i + 1) if isinstance(chunk, dict) else i + 1

        action = heading_out = reason = None
        confidence = "high"
        toc_match = None   # True/False only when a TOC exists and we keep

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
                current_from_toc = True if toc_exists else None
                action, heading_out = "log", match
                toc_match = True if toc_exists else None
                reason = ("heading matches the Table of Contents"
                          if toc_exists
                          else "recognised section heading")
            elif heading and is_numbered_heading(heading):
                # Numbered headings are always valid, even if the TOC/LLM
                # list missed them (alr's mandatory rule).
                current_heading = heading
                current_from_toc = False if toc_exists else None
                action, heading_out = "log", heading
                toc_match = False if toc_exists else None
                reason = ("numbered/designated heading — always valid"
                          + (" (NOT in the TOC)" if toc_exists else ""))
            elif current_heading:
                action, heading_out = "log", current_heading
                toc_match = current_from_toc
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
            "toc_match": toc_match,
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
    decided = (counts.get("log", 0) + counts.get("skip", 0)
               + counts.get("merge", 0))
    return {
        "total": total,
        "log": counts.get("log", 0),
        "skip": counts.get("skip", 0),
        "review": counts.get("review", 0),
        "merge": counts.get("merge", 0),
        "low_confidence": sum(1 for p in proposals
                              if p["confidence"] == "low"),
        "auto_handled_pct": (round(100.0 * decided / total, 1)
                             if total else 0.0),
    }


def analyze_chunks(chunks, llm=None, log=None, pdf_path=None, tables=None,
                   toc=None):
    """Run the whole triage: TOC → heading refinement → per-chunk proposals.
    Returns ``{proposals, summary, toc, refinement, headings}``.

    The Table of Contents is resolved by :func:`resolve_toc`, which tries, in
    order: the PDF's embedded outline, the printed TOC in the PDF text, the
    TOC reassembled from the chunks, an extracted **table** on the contents
    pages (``tables`` — the extractor's table records from the cache), and
    finally the LLM reading the contents text. Only a document that offers no
    Table of Contents at all falls through to the LLM heading check.

    ``toc`` short-circuits the cascade with an already-resolved TOC (the UI
    passes the one the reviewer looked at, so nothing is read twice)."""
    chunks = list(chunks)
    headings = collect_headings(chunks)
    if toc is None:
        toc = resolve_toc(pdf_path=pdf_path, chunks=chunks, tables=tables,
                          llm=llm, log=log)
    refinement = refine_headings(headings, toc=toc, llm=llm, log=log)
    if refinement["source"] == "toc" and toc.get("source"):
        refinement["note"] += f" ({toc['source']})"
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


# ---------------------------------------------------------------- merging -- #
# What separates a merged chunk's text from the text it is appended to.
MERGE_SEPARATOR = "\n\n"


def merge_block_text(entry):
    """The text a chunk marked ``merge`` contributes to the chunk above it.

    ``merge_add_heading`` decides whether the chunk's heading is written as a
    line before its text — the reviewer is asked this every time a merge is
    made, because a split paragraph must NOT get a heading inserted in the
    middle of it while a genuine sub-section usually must.
    ``merge_heading_text`` optionally overrides which heading is written."""
    body = str(entry.get("text") or "").rstrip()
    if not entry.get("merge_add_heading"):
        return body
    heading = str(entry.get("merge_heading_text")
                  or entry.get("heading") or "").strip()
    if not heading:
        return body
    return f"{heading}\n{body}" if body else heading


def find_merge_target(decisions, index):
    """The decision a chunk at list position ``index`` merges into, or None.

    That is the nearest preceding chunk that survives on its own: ``skip``
    chunks are passed over (they are not written at all) and so are other
    ``merge`` chunks (they are themselves folded further up), which is what
    makes a run of consecutive merges land in one and the same chunk."""
    for i in range(int(index) - 1, -1, -1):
        entry = decisions[i]
        if not isinstance(entry, dict):
            continue
        if entry.get("action") in ("skip", "merge"):
            continue
        return entry
    return None


def resolve_merges(decisions):
    """Apply every ``merge`` decision, without touching the input.

    Returns ``(resolved, merged_into)``: ``resolved`` is a copy of
    ``decisions`` in which each merge target has absorbed the text, labels
    and page numbers of the chunks merged into it (their indices land in its
    ``merged_from``), and ``merged_into`` maps the chunk index of every
    merged chunk to the chunk index it went into.

    Merged chunks stay in ``resolved`` — marked, with their own text intact —
    so the session history can record what the reviewer did and a resumed
    session re-applies the merge instead of duplicating the text. A chunk
    marked ``merge`` with nothing above it to merge into is kept as its own
    chunk rather than dropped, so no text is ever lost."""
    resolved, merged_into = [], {}
    target = None
    for entry in decisions:
        if not isinstance(entry, dict):
            continue
        current = dict(entry)
        # Own the mutable fields: the caller's proposals must not change.
        current["labels"] = list(entry.get("labels") or [])
        current["page_num"] = list(entry.get("page_num") or [])
        current["merged_from"] = list(entry.get("merged_from") or [])
        action = current.get("action")

        if action == "merge":
            if target is None:
                current["action"] = "log"
                current["merged_into"] = None
                current["merge_unresolved"] = True
                current["reason"] = ("marked to merge, but no kept chunk "
                                     "precedes it — left as its own chunk")
                resolved.append(current)
                target = current
                continue
            piece = merge_block_text(current)
            if piece.strip():
                base = str(target.get("text") or "").rstrip()
                target["text"] = (base + MERGE_SEPARATOR + piece
                                  if base else piece)
            for label in current["labels"]:
                if label not in target["labels"]:
                    target["labels"].append(label)
            for page in current["page_num"]:
                if page not in target["page_num"]:
                    target["page_num"].append(page)
            index = current.get("chunk_index")
            if index is not None and index not in target["merged_from"]:
                target["merged_from"].append(index)
            current["merged_into"] = target.get("chunk_index")
            merged_into[index] = target.get("chunk_index")
            resolved.append(current)
            continue

        resolved.append(current)
        if action != "skip":
            target = current
    return resolved, merged_into


# ------------------------------------------------------------- the output -- #
def build_merged_sections(decisions, resolve=True):
    """Build the ``merged_headings`` payload from accepted decisions.

    ``decisions`` are proposal dicts (possibly edited by the reviewer) with an
    ``action``; ``skip`` entries are excluded, exactly as the manual review
    does, and ``merge`` entries are excluded too because their text has
    already been folded into the chunk above them. Grouping is by
    **normalised** heading, which is what stops 'Part 21' and 'Part 21 '
    becoming two sections — the original spelling of the first occurrence is
    kept for display.

    ``resolve`` runs :func:`resolve_merges` first; pass ``False`` when the
    decisions were already resolved by the caller, so the merged text is not
    appended twice."""
    if resolve:
        decisions, _ = resolve_merges(decisions)
    merged, order = {}, []
    for entry in decisions:
        if entry.get("action") in ("skip", "merge"):
            continue
        heading = entry.get("heading") or ""
        key = normalize_heading(heading)
        indices = ([entry.get("chunk_index")]
                   + list(entry.get("merged_from") or []))
        if key not in merged:
            merged[key] = {
                "heading": [heading] if heading else [],
                "merged_text": entry.get("text", ""),
                "chunk_indices": indices,
                "types_of_docitem": list(entry.get("labels") or []),
                "page_nums": list(entry.get("page_num") or []),
            }
            order.append(key)
        else:
            target = merged[key]
            target["merged_text"] += "\n\n" + entry.get("text", "")
            target["chunk_indices"].extend(indices)
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
    keep working.

    The history stays a faithful **per-chunk** record: a merged chunk keeps
    its own text under the status ``merged`` and the chunk it went into keeps
    its own pre-merge text. Only ``merged_headings`` carries the combined
    text — that is what makes resuming a session idempotent (restoring the
    history and accepting again re-applies the merge instead of appending
    the same text a second time)."""
    resolved, merged_into = resolve_merges(decisions)
    history = []
    for entry in decisions:
        if not isinstance(entry, dict):
            continue
        action = entry.get("action")
        if action == "skip":
            status = "skipped"
        elif action == "merge":
            status = "merged"
        else:
            status = entry.get("status") or "logged (auto)"
        item = {
            "chunk_index": entry.get("chunk_index"),
            "status": status,
            "heading": [entry["heading"]] if entry.get("heading") else [],
            "chunk_text": entry.get("text", ""),
            "type_of_docitem": list(entry.get("labels") or []),
            "page_num": list(entry.get("page_num") or []),
            "triage_reason": entry.get("reason", ""),
        }
        if action == "merge":
            item["merged_into"] = merged_into.get(entry.get("chunk_index"))
            item["merge_add_heading"] = bool(entry.get("merge_add_heading"))
            if entry.get("merge_heading_text"):
                item["merge_heading_text"] = entry["merge_heading_text"]
        history.append(item)
    return {"merged_headings": build_merged_sections(resolved, resolve=False),
            "raw_session_history": history}