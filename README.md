# Data Extraction & Regulation Mapping

Tools for turning regulatory source documents (PDFs and EASA e-Rules XML) into
structured, human-curated data — chunked text, tables, images, section trees,
metadata indexes, and cross-reference graphs.

Everything lives in the `data_extraction` Python package and is reached through
two notebook launchers.

## Quick start

```bash
# from the repository root
python run_studio.py        # PDF/chunk tools: extraction & review (incl. cache),
                            # section review, AI review, PDF -> Markdown
python run_easa_studio.py   # EASA tools: XML extraction, structured-JSON review
```

Each launcher opens a tabbed window. Tabs load lazily and are fault-isolated: if
a tool's heavy dependency isn't installed, only that tab shows an error (or the
failure surfaces when you actually run it) — the other tabs keep working.

## Layout

```
data_extraction/
  chunking/    logic.py, chunk_review_ui.py, section_review_ui.py,
               ai_review_ui.py, table_image_extractor.py, workspace_config.py
  easa/        parser.py, graph_builder.py, run_main.py,
               extraction_ui.py, json_review_ui.py
  evaluation/  column_evaluator.py (AI-review output evaluation) +
               Lexical_Overlap_Metrics.py, Distance_w_Structural _Alignment.py,
               data_evaluator.py / metric_evaluator.py (alr reference sources)
  markdown/    converter.py
  studio/      base.py (shared _BaseStudio + tab classes),
               main.py (DataExtractionStudio), easa.py (EASAStudio)
run_studio.py, run_easa_studio.py     # entry points
lib/           bundled UI resources
archive/       experimental / superseded scripts (not imported)
```

All heavy third-party deps (docling, markitdown, PyMuPDF, xmltodict, openpyxl…)
are imported **lazily**, so every module imports even when they are absent —
they are only needed when an extraction actually runs.

## Requirements

- Python 3.10+ with Tkinter (ships with the standard python.org installer)
- Install everything from the pinned list:

  ```bash
  pip install -r requirements.txt
  ```

  | Tool | Needs |
  |---|---|
  | PDF Extraction & Review | `docling`, `docling-core`, `pymupdf`, `pandas`, `openpyxl`, `tqdm`, `colorama` |
  | Section Review | (stdlib only) |
  | EASA XML Extraction | `xmltodict`, `openpyxl` |
  | PDF → Markdown | `markitdown` |
  | AI Review evaluation | `nltk`, `rouge-score`, `Levenshtein`, `jiwer` (all optional — grounding, Jaccard and a difflib similarity ratio work without them) |

## What each tab does

### PDF Extraction & Review
`data_extraction/chunking/` — `logic.py`, `table_image_extractor.py`,
`chunk_review_ui.py`

- Converts a PDF with **Docling** and splits it into token-aware chunks
  (**HybridChunker**), each tagged with headings, page numbers, doc-item types.
- Same pass extracts **tables → CSV**, **images → PNG**, and layout headings.
- Caches the parse to JSON (re-runs hit the cache). Conversion runs on a
  background thread so the window stays responsive.
- Interactive review: step through each chunk, edit text/headings, **Log / Skip /
  Use-previous-heading**; logged chunks auto-merge under common headings, then
  chain into **Section Review**. **Resume** or **Reset** prior progress.
- The "Curation of the chunks extracted" panel is also the cache launcher (the
  former separate Cache Review tab was folded in here): pick an existing cache
  JSON + storage, get warned about existing processing footprints (with the
  option to switch to a fresh base folder), and file↔storage mappings are
  remembered in the workspace registry so browsing a known file auto-fills its
  storage destination.

### AI Review (chunks)
`data_extraction/chunking/ai_review_ui.py`

Loads a chunks cache (`*_docling_chunks_cache.json`) **or** a review output
(`Processed_chunks.json` with `merged_headings`) and lists the sections with
✓ checkboxes (Select all / Clear checks, heading toggle, search filter). The
right side is the same AI Review workbench as the EASA studio — shared
sections queue, free-form review and column analysis with all their features
(see the EASA *AI Review* bullet below) — implemented once in
`data_extraction/ai_utils/review_panel.py` and reused by both studios.

### Section Review
`data_extraction/chunking/section_review_ui.py`

Loads the merged sections from a chunk-review output JSON; select/edit/add/delete
sections and export back to the file. Also available as its own tab to reopen an
existing output later.

### EASA XML Extraction
`data_extraction/easa/parser.py` (+ `run_main.py`, `graph_builder.py`)

Point at an EASA e-Rules **XML ZIP** (single file or a folder) to produce raw
XML→JSON, a recursive **rules-hierarchy JSON**, extracted **images** and
**tables (→ Excel)**, and a **Master Structural Index** Excel. Optionally builds
the **Cosmograph** node/edge graph (CSV + Excel; unmatched links become
"External Reference" nodes). Runs on a background thread with a live log.

### EASA JSON Review
`data_extraction/easa/json_review_ui.py`

Interactive tree viewer for the structured JSON: navigate the hierarchy, and per
node see the text, EASA attributes, hyperlinks, and extracted assets:

- **Image preview** — inline, using Pillow when available (JPG/BMP/TIFF/PNG/GIF)
  and Tk's built-in PNG/GIF otherwise; anything else opens externally.
- **Table preview** — `.xlsx` tables render as a grid (via openpyxl).
- **Search/filter** the tree; **Summary** gives a document overview (totals and a
  per-element-type breakdown).
- **Checkbox selection** — every node has a ✓ checkbox (click it, or press
  Space on the highlighted row); **Select all** / **Clear checks** buttons and
  the ✓ column heading toggle all visible nodes at once. Checks survive
  filtering, and **＋ Add checked to batch** queues all checked sections for a
  looped AI run in document order (duplicates are skipped).
- **Export** — node index → CSV/Excel, full text → Markdown, or the selected
  subtree → JSON.
- **AI Review** (`data_extraction/ai_utils/`) — a full-width top-level page
  next to *Browse & Review*, fed by a shared **sections queue** (add the
  current node or all checked nodes; remove/clear). Two modes:
  - *Free-form review* — preset or custom instruction per section, run on the
    current node or loop over the queue. An **Answer as** picker (Plain text /
    Markdown / JSON / CSV table) is sent to the model so responses come back in
    that format — queued runs ask for it explicitly before starting. Every run
    first asks **where to store the results** (the file type — Markdown / CSV /
    Excel / JSON — is the export format) and writes them there automatically
    when the run finishes; cancel the dialog to abort the run.
  - *Column analysis* — define **columns** (name + what the LLM should
    extract); a live **prompt preview** rebuilds as you add/edit/remove them.
    Each queued section is analyzed into one table row per section with one
    cell per column (the model is asked for strict JSON; unparseable replies
    are surfaced as `[unparsed]`). Every analysis first asks for a **storage
    file**: choosing the *same* `.xlsx` again adds each new run as its own
    snapshot sheet (`Run N <timestamp>`) holding all sections of that batch,
    so one workbook accumulates the whole history (CSV/JSON hold the latest
    batch). Saving is **row by row**: the file (and the run's sheet) is
    created with the first LLM reply and re-saved after every further
    section, so nothing is lost if a long batch is interrupted. After each
    saved row a companion `Uniq Run N …` sheet is refreshed via
    `scripts/excel_file_utils.save_unique_elements_to_new_sheet`: every
    column definition has a ✓ checkbox (heading click toggles all; new
    columns start checked) and only the **checked** columns are collected,
    each unique element with its occurrence **count** in the adjacent cell. Double-click a row for full
    values; **Export table…** appends the same kind of snapshot.
  - *Evaluation* (`data_extraction/evaluation/column_evaluator.py`) — its own
    page next to *Column analysis*, fully independent of it: it evaluates
    **files that already hold analyzed data**. Pick any analysis workbook,
    one `Run N …` sheet or *All runs*, and the **sections JSON the analysis
    was made from** (always asked; chunks cache, `Processed_chunks.json` or
    an EASA structured JSON — the shape is auto-detected). Each run-sheet
    row is matched automatically: its `Section` value is looked up among the
    JSON's section titles and evaluated against that section's text. Every
    evaluation is an individual checkbox — **substring check** (each cell item grounded in its
    reference, adapted from the alr `data_evaluator`), **Jaccard**,
    **ROUGE-1/2/L**, **BLEU**, **Levenshtein distance**, **similarity ratio**
    and **word error rate** per cell vs. reference (adapted from the alr
    `metric_evaluator`; each guarded, so a missing library just leaves that
    metric empty). Results go into an `Eval Run N …` sheet (per-column
    True/False counts, metric columns, per-section **Coverage %**, an average
    row) — written **into the same workbook or into a new evaluation
    workbook** (your choice; the new file also gets a copy of the run sheet).
    Optionally the unique generated values are identified via
    `scripts/excel_file_utils.save_unique_elements_to_new_sheet` (the
    `Uniq Run N …` sheet) and every unique element is grounded against the
    combined reference text into a `Uniq Eval Run N …` sheet with per-column
    coverage summaries. An **Auto-evaluate after the analysis** checkbox sits
    next to *Analyze queued* on the Column analysis tab (on by default): when
    it is ticked, starting an analysis first pops up a picker with all the
    possible evaluations so you choose what gets evaluated for that batch —
    *Skip evaluation* opts out for that run; the picker remembers its own
    choices, independent of the Evaluation tab. The chosen evaluations then
    run **section by
    section**: right after each section's result row is saved it is evaluated
    against its reference text and the `Eval Run N …` (and `Uniq Eval Run N
    …`) sheets are refreshed before the next section is analyzed, so an
    interrupted batch already has every finished section evaluated.
    Standalone:
    `python -m data_extraction.evaluation.column_evaluator results.xlsx
    sections.json [--metrics …] [--out eval.xlsx]` (sections JSON = chunks
    cache or `Processed_chunks.json`).

  A **progress bar with a console** sits at the bottom of the AI Review page
  and logs every step of all three modes: which section is being analyzed,
  the prompt sent to the LLM, its response, and — during evaluations — the
  reference text, the evaluated cell texts and the computed results
  (previews are truncated; the console keeps the last ~2000 lines).

  LLM calls run on a background thread. Manage keys with
  the **API keys…** button on the tab (add / edit / clear, persisted to
  `API_keys_config.json` under `ALR_MAIN_FOLDER`); the
  `Ollama_DLR_API_Key` / `BlaBla_API_Key` env vars still take precedence.

### PDF → Markdown
`data_extraction/markdown/converter.py`

Batch-converts PDFs to Markdown via **MarkItDown** (file or folder), one `.md`
per PDF. Threaded with a live log.

## How the launchers host the tools

`studio/base.py` provides `_BaseStudio`: an `_EmbeddedRoot` shim lets each
window-owning UI class run inside a notebook tab, and `launch_review_app` is
patched so chunk review opens as a modal `Toplevel` (then chains Section Review)
instead of spawning a second root window. `main.py` and `easa.py` just subclass
`_BaseStudio` and declare their `TAB_SPECS`.

## Standalone use

Individual tools also run on their own via `-m` from the repo root:

```bash
python -m data_extraction.easa.extraction_ui         # EASA extraction window
python -m data_extraction.easa.json_review_ui [file.json]   # EASA JSON review
python -m data_extraction.chunking.ai_review_ui [chunks.json]  # chunk AI review
python -m data_extraction.evaluation.column_evaluator results.xlsx sections.json
                                                      # evaluate a stored analysis
python -m data_extraction.studio.main                # same as run_studio.py
```

## Configuration

- **Workspace registry path** — defined once in
  `data_extraction/chunking/workspace_config.py`; OS-appropriate default,
  overridable with the `DOCLING_WORKSPACE_REGISTRY` environment variable.
- **Image de-duplication** — the PDF extractor de-duplicates images perceptually
  only when the optional `imagehash` package is installed; otherwise every image
  is kept.
- **Crash logs** — every launcher installs `data_extraction/crash_logging.py` at
  startup, so any uncaught exception (main thread, background thread, or Tk
  callback) is written with its full traceback to
  `~/.data_extraction/logs/crashes.log` (override the folder with the
  `DATA_EXTRACTION_LOG_DIR` environment variable). Tk-callback crashes also show
  an error dialog pointing at the log file instead of failing silently.

## Known limitations

- Import- and headless-Tk verified; the full GUI with real extractions needs the
  heavy deps installed — validate on the target (Windows) environment.
- Some backend modules keep hardcoded example paths in their `__main__` blocks;
  adjust when running them directly.
