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

### Extract & Review (unified)
`data_extraction/studio/extract_review_tab.py`

One workspace-driven surface that folds extraction and review into a single
flow for both document types — the first tab in both studios. A **source
selector** (PDF · Docling chunks / EASA · XML), a shared **workspace bar** (the
storage folder backed by the SQLite store), and a **document rail** listing what
the workspace already holds (from the database). Press **＋ Extract new…** to run
the source's extractor with the workspace pre-filled; pick a document in the
rail to review it, with mode buttons switching between the review surfaces the
source offers (PDF: *Section review* / *AI review*; EASA: *Browse + AI review*).

It is an orchestrator, not a rewrite: it *hosts* the existing tools (below)
unchanged, so every feature they have is still here. The original per-tool tabs
remain available alongside it.

### PDF Extraction & Review
`data_extraction/chunking/` — `logic.py`, `table_image_extractor.py`,
`chunk_review_ui.py`

- Converts a PDF with **Docling** and splits it into token-aware chunks
  (**HybridChunker**), each tagged with headings, page numbers, doc-item types.
- Same pass extracts **tables → CSV**, **images → PNG**, and layout headings.
- Caches the parse to JSON (re-runs hit the cache). Conversion runs on a
  background thread so the window stays responsive.
- **Automatic triage before review** (`chunk_triage.py`, modelled on the alr
  pipeline's section processing): the extracted chunks are sorted on their own
  before anyone looks at them. The **Table of Contents is read from the PDF
  itself** first — the embedded outline (bookmarks) when present, else the
  printed TOC parsed from the first pages' text — and the Docling chunk
  headings are **verified against it**; a document-internal (chunk-level) TOC
  is the next fallback. Only when the PDF offers no usable TOC are the
  collected headings sent to an **LLM that keeps only the valid ones**
  (numbered/designated headings such as `Part 21`, `AMC1 ORO.GEN.200`,
  `1.2 Scope` are always kept), and failing that deterministic rules apply.
  Every kept chunk records whether its heading **matched the TOC** — the
  review table shows it as a **TOC ✓/✗ column** (✗ = kept, e.g. a numbered
  heading, but absent from the TOC) with a `TOC-verified: n ✓ / m ✗` count in
  the header; reviewer overrides clear the flag. An **Open PDF** button next
  to the file picker opens the selected PDF with the system viewer, and a
  **View TOC** button shows exactly the TOC the triage reads from the PDF via
  PyMuPDF (entries, pages, source: embedded outline vs. printed text) — or
  explains the LLM fallback when the PDF has none. Each chunk then gets a proposed decision with a reason: page
  headers/footers, repeated running text, the TOC page itself and empty
  fragments are proposed as **Skip**; content whose heading matches a real
  section is **Kept** under it; content under a noise heading — or with no
  heading — is folded into the preceding valid section, which is what the
  manual *Use previous heading* click used to do one chunk at a time.
- **Bulk review, embedded in the tab** (`chunk_triage_ui.py`): the tab is one
  page — pick the **PDF** (or an existing **cache JSON** to skip
  re-extraction) and the **storage destination**, press **▶ Run Extraction &
  Review**, and the triage appears right below: all chunks in one sortable
  table (decision, heading, page, type, why, preview) filtered to **Needs
  review** by default, with bulk *Keep* / *Skip* / *Set heading* / *Use
  heading above*, per-chunk editing, and a merged-section preview. The
  triage's proposal is shown as the **LLM decision**; a separate **User
  decision** column offers a per-row dropdown (click the cell) with **Keep /
  Merge with above / Skip / Need review** — *Merge with above* folds the
  chunk into the nearest preceding kept section. The preview shows both
  variants side by side (**LLM suggested** and **user decided** tabs), and
  when User decisions exist *Accept & save* first asks **which variant to
  write** (user-decided rows are marked `logged (user)` in the output).
  **Nothing
  is written until you press *Accept & save*** — then the same
  `merged_headings` / `raw_session_history` payload as before is written
  **once** (the old tool rewrote the whole file on every click), and **Section
  Review** opens on the saved file. Sections are grouped by *normalised*
  heading, so `Part 21` and `PART 21 ` no longer become two sections.
- **Automatic resume + versioned saves**: when the chosen document + storage
  match a previous run, the most recent reviewed `Processed_chunks*.json`
  under that document's storage root (any dated folder) is found and its
  decisions are **pre-applied automatically** — accepts, skips, reassigned
  headings and edited text return marked *restored from the previous session*.
  On *Accept & save*: **no changes → the existing review file is kept as it
  is**; **any change → the review is written as a NEW file** (a timestamped
  name when today's file already exists), so an earlier review is never
  overwritten. The original one-chunk-at-a-time tool remains available via
  `launch_review_app(..., bulk=False)`.
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
    extract); every column has a **Run ✓** checkbox (only checked columns
    are analyzed — heading click toggles all) next to its **Uniq ✓**
    checkbox, and the definitions with both check states are
    **remembered across sessions** (saved to
    `~/.data_extraction/column_analysis_columns.json` on every change), so
    closing and reopening the studio brings them back. The rest of the
    page's session state persists too
    (`~/.data_extraction/review_ui_state.json`): the last storage file,
    the LLM service and selected models, the auto-evaluate box and its
    picker choices, the Evaluation tab's workbook/metric/output choices,
    and the window geometry with the console-sash position. Every AI
    Review tab (Free-form review, Column analysis, Evaluation, Unique
    elements) sits in a **vertically scrollable pane** — on a small
    window the tab scrolls (scrollbar or mouse wheel; text boxes and
    tables keep their own wheel), on a large one the content stretches
    to fill it, so no widget can end up unreachable. The **progress
    console** colours its lines by severity (errors red, warnings amber,
    saves green, run headers blue), has a live **filter** box (hides
    non-matching lines), an **auto-scroll** toggle, a **Save log…**
    button, and an **Open results** button that opens the latest result
    file (analysis workbook, free-form autosave, evaluation output,
    uniques/entities sheet) with the system's default app. The LLM row
    hosts an **inline model picker** plus an *Active: service · model* chip
    showing exactly what the next run uses. **Blablador is the first and default
    service, then Chat AI (chat-ai.academiccloud.de), then DLR Ollama**
    — the same preference order drives the automatic fallback when a
    chat call fails — and a background **availability probe** at startup only
    offers a service when its **API key is stored and a model list is
    known** (with nothing usable, the chip warns and points to *API keys…*).
    Model lists are **cached**: they are fetched once and stored per service
    (`model_lists_cache.json` under `ALR_MAIN_FOLDER`), then reused on every
    startup and probe without another network call — the **↻ button** is the
    explicit *refresh* that re-fetches a service's list live and updates the
    cache (a failed refresh keeps the stored list). The same
    gated **service + model picker sits on the PDF Extraction tab** for
    the triage's LLM heading check (used only when the document has no
    usable Table of Contents). Every Evaluation-tab metric checkbox has a
    **hover tooltip** explaining the metric and its direction
    (higher/lower is better). The page uses the *clam* ttk theme and
    keyboard shortcuts: **Ctrl/Cmd+Return** = Analyze queued,
    **Ctrl/Cmd+P** = Pause/Resume, **Ctrl/Cmd+.** = Stop. A **＋ Specific
    entities** button adds the predefined aviation entity-chain column
    (`data_extraction/ai_utils/entity_chains.py`): the LLM extracts
    `Reference|System Info|Process|Personal|QuantityValue` chains
    (components **pipe-separated**, `;` between chains, `#` = missing
    optional component, Reference mandatory). The connector is a pipe (`|`)
    rather than a hyphen precisely because References contain hyphens
    themselves (`FAA AC 120-76D`, `RTCA DO-178C`), which used to break the
    split; the pipe never appears inside a value, so parsing is an
    unambiguous positional split (legacy hyphen-format values still parse
    via a fallback). When this column is part of a run its chains are
    **auto-parsed into an `Entities <run sheet>` sheet** after every saved
    row — one row per chain, one column per component plus the raw chain.
    Because the chain format is itself a valid answer, a reply that is
    **bare chain text rather than JSON is accepted and recorded as-is**
    instead of triggering the JSON retries (a fenced ```` ``` ```` wrapper is
    stripped; a JSON-wrapped chain is unwrapped) — only genuinely unusable
    replies still cost retries and end up `[unparsed]`. In
    per-column call mode this column's prompt is sent as-is and the raw
    reply kept (no JSON wrapper). The *Unique elements* tab has the same
    parser **standalone**: pick any workbook/sheet/column holding chain
    values and *Parse chains → sheet* writes `Entities <sheet>` into the
    same file. A live **prompt
    preview** rebuilds as you add/edit/remove them
    and is **editable**: whatever you type above the `---` marker is sent as
    the prompt of the next run (changing the column definitions rebuilds it).
    Each queued section is analyzed into one table row per section with one
    cell per column (the model is asked for strict JSON; an unparseable
    reply is retried — the same prompt is re-sent up to 3 more times — and
    only then surfaced as `[unparsed]`). Starting an analysis first asks **how to
    call the LLM**: *one call per section* (all columns in a single JSON
    answer) or *one call per column value* (each column of a section is its
    own focused LLM call; non-JSON replies are kept as the raw value). While
    a run is going, **Pause / Resume / Stop** buttons next to *Analyze
    queued* control it: pausing waits after the current LLM call (a
    **⏸ PAUSED** indicator appears beside the progress bar), stopping
    ends the run after it — every row already saved (and its evaluation)
    is kept. A **Re-run failed** button lights up after a run that left
    `[ERROR]`/`[unparsed]` rows behind: it re-analyzes **only those
    sections** with the run's own settings (prompt, call mode, service,
    model) and **updates their existing rows in place** — in the table,
    the run sheet, the Uniq/Entities sheets and the evaluation — instead
    of appending. The **sections queue shows a live status per section**
    (⏳ queued / ▶ running / ✔ done / ✖ error — error covers `[ERROR]` and
    `[unparsed]` results) for both run modes, and the progress bar carries
    a live **`done/total · ETA · LLM calls` counter** (retries and
    per-column calls included). Every analysis also asks for a **storage
    file**: choosing the *same* `.xlsx` again adds each new run as its own
    snapshot sheet (`Run N <timestamp>`) holding all sections of that batch,
    so one workbook accumulates the whole history (CSV/JSON hold the latest
    batch). Saving is **row by row**: the file (and the run's sheet) is
    created with the first LLM reply and re-saved after every further
    section, so nothing is lost if a long batch is interrupted. Every saved
    row also carries three trailing **provenance columns** — *Service Used*,
    *Model Used*, *Fallback?* — filled from `get_last_call_info()` so the
    file records which service/model actually answered each section and
    flags any cross-service fallback (the free-form export gets the same
    `service_used` / `model_used` / `fallback` columns). They are written to
    the file only, never mixed into the analysis columns the entity and
    evaluation passes read. After each
    saved row a companion `Uniq Run N …` sheet is refreshed via
    `scripts/excel_file_utils.save_unique_elements_to_new_sheet`: every
    column definition has a ✓ checkbox (heading click toggles all; new
    columns start checked) and only the **checked** columns are collected,
    each unique element with its occurrence **count** in the adjacent cell. Double-click a row for full
    values; **Export table…** appends the same kind of snapshot.
  - *Evaluation* (`data_extraction/evaluation/column_evaluator.py`) — its own
    page next to *Column analysis*, fully independent of it: it evaluates
    **files that already hold analyzed data**. Pick any workbook and **any
    sheet** — the dropdown lists every sheet of the file, exactly like the
    Unique elements tab: `Run N …` snapshots, hand-made sheets and even the
    pipeline's own result sheets (only the *All sheets* bulk option skips
    the result sheets — `Eval`/`Uniq`/`Entities`/per-metric — so an
    evaluation never evaluates its own outputs). Sheets with a `Section`
    column work as before; for a
    sheet **without** one, a dialog asks **which column substitutes it** —
    that column's values become the per-row key matched against the
    reference section titles and every other column is evaluated
    (`section_column=` / `section_columns=` in the evaluator API) — plus
    the **sections JSON the analysis
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
    Alongside the per-metric item sheets, the **complete sentence-level
    record** is stored as JSON in a `Metric_Sentence_Details/` folder next to
    the evaluation workbook (one file per run, mirroring alr): for every
    extracted item, every reference sentence with every selected metric's
    value against it — the full item×sentence matrix the per-item sheets keep
    only the best value of — plus the kept best value/sentence per metric.
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

  A **progress bar with a console** sits in its own pane at the bottom of the
  AI Review page (~20 % of the window; drag the divider to resize) and logs
  every step of all three modes: which section is being analyzed,
  the prompt sent to the LLM, its response, and — during evaluations — the
  reference text, the evaluated cell texts and the computed results
  (previews are truncated; the console keeps the last ~2000 lines).

  LLM calls run on a background thread. Manage keys with
  the **API keys…** button on the tab (add / edit / clear, persisted to
  `API_keys_config.json` under `ALR_MAIN_FOLDER`); the
  `Ollama_DLR_API_Key` / `BlaBla_API_Key` / `ChatAI_API_Key` env vars still
  take precedence. Chat AI (chat-ai.academiccloud.de) is available for both
  chat completions and embeddings (`e5-mistral-7b-instruct` by default).
  Every remote request retries transient failures (429/5xx, timeouts) with
  backoff on the same service — honouring `Retry-After` — before the
  cross-service fallback kicks in, and `llm_call(allow_fallback=False)`
  pins a batch to one service so results never mix models mid-run.
  `get_last_call_info()` reports which service/model actually answered the
  most recent call (and whether it was a fallback) for provenance records.

### Data & Analysis
`data_extraction/studio/data_analysis_tab.py`

A read-across view of everything the extraction and AI-review flows persisted
to SQLite (see *SQLite persistence* below), present in both studios. Open one
workspace's database (or **Scan all workspaces** to list documents across every
project from the global registry), pick a document, and browse its stored
**Sections**, **AI Reviews** (with Service / Model / Fallback provenance) and
**Entities**. The **Cross-document AI analysis** panel builds a corpus from the
workspace's documents (all, or PDF-only / EASA-only), runs it through a gated
LLM picker (BlaBla first), shows the answer, and stores it — *Past analyses…*
lists everything recorded. Additive: it only reads what the other tabs wrote.

### PDF → Markdown
`data_extraction/markdown/converter.py`

Batch-converts PDFs to Markdown via **MarkItDown** (file or folder), one `.md`
per PDF. Threaded with a live log.

## SQLite persistence (`data_extraction/db/`)

Alongside the JSON/Excel outputs, every extraction and AI review is also
recorded in SQLite so the data can be reviewed in the tool and reused for
downstream AI analysis. Persistence is **best-effort** — a database problem is
logged and swallowed, never blocking a review or an extraction.

Two tiers:

- **Per-workspace data DB** — `<workspace>/extraction_index.db` holds the heavy
  data: `documents`, `sections` (PDF merged headings *and* EASA hierarchy nodes,
  flattened with parent/level), `ai_runs`, `ai_reviews` (with the
  Service/Model/Fallback provenance kept in their own columns), `entities`
  (parsed Specific-entities chains), and `analysis_results` (cross-document AI
  analysis output). The *workspace* is the storage destination you already pick
  for a run.
- **Global registry DB** — `~/.data_extraction/registry.db` (override with
  `DATA_EXTRACTION_HOME`) indexes which workspaces exist and mirrors each
  document, so the tool can list everything across projects.

The facade (`db/facade.py`) is the one entry point: `open_workspace(dir)` plus
mappers `persist_pdf_review`, `persist_easa_document`, `persist_ai_reviews_for`.
Migrations only ever ADD columns, and `upsert_document` preserves review/analysis
columns with COALESCE, so re-extracting a document never wipes its review data.
AI-review batches attach to the already-extracted document by name.

**Cross-document analysis** (`db/analysis.py`): `run_cross_document_analysis`
builds one corpus from many stored documents (size-guarded), asks an injected
LLM a question, records the answer + provenance + scope in `analysis_results`,
and returns it. The LLM is injected as a `llm(prompt, system_prompt)` callable
(`default_llm('b'|'c'|'o', model)` builds one from `llm_call`), so it tests
headless.

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
