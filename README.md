# Data Extraction & Regulation Mapping

Tools for turning regulatory source documents (PDFs and EASA e-Rules XML) into
structured, human-curated data — chunked text, tables, images, section trees,
metadata indexes, and cross-reference graphs.

Everything lives in the `data_extraction` Python package and is reached through
two notebook launchers.

## Quick start

```bash
# from the repository root
python run_studio.py        # PDF/chunk tools: extraction & review, cache review,
                            # section review, PDF -> Markdown
python run_easa_studio.py   # EASA tools: XML extraction, structured-JSON review
```

Each launcher opens a tabbed window. Tabs load lazily and are fault-isolated: if
a tool's heavy dependency isn't installed, only that tab shows an error (or the
failure surfaces when you actually run it) — the other tabs keep working.

## Layout

```
data_extraction/
  chunking/    logic.py, chunk_review_ui.py, section_review_ui.py,
               cache_launcher.py, table_image_extractor.py, workspace_config.py
  easa/        parser.py, graph_builder.py, run_main.py,
               extraction_ui.py, json_review_ui.py
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
  | PDF Extraction & Review, Cache Review | `docling`, `docling-core`, `pymupdf`, `pandas`, `openpyxl`, `tqdm`, `colorama` |
  | Section Review | (stdlib only) |
  | EASA XML Extraction | `xmltodict`, `openpyxl` |
  | PDF → Markdown | `markitdown` |

## What each tab does

### PDF Extraction & Review / Cache Review Launcher
`data_extraction/chunking/` — `logic.py`, `table_image_extractor.py`,
`chunk_review_ui.py`, `cache_launcher.py`

- Converts a PDF with **Docling** and splits it into token-aware chunks
  (**HybridChunker**), each tagged with headings, page numbers, doc-item types.
- Same pass extracts **tables → CSV**, **images → PNG**, and layout headings.
- Caches the parse to JSON (re-runs hit the cache). Conversion runs on a
  background thread so the window stays responsive.
- Interactive review: step through each chunk, edit text/headings, **Log / Skip /
  Use-previous-heading**; logged chunks auto-merge under common headings, then
  chain into **Section Review**. **Resume** or **Reset** prior progress.

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
    that format — queued runs ask for it explicitly before starting. Export to
    Markdown / CSV / Excel / JSON.
  - *Column analysis* — define **columns** (name + what the LLM should
    extract); a live **prompt preview** rebuilds as you add/edit/remove them.
    Each queued section is analyzed into one table row per section with one
    cell per column (the model is asked for strict JSON; unparseable replies
    are surfaced as `[unparsed]`). Double-click a row for full values; export
    the table to Excel / CSV / JSON.

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
