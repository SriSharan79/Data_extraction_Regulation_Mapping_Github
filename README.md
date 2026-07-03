# Data Extraction & Regulation Mapping

Tools for turning regulatory source documents (PDFs and EASA e-Rules XML) into
structured, human-curated data — chunked text, tables, images, section trees,
metadata indexes, and cross-reference graphs.

Every tool is reachable from a single launcher, **`Data_Extraction_Studio.py`**,
which hosts each one as a tab. Each tool also still runs standalone.

## Quick start

```bash
# from the repository root
python3 Data_Extraction_Studio.py
```

The launcher opens a tabbed window. Tabs load lazily and are fault-isolated: if a
tool's dependencies aren't installed, only that tab shows an error — the others
keep working.

## Requirements

- Python 3.10+ with Tkinter (ships with the standard python.org installer)
- Python packages (install what you need for the tabs you use):

  ```bash
  pip install docling docling-core markitdown pymupdf openpyxl xmltodict colorama pandas tqdm
  ```

  | Tool / tab | Needs |
  |---|---|
  | PDF Extraction & Review, Cache Review Launcher | `docling`, `docling-core`, `pymupdf`, `pandas`, `openpyxl`, `tqdm`, `colorama` |
  | Section Re-Writer | (stdlib only) |
  | EASA XML Extraction | `xmltodict`, `openpyxl` |
  | PDF → Markdown | `markitdown` |

## What each tab does

### PDF Extraction & Review  /  Cache Review Launcher
`Manual_Chunking_With_UI/Chunk_review_logic.py`, `Table_image_extractor.py`,
`chunk_review_ui.py`, `cache_launcher.py`

- Converts a PDF with **Docling** and splits it into token-aware chunks
  (**HybridChunker**), each tagged with headings, page numbers and doc-item types.
- Same pass extracts **tables → CSV**, **images → PNG**, and **layout headings**.
- Caches the parse to JSON (re-runs hit the cache and skip reprocessing).
- Interactive review window: step through each chunk, edit text/headings, then
  **Log / Skip / Use-previous-heading**. Logged chunks are auto-merged under
  common headings into a structured output JSON.
- **Resume** prior progress or **Reset** it. A per-document / per-date workspace
  is derived automatically, and a JSON registry remembers which storage folder
  pairs with each source file.
- Run modes: *generate cache only*, *cache + full review*, *review from an
  existing cache*.

### Section Re-Writer
`Manual_Chunking_With_UI/Raw_sec_rewriter.py`

Load a raw-chunks JSON (browsable tree + text inspector on the left) and
build/edit target **sections** on the right (add, delete, rename, edit body).
"Auto-reconstruction" bootstraps sections by grouping chunks under their first
heading. Exports a clean sections JSON.

### EASA XML Extraction
`EASA_Data_Extractors/EASA_Parser.py`

Point at an EASA e-Rules **XML ZIP** (single file or a folder of them) to produce:

- raw XML → JSON,
- a recursive **rules-hierarchy JSON** (topics, attributes, text, hyperlinks, children),
- extracted **images** and **tables (→ Excel)** tagged to their sections,
- a **Master Structural Index** Excel (22 EASA metadata attributes + per-node metrics).

Optionally builds the cross-reference graph below in the same run. Runs on a
background thread with a live log pane.

### (EASA cross-reference graph)
`EASA_Data_Extractors/EASA_Graph_builder.py`

Turns an EASA extraction JSON into **Cosmograph node/edge tables** (CSV *and*
Excel) by matching hyperlinks between topics. Unmatched links become
"External Reference" nodes so cross-document references stay visible.
Enabled via the checkbox on the EASA tab.

### PDF → Markdown
`Markdown_extraction/Process_files_to _Markdown.py`

Batch-converts PDFs to Markdown via **MarkItDown** (single file or a folder),
one `.md` per PDF. Threaded with a live log pane.

## How the launcher hosts the tools (non-invasive)

`Data_Extraction_Studio.py` does not modify any tool file:

- An `_EmbeddedRoot` shim lets each window-owning UI class run inside a notebook
  tab unchanged (it no-ops window-only calls like `title`/`geometry`/`destroy`).
- `launch_review_app` is monkeypatched **on the imported modules only** so the
  chunk-review window opens as a modal `Toplevel` of the launcher instead of a
  second root window.

## Standalone use

Each module still runs on its own, e.g.:

```bash
python3 Manual_Chunking_With_UI/Chunk_review_logic.py   # extraction + review launcher
python3 Manual_Chunking_With_UI/Raw_sec_rewriter.py     # section re-writer
python3 EASA_Data_Extractors/EASA_Parser.py             # EASA extraction (edit paths in __main__)
```

## Configuration

- **Workspace registry path** — where the file-to-storage registry lives is
  defined once in `Manual_Chunking_With_UI/workspace_config.py`. It defaults to
  an OS-appropriate path and can be overridden with the
  `DOCLING_WORKSPACE_REGISTRY` environment variable.
- **Image de-duplication** — the PDF extractor de-duplicates images
  perceptually only when the optional `imagehash` package is installed;
  otherwise every image is kept.

## Known limitations / TODO

- The launcher is byte-compile / import verified; the **live Tkinter GUI has not
  been run** in CI — validate on the target (Windows) environment.
- Some modules still contain **hardcoded example paths** in their `__main__`
  blocks (`C:\Users\...`, `U:\...`); adjust these when running a module
  standalone.
