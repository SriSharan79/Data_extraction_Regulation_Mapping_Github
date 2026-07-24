"""
Headless orchestration entry for EASA XML extraction.

Wraps EASA_Parser.extract_easa_from_zip_v3 (and, optionally, the Cosmograph
graph builder) behind a single ``main()`` that accepts either one .zip file or
a folder of them. Used both by the standalone EASA extraction UI
(EASA_Extraction_UI.py) and by the "EASA XML Extraction" tab in
Data_Extraction_Studio.py.

It prints progress to stdout so a GUI can capture and display it.
"""

import os
from pathlib import Path


def collect_zip_targets(src_path: str) -> list:
    """Return the list of .zip files to process from a file or a folder."""
    if os.path.isdir(src_path):
        return sorted(str(p) for p in Path(src_path).rglob("*.zip"))
    if src_path.lower().endswith(".zip"):
        return [src_path]
    return []


def main(src_path: str, storage_base: str, build_cosmograph: bool = False) -> list:
    """
    Extract one or more EASA XML ZIP archives into ``storage_base``.

    Args:
        src_path: a single ``.zip`` file, or a folder containing ``.zip`` files.
        storage_base: workspace/storage directory for the outputs.
        build_cosmograph: also build Cosmograph node/edge CSV+Excel from each
            resulting structured JSON.

    Returns:
        List of structured-JSON output paths that were produced (one per archive).
    """
    # Imported lazily so simply importing this module (e.g. from the studio) does
    # not require xmltodict/openpyxl until an extraction is actually run.
    from .parser import extract_easa_from_zip_v3, resolve_paths
    from .graph_builder import export_to_cosmograph_csv

    targets = collect_zip_targets(src_path)
    if not targets:
        print(f"No .zip files found at the source path: {src_path}")
        return []

    print(f"Found {len(targets)} archive(s) to process.")
    produced = []
    for zip_path in targets:
        print(f"\n=== Processing: {zip_path} ===")
        extract_easa_from_zip_v3(zip_path, storage_base)

        paths = resolve_paths(storage_base, zip_path)
        output_json = paths["output_json"]
        produced.append(output_json)

        # Best-effort: record the document + its nodes in the workspace's
        # SQLite store (the storage base is the workspace). Never fatal —
        # a DB problem must not abort the extraction of the remaining files.
        _persist_easa_to_db(zip_path, storage_base, output_json)

        if build_cosmograph:
            try:
                export_to_cosmograph_csv(output_json)
            except Exception as exc:  # noqa: BLE001 - keep processing the rest
                print(f"[graph] Skipped graph build for {zip_path}: {exc}")

    print(f"\nDone. Produced {len(produced)} structured JSON file(s).")
    return produced


def _persist_easa_to_db(zip_path: str, storage_base: str, output_json: str) -> None:
    """Record one extracted EASA document + its hierarchy nodes in the
    workspace SQLite store. Best-effort: any failure is printed and swallowed
    so it never interrupts the extraction run."""
    import json as _json
    import os as _os

    try:
        from data_extraction.db import facade
    except Exception as exc:  # noqa: BLE001 - DB layer optional
        print(f"[db] SQL persistence unavailable: {exc}")
        return
    try:
        if not output_json or not _os.path.exists(output_json):
            return
        with open(output_json, "r", encoding="utf-8") as f:
            structured = _json.load(f)
        ws = facade.open_workspace(storage_base)
        doc_uuid = facade.persist_easa_document(
            ws, source_path=zip_path, storage_root=_os.path.dirname(output_json),
            structured_data=structured, review_path=output_json)
        if doc_uuid:
            print(f"[db] Recorded EASA document in workspace store (doc {doc_uuid}).")
    except Exception as exc:  # noqa: BLE001 - persistence is best-effort
        print(f"[db] Skipped DB persistence for {zip_path}: {exc}")


if __name__ == "__main__":
    import argparse

    from ..crash_logging import install
    install()

    parser = argparse.ArgumentParser(description="Extract EASA XML ZIP archives.")
    parser.add_argument("source", help="A .zip file or a folder of .zip files")
    parser.add_argument("storage", help="Workspace / storage directory")
    parser.add_argument("--graph", action="store_true", help="Also build Cosmograph CSV/Excel")
    args = parser.parse_args()

    main(args.source, args.storage, build_cosmograph=args.graph)
