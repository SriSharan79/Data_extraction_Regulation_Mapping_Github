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
    from EASA_Parser import extract_easa_from_zip_v3, resolve_paths
    from EASA_Graph_builder import export_to_cosmograph_csv

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

        if build_cosmograph:
            try:
                export_to_cosmograph_csv(output_json)
            except Exception as exc:  # noqa: BLE001 - keep processing the rest
                print(f"[graph] Skipped graph build for {zip_path}: {exc}")

    print(f"\nDone. Produced {len(produced)} structured JSON file(s).")
    return produced


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Extract EASA XML ZIP archives.")
    parser.add_argument("source", help="A .zip file or a folder of .zip files")
    parser.add_argument("storage", help="Workspace / storage directory")
    parser.add_argument("--graph", action="store_true", help="Also build Cosmograph CSV/Excel")
    args = parser.parse_args()

    main(args.source, args.storage, build_cosmograph=args.graph)
