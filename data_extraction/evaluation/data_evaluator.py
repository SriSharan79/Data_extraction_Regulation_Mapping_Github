import os
import sys
from pathlib import Path
import pandas as pd

from alr.common.file_manager import DataAnalyzeManager, Vec_DB_Manager
from alr.common.sections import build_sections_eval_map
from alr.rag_builders.master_excel_db_builder import _fetch_metadata, _load_abstract_json, _load_recorded_abstracts
from alr.common.json_utils import get_key_from_file, get_value_by_pair
from alr.rag_builders.vector_db_updater import add_new_strings_to_index, create_faiss_index_cosine, load_index_file, save_index_file, search_similar, vectorize_strings
from colorama import Fore, Style


# =====================================================================
# MODULAR EXCEL HELPERS
# =====================================================================

def _is_duplicate_in_sheet(file_path, sheet_name, target_uuid):
    """Checks if a UUID already exists in a given Excel sheet."""
    if file_path.exists() and file_path.stat().st_size > 0:
        try:
            with pd.ExcelFile(file_path, engine='openpyxl') as xls:
                if sheet_name in xls.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
                    if not df.empty and "UUID" in df.columns:
                        if target_uuid in df["UUID"].astype(str).values:
                            return True
        except Exception as e:
            print(Fore.YELLOW + f"⚠️ Sheet '{sheet_name}' read error: {e}")
    return False


def _write_section_sheet_flat(file_path, sheet_name, data_entry):
    """Appends or updates a data entry for section files ensuring single row flat format per UUID."""
    try:
        df_new = pd.DataFrame([data_entry])
        all_sheets = {}
        sheet_order = []
        target_uuid = str(data_entry.get("UUID"))

        if file_path.exists() and file_path.stat().st_size > 0:
            with pd.ExcelFile(file_path, engine='openpyxl') as xls:
                sheet_order = list(xls.sheet_names)
                for s in sheet_order:
                    all_sheets[s] = pd.read_excel(file_path, sheet_name=s, engine='openpyxl')
            
            if sheet_name in all_sheets:
                df_sheet = all_sheets[sheet_name]
                df_sheet["UUID"] = df_sheet["UUID"].astype(str)
                match_mask = df_sheet["UUID"] == target_uuid
                
                if match_mask.any():
                    # Update row to dynamically balance any column structure changes
                    for col in df_new.columns:
                        df_sheet.loc[match_mask, col] = df_new[col].values[0]
                    all_sheets[sheet_name] = df_sheet
                else:
                    all_sheets[sheet_name] = pd.concat([df_sheet, df_new], ignore_index=True)
            else:
                all_sheets[sheet_name] = df_new
                sheet_order.append(sheet_name)
        else:
            all_sheets[sheet_name] = df_new
            sheet_order.append(sheet_name)

        with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
            for s in sheet_order:
                all_sheets[s].to_excel(writer, sheet_name=s, index=False)
    except Exception as e:
        print(Fore.RED + f"❌ Failed to write to section sheet '{sheet_name}': {e}")


def _update_master_overview(storage_dir, sheet_name, uuid, title, filename, text_content, true_count, false_count, full_section_entry, overview_path=None):
    """
    Updates the centralized Master Overview tracking true/false count per section
    and bulleted texts. ``overview_path`` selects the workbook (defaults to the
    abstract evaluation overview; the intro evaluation passes its own).
    """
    VDB = Vec_DB_Manager(storage_dir)
    try:
        master_overview_path = Path(overview_path) if overview_path else VDB.Abstract_Eval_Overview
        overview_sheet = "Overview"
        all_sheets = {}

        if master_overview_path.exists() and master_overview_path.stat().st_size > 0:
            with pd.ExcelFile(master_overview_path, engine='openpyxl') as xls:
                for s in xls.sheet_names:
                    all_sheets[s] = pd.read_excel(master_overview_path, sheet_name=s, engine='openpyxl')
        
        df_overview = all_sheets.get(overview_sheet, pd.DataFrame(columns=["UUID", "Title", "Filename"]))
        df_overview["UUID"] = df_overview["UUID"].astype(str)

        match_mask = df_overview["UUID"] == str(uuid)

        # Dynamic target tracking column assignments
        true_col = f"{sheet_name}_True_Count"
        false_col = f"{sheet_name}_False_Count"

        if match_mask.any():
            df_overview.loc[match_mask, sheet_name] = text_content
            df_overview.loc[match_mask, true_col] = true_count
            df_overview.loc[match_mask, false_col] = false_count
            if title: df_overview.loc[match_mask, "Title"] = title
            if filename: df_overview.loc[match_mask, "Filename"] = filename
        else:
            new_row = {
                "UUID": str(uuid),
                "Title": title,
                "Filename": filename,
                sheet_name: text_content,
                true_col: true_count,
                false_col: false_count
            }
            df_overview = pd.concat([df_overview, pd.DataFrame([new_row])], ignore_index=True)

        all_sheets[overview_sheet] = df_overview

        # First write overview updates back to dictionary structures
        with pd.ExcelWriter(master_overview_path, engine='openpyxl', mode='w') as writer:
            # Force 'Overview' tab to stay strictly at index position 0
            all_sheets[overview_sheet].to_excel(writer, sheet_name=overview_sheet, index=False)
            for s, df in all_sheets.items():
                if s != overview_sheet:
                    df.to_excel(writer, sheet_name=s, index=False)
                    
        # Reuse helper function to cleanly inject or update individual section tabs inside Master Overview
        _write_section_sheet_flat(master_overview_path, sheet_name, full_section_entry)

    except Exception as e:
        print(Fore.RED + f"❌ Failed to update single Master Overview file: {e}")


# =====================================================================
# CORE PROCESSING LAYER
# =====================================================================


def _sync_sections_for_uuid(UUID, title, file_name, json_data, sections, storage_path,
                            text_key="Abstract Text identified:", overview_path=None,
                            overwrite=False):
    """
    Iterates through layout mappings to transform and evaluate section lists or
    strings. Returns a stats dict ``{section_key: {"true": t, "false": f}}`` that
    records, per section, how many extracted items were grounded in (a subset of)
    the reference text (``text_key`` -- the identified abstract or introduction
    text). ``overview_path`` selects the master-overview workbook to update.

    ``overwrite=True`` rewrites the Excel rows of already-evaluated UUIDs in
    place with the freshly computed values (used by ``mode="generate"``);
    ``overwrite=False`` (default, previous behaviour) leaves existing rows
    untouched and only appends UUIDs not yet in the sheets.
    """
    stats = {}
    for key, (ex_path, _) in sections.items():
        content_value = json_data.get(key, "Not Found")
        reference_text = json_data.get(text_key, "Not Found")

        if isinstance(content_value, list):
            t, f = _save_list_section(
                UUID=UUID,
                key=key,
                content_list=content_value,
                ex_path=ex_path,
                title=title,
                file_name=file_name,
                abs_txt=reference_text,
                storage_path=storage_path,
                overview_path=overview_path,
                overwrite=overwrite
            )
        else:
            t, f = _save_single_section(
                UUID=UUID,
                key=key,
                content_value=content_value,
                ex_path=ex_path,
                title=title,
                file_name=file_name,
                abs_txt=reference_text,
                storage_path=storage_path,
                overview_path=overview_path,
                overwrite=overwrite
            )
        stats[key] = {"true": t, "false": f}
    return stats


def _save_list_section(UUID, key, content_list, ex_path, title, file_name, abs_txt, storage_path, overview_path=None, overwrite=False):
    """
    Flattens lists horizontally using alternation layouts and uploads aggregate
    statistics. Always computes the true/false subset counts (returned for SQL),
    but only writes the Excel sheet when this UUID has not already been evaluated
    -- unless ``overwrite=True``, which updates the existing row in place
    (_write_section_sheet_flat upserts by UUID).
    """
    entry = {
        "UUID": str(UUID),
        "Title": title,
        "Filename": file_name,
    }

    bullet_lines = []
    true_count = 0
    false_count = 0

    for idx, item in enumerate(content_list):
        content_str = str(item)
        is_subset = content_str.lower() in str(abs_txt).lower() if abs_txt != "Not Found" else False

        if is_subset:
            true_count += 1
        else:
            false_count += 1

        # Populate alternating structural sequence columns
        count_label = idx + 1
        # entry[f"{key} {count_label}"] = f"{key} {count_label}"
        entry[f"{key} {count_label} value"] = content_str
        entry[f"{key} {count_label} Is_Subset"] = is_subset

        bullet_lines.append(f"- {content_str}")

    # Skip the Excel write if already evaluated (unless overwriting), but
    # still return the counts.
    if overwrite or not _is_duplicate_in_sheet(Path(ex_path), key, str(UUID)):
        _write_section_sheet_flat(Path(ex_path), key, entry)
        aggregate_bullets = "\n ".join(bullet_lines)
        _update_master_overview(storage_path, key, UUID, title, file_name, aggregate_bullets, true_count, false_count, entry, overview_path=overview_path)
        print(Fore.GREEN + f"✅ Evaluated flat list '{key}' ({len(content_list)} items) for UUID: {UUID}")

    return true_count, false_count


def _save_single_section(UUID, key, content_value, ex_path, title, file_name, abs_txt, storage_path, overview_path=None, overwrite=False):
    """
    Saves single text objects cleanly mapping their matching criteria directly.
    Always computes the true/false subset counts (returned for SQL), but only
    writes the Excel sheet when this UUID has not already been evaluated --
    unless ``overwrite=True``, which updates the existing row in place.
    """
    content_str = str(content_value)
    is_subset = content_str.lower() in str(abs_txt).lower() if abs_txt != "Not Found" else False

    true_count = 1 if is_subset else 0
    false_count = 0 if is_subset else 1

    entry = {
        "UUID": str(UUID),
        "Title": title,
        "Filename": file_name,
        "Content": content_str,
        "Is_Subset": is_subset
    }

    if overwrite or not _is_duplicate_in_sheet(Path(ex_path), key, str(UUID)):
        _write_section_sheet_flat(Path(ex_path), key, entry)
        # Save overview text and append sectional layout data onto a unique sheet inside Master_Overview.xlsx
        _update_master_overview(storage_path, key, UUID, title, file_name, content_str, true_count, false_count, entry, overview_path=overview_path)
        print(Fore.GREEN + f"✅ Evaluated single_section '{key}' for UUID: {UUID}")

    return true_count, false_count


def _eval_score(stats):
    """Overall subset-coverage from per-section stats: (percent, total_true, total_false)."""
    total_true = sum(int(v.get("true", 0)) for v in stats.values())
    total_false = sum(int(v.get("false", 0)) for v in stats.values())
    total = total_true + total_false
    percent = round(100.0 * total_true / total, 1) if total else None
    return percent, total_true, total_false


# target -> (SQL json column, SQL score column) for the evaluation summary.
_EVAL_SQL_COLUMNS = {
    "abstract": ("evaluation_json", "evaluation_score"),
    "intro": ("intro_evaluation_json", "intro_evaluation_score"),
}


def _push_eval_to_sql(uuid, stats, db_path=None, target="abstract"):
    """
    Write the evaluation summary (per-section subset counts + overall score) into
    the SQLite document row, if that row exists. ``target`` selects the column
    pair (abstract vs intro). Returns True on update.
    """
    import json as _json
    from alr.common.sql_store import AnalyzedDataStore, DB_PATH

    json_col, score_col = _EVAL_SQL_COLUMNS[target]
    store = AnalyzedDataStore(db_path or DB_PATH)
    if not store.get_document(uuid):
        return False
    percent, _t, _f = _eval_score(stats)
    store.update_document(uuid, {
        json_col: _json.dumps(stats),
        score_col: "" if percent is None else str(percent),
    })
    return True


def _existing_evaluation(uuid, db_path=None, target="abstract"):
    """
    Return the stored per-section stats for ``uuid`` if this document has already
    been evaluated for ``target`` (its SQL score column is present), else ``None``.
    Used by ``mode="copy"`` to reuse a prior evaluation instead of recomputing.
    """
    import json as _json
    from alr.common.sql_store import AnalyzedDataStore, DB_PATH

    json_col, score_col = _EVAL_SQL_COLUMNS[target]
    try:
        store = AnalyzedDataStore(db_path or DB_PATH)
        row = store.get_document(uuid)
    except Exception:
        return None
    if not row:
        return None
    score = row.get(score_col)
    if score is None or str(score).strip() in ("", "nan"):
        return None
    raw = row.get(json_col)
    if raw:
        try:
            return _json.loads(raw)
        except (ValueError, TypeError):
            pass
    return {}  # score present but no detail: still counts as "already evaluated"


def _load_intro_json(MF, uuid):
    """Load {uuid}_Intro.json from the space's Introduction data folder (or None)."""
    import json as _json

    path = Path(os.path.join(MF.AD_Intro, f"{uuid}_Intro.json"))
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return _json.load(f)
    except (OSError, _json.JSONDecodeError) as e:
        print(Fore.YELLOW + f"⚠️ Could not read intro JSON for {uuid}: {e}")
        return None


def _load_recorded_intros(MF):
    """UUIDs with a recorded introduction analysis (mirrors _load_recorded_abstracts)."""
    from alr.common.excel_utils import extract_column

    if MF.AD_Intro_log_path and Path(MF.AD_Intro_log_path).exists():
        return extract_column(MF.AD_Intro_log_path, "UUID")
    print(Fore.RED + "Introduction Log is not Available")
    return []


def evaluate_document(storage_path, uuid, db_path=None, push_sql=True, mode="generate", target="abstract"):
    """
    Evaluate one analyzed document (by ``uuid``): write its per-section evaluation
    Excel sheets and, when ``push_sql`` is set, mirror the summary into the SQLite
    store so overviews can use it.

    ``target`` selects what is evaluated: ``"abstract"`` (default) grounds the 7
    abstract sections against the identified abstract text (SQL columns
    ``evaluation_json`` / ``evaluation_score``); ``"intro"`` grounds the analyzed
    introduction keys (Background, Motivation, Gaps & Limitations, RQs & Scope)
    against the identified introduction text (SQL columns
    ``intro_evaluation_json`` / ``intro_evaluation_score``).

    ``mode="copy"`` reuses a prior evaluation when one already exists in SQL
    (skips recomputation; unevaluated documents are still computed);
    ``mode="generate"`` (default) always recomputes AND rewrites the Excel rows
    of already-evaluated UUIDs in place with the fresh values.

    Returns the per-section stats dict, or ``None`` if the source JSON is missing.
    Safe to call repeatedly: the Excel writes upsert per UUID (no duplicates).
    """
    from alr.common.sections import build_intro_sections_eval_map, INTRO_TEXT_KEY, ABSTRACT_TEXT_KEY

    if mode == "copy":
        prior = _existing_evaluation(uuid, db_path, target=target)
        if prior is not None:
            print(Fore.YELLOW + f"⏭️ Reusing existing {target} evaluation for {uuid} (copy mode).")
            return prior

    MF = storage_path if isinstance(storage_path, DataAnalyzeManager) else DataAnalyzeManager(storage_path)
    VDB = Vec_DB_Manager(MF.folder)

    if target == "intro":
        sections = build_intro_sections_eval_map(VDB)
        text_key = INTRO_TEXT_KEY
        overview_path = VDB.Introduction_Eval_Overview
        loader = _load_intro_json
    else:
        sections = build_sections_eval_map(VDB)
        text_key = ABSTRACT_TEXT_KEY
        overview_path = VDB.Abstract_Eval_Overview
        loader = _load_abstract_json

    MF.update_id_files(uuid)
    title, file_name = _fetch_metadata(MF, uuid)
    json_data = loader(MF, uuid)
    if not json_data:
        return None

    stats = _sync_sections_for_uuid(
        UUID=uuid, title=title, file_name=file_name,
        json_data=json_data, sections=sections, storage_path=str(MF.folder),
        text_key=text_key, overview_path=overview_path,
        overwrite=(mode == "generate"),
    )
    if push_sql:
        try:
            _push_eval_to_sql(uuid, stats, db_path, target=target)
        except Exception as e:
            print(Fore.YELLOW + f"⚠️ Could not push {target} evaluation to SQL for {uuid}: {e}")
    return stats


def evaluate_space(storage_path, db_path=None, progress_callback=None, should_cancel=None,
                   mode="generate", target="abstract") -> int:
    """
    Evaluate every analyzed document in a storage space (build the evaluation Excel
    DBs and sync the summary into SQLite). Returns the number of documents evaluated.

    ``target`` selects the data evaluated: ``"abstract"`` (default, documents from
    the abstract log) or ``"intro"`` (documents from the introduction log).
    ``mode="copy"`` reuses prior evaluations already recorded in SQL (skips
    recompute); ``mode="generate"`` (default) recomputes. ``progress_callback(done,
    total)`` is called after each document if given, and ``should_cancel`` is
    checked before each for cooperative cancellation.
    """
    MF = storage_path if isinstance(storage_path, DataAnalyzeManager) else DataAnalyzeManager(storage_path)

    recorded = _load_recorded_intros(MF) if target == "intro" else _load_recorded_abstracts(MF)
    if not recorded:
        return 0

    count = 0
    total = len(recorded)
    for i, uuid in enumerate(recorded, 1):
        if should_cancel is not None and should_cancel():
            print("Evaluation cancelled by user.")
            break
        try:
            evaluate_document(MF, uuid, db_path=db_path, push_sql=True, mode=mode, target=target)
            count += 1
        except Exception as e:
            print(Fore.YELLOW + f"⚠️ {target.title()} evaluation failed for {uuid}: {e}")
        if progress_callback:
            progress_callback(i, total)

    print(Fore.GREEN + f"✅ Evaluated {count} document(s) ({target}); results synced to SQL.")
    return count


def generate_databases(Storage_path):
    """Back-compatible entry point: evaluate the whole space (Excel DBs + SQL sync)."""
    return evaluate_space(Storage_path)


def generate_combined_databases(Source_path, Storage_path):
    MF = DataAnalyzeManager(Source_path)
    VDB = Vec_DB_Manager(Storage_path)

    recorded_abstracts = _load_recorded_abstracts(MF)
    if not recorded_abstracts:
        return

    sections = build_sections_eval_map(VDB)

    for UUID in recorded_abstracts:
        MF.update_id_files(UUID)

        title, file_name = _fetch_metadata(MF, UUID)
        json_data = _load_abstract_json(MF, UUID)
        if not json_data:
            continue

        _sync_sections_for_uuid(
            UUID=UUID,
            title=title,
            file_name=file_name,
            json_data=json_data,
            sections=sections,
            storage_path=Storage_path
        ) 

if __name__ == "__main__":
    Source_paths=[
        '/remotedata/U/DLR+kata_du/ALR DATA/AI_RM/AI_REQ_Results',    
        '/remotedata/U/DLR+kata_du/ALR DATA/AI_SE_Domains_main/AI_SE_Processed_results',
        '/remotedata/U/DLR+kata_du/ALR DATA/LLM_Safety/LLM_Safety_Results',
        '/remotedata/U/DLR+kata_du/ALR DATA/MBSE_MBSA_Aviation/MBSE_MBSA_Aviation_Results',
        '/remotedata/U/DLR+kata_du/ALR DATA/Only_MBSA/Only_MBSA_results'
    ]
    storage_path='/remotedata/U/DLR+kata_du/ALR DATA/00_Container/Combined_DB/AI_SE_Domains'
    storage_path='/remotedata/U/DLR+kata_du/ALR DATA/Only_MBSA/Only_MBSA_results'
    generate_databases(storage_path)

    # generate_databases(storage_path)
    # for S_path in Source_paths:
    #     generate_combined_databases(S_path, storage_path)