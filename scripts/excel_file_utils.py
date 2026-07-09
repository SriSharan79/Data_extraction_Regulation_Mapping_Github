import pandas as pd


def get_unique_elements(file_path, column_name, separator=";"):
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Check if the column exists in the file
        if column_name not in df.columns:
            print(f"Error: Column '{column_name}' not found in the Excel file.")
            return set()

        unique_elements = set()

        # Drop NaN values and iterate through the rows of the specified column
        for row in df[column_name].dropna():
            # Split the row content by the separator and strip whitespace from each element
            elements = [el.strip() for el in str(row).split(separator)]
            # Add elements to the set (sets automatically handle uniqueness)
            unique_elements.update(elements)

        return unique_elements

    except Exception as e:
        print(f"An error occurred: {e}")
        return set()

def save_unique_elements_to_new_sheet(
    file_path,
    column_names,
    new_sheet_name="Unique Elements",
    source_sheet=0,
):
    """Collect the unique elements of one or more columns and write them to
    ``new_sheet_name`` as `Unique_<col> | Count_<col>` pairs (each element with
    how often it occurs next to it).

    ``column_names`` may be a single column name or a list of them;
    ``source_sheet`` picks the sheet to read (index or name). The target sheet
    is rebuilt from scratch on every call, so the function is safe to re-run
    after each new data row (idempotent).  Returns True on success.
    """
    try:
        from collections import Counter

        # 1. Read the source sheet to extract data
        df = pd.read_excel(file_path, sheet_name=source_sheet)

        if isinstance(column_names, str):
            column_names = [column_names]

        # Define the list of candidate separators to check
        candidate_separators = [',',  ';']

        blocks = []
        for column_name in column_names:
            if column_name not in df.columns:
                print(f"Error: Column '{column_name}' not found in the Excel file.")
                continue

            counts = Counter()

            # 2. Extract elements with dynamic separator detection, counting
            # every occurrence so the sheet can show element + count.
            for row in df[column_name].dropna():
                text = str(row)

                # Count how many times each separator appears in the current text
                sep_counts = {sep: text.count(sep) for sep in candidate_separators}

                # Find the separator that has the highest count
                best_sep = max(sep_counts, key=sep_counts.get)

                # If the most frequent separator actually exists in the string (count > 0)
                if sep_counts[best_sep] > 0:
                    # Split by the best separator and remove empty strings/whitespace
                    elements = [el.strip() for el in text.split(best_sep) if el.strip()]
                else:
                    # If none of the separators are in the string, treat it as a single element
                    elements = [text.strip()] if text.strip() else []

                counts.update(elements)

            blocks.append(pd.DataFrame(
                sorted(counts.items()),
                columns=[f"Unique_{column_name}", f"Count_{column_name}"],
            ))

        if not blocks:
            return False

        # 3. Columns of different lengths sit side by side, padded with blanks;
        # keep the counts as (nullable) integers.
        target_df = pd.concat(blocks, axis=1)
        for col in target_df.columns:
            if col.startswith("Count_"):
                target_df[col] = target_df[col].astype("Int64")

        # 4. Replace the target sheet with the fresh unique-element table
        with pd.ExcelWriter(
            file_path, mode="a", engine="openpyxl", if_sheet_exists="replace"
        ) as writer:
            target_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

        print(
            f"Success! Wrote unique elements (+counts) for "
            f"{[c for c in target_df.columns if c.startswith('Unique_')]} "
            f"into the sheet '{new_sheet_name}'."
        )
        return True

    except Exception as e:
        print(f"An error occurred: {e}")
        return False


import os
import shutil
from pathlib import Path
import pandas as pd


def organize_pdf_files(
    excel_path, file_col, category_col, source_dir, target_dir
):
    try:
        # 1. Load Excel and build mapping
        df = pd.read_excel(excel_path)

        # Clean data: drop rows where filename or category is missing
        df = df[[file_col, category_col]].dropna()

        # Create a dictionary mapping: { filename: target_category_folder }
        # Stripping spaces and enforcing standard string format
        file_to_category = {
            str(row[file_col]).strip(): str(row[category_col]).strip()
            for _, row in df.iterrows()
        }

        print(
            f"Loaded {len(file_to_category)} file-to-category mappings from Excel."
        )

        # 2. Recursively scan source folder for PDF files
        print(f"Scanning '{source_dir}' recursively for PDFs...")
        source_path = Path(source_dir)
        found_files_count = 0

        # rglob handles recursive searching (looks inside all subfolders)
        for file_path in source_path.rglob("*.pdf"):
            filename = file_path.name

            # Check if this PDF is one of the files listed in our Excel mapping
            if filename in file_to_category:
                category_name = file_to_category[filename]

                # Define and create the target category folder safely
                category_folder = Path(target_dir) / category_name
                category_folder.mkdir(parents=True, exist_ok=True)

                # Define final destination path
                destination_path = category_folder / filename

                # Copy the file (use shutil.move if you want to cut-and-paste instead)
                shutil.copy2(file_path, destination_path)
                found_files_count += 1
                print(f"Copied: {filename} -> [{category_name}]")

        print("\n--- Processing Complete ---")
        print(
            f"Successfully organized {found_files_count} files into '{target_dir}'."
        )

    except Exception as e:
        print(f"An error occurred: {e}")


# --- Configuration ---
if __name__ == "__main__":
    
        # Replace with your actual file path, column name, and desired new sheet name
    EXCEL_FILE = r"C:\Users\kata_du\Documents\Literature\EASA\XML _Data_extractions\EAR for CS-25 Amdt 27 (xml) fix 12.22 FINAL (1)\CS-25 _Section_analysis_by AI.xlsx"
    COLUMNS_TO_PROCESS = [
                        "System Info",
                        "Personal Involved",
                        "References"
                         ]
    NEW_SHEET = "Unique Categories"
    # One call with all columns: the sheet is rebuilt per call, so per-column
    # looping would leave only the last column in it.
    save_unique_elements_to_new_sheet(
        EXCEL_FILE, COLUMNS_TO_PROCESS, new_sheet_name=NEW_SHEET
    )
    for col in COLUMNS_TO_PROCESS:
        unique_results = get_unique_elements(EXCEL_FILE, col)

        print(f"\nFound {len(unique_results)} unique elements:\n")
        for item in sorted(unique_results):
            print(f"- {item}")

    # Update these paths and column names to match your environment
    # EXCEL_FILE_PATH = r"U:\ALR DATA\IEEE Systems Conference\publications_metadata.xlsx"
    # FILENAME_COLUMN = "File_Name"  # e.g., 'document_1.pdf'
    # CATEGORY_COLUMN = "Container"  # e.g., 'Computing and Processing'

    # SOURCE_FOLDER = r"U:\ALR DATA\IEEE Systems Conference\failed_pdfs"
    # STORAGE_FOLDER = r"U:\ALR DATA\IEEE Systems Conference"

    # organize_pdf_files(
    #     excel_path=EXCEL_FILE_PATH,
    #     file_col=FILENAME_COLUMN,
    #     category_col=CATEGORY_COLUMN,
    #     source_dir=SOURCE_FOLDER,
    #     target_dir=STORAGE_FOLDER,
    # )
