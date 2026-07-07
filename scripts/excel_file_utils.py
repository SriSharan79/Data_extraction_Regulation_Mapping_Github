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
    column_name,
    new_sheet_name="Unique Elements",
    separator=";",
):
    try:
        # 1. Read the Excel file to extract data
        df = pd.read_excel(file_path)

        # Check if the column exists
        if column_name not in df.columns:
            print(f"Error: Column '{column_name}' not found in the Excel file.")
            return

        unique_elements = set()

        # Extract unique elements
        for row in df[column_name].dropna():
            elements = [el.strip() for el in str(row).split(separator)]
            unique_elements.update(elements)

        # Convert the sorted unique elements into a DataFrame
        unique_df = pd.DataFrame(
            sorted(list(unique_elements)), columns=[f"Unique_{column_name}"]
        )

        # 2. Append the new sheet to the same Excel file
        # 'mode="a"' appends to the file, and 'if_sheet_exists="replace"'
        # overwrites the sheet if you run the script multiple times.
        with pd.ExcelWriter(
            file_path, mode="a", engine="openpyxl", if_sheet_exists="replace"
        ) as writer:
            unique_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

        print(
            f"Success! Extracted {len(unique_elements)} unique elements and saved them "
            f"into the sheet '{new_sheet_name}' inside '{file_path}'."
        )

    except Exception as e:
        print(f"An error occurred: {e}")


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
    
    #     # Replace with your actual file path, column name, and desired new sheet name
    # EXCEL_FILE = r"U:\ALR DATA\AI_SE-Domains_pdfs\IEEE Excels on the Content in IEEE Explore\IEEEXplore_Global_All-Conference-Series.xlsx"
    # COLUMN_TO_PROCESS = "subjects"
    # NEW_SHEET = "Unique Categories"
    
    # save_unique_elements_to_new_sheet(
    #     EXCEL_FILE, COLUMN_TO_PROCESS, new_sheet_name=NEW_SHEET
    # )
    
    # unique_results = get_unique_elements(EXCEL_FILE, COLUMN_TO_PROCESS)

    # print(f"\nFound {len(unique_results)} unique elements:\n")
    # for item in sorted(unique_results):
    #     print(f"- {item}")

    # Update these paths and column names to match your environment
    EXCEL_FILE_PATH = r"U:\ALR DATA\IEEE Systems Conference\publications_metadata.xlsx"
    FILENAME_COLUMN = "File_Name"  # e.g., 'document_1.pdf'
    CATEGORY_COLUMN = "Container"  # e.g., 'Computing and Processing'

    SOURCE_FOLDER = r"U:\ALR DATA\IEEE Systems Conference\failed_pdfs"
    STORAGE_FOLDER = r"U:\ALR DATA\IEEE Systems Conference"

    organize_pdf_files(
        excel_path=EXCEL_FILE_PATH,
        file_col=FILENAME_COLUMN,
        category_col=CATEGORY_COLUMN,
        source_dir=SOURCE_FOLDER,
        target_dir=STORAGE_FOLDER,
    )
