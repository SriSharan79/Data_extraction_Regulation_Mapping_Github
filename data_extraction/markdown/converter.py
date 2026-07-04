import os
from pathlib import Path


def convert_pdf_to_markdown(pdf_path, output_md_path):
    # markitdown is heavyweight and only needed at conversion time; import it
    # lazily so this module loads even when markitdown is not installed.
    from markitdown import MarkItDown

    # Initialize the MarkItDown converter
    md = MarkItDown()
    
    try:
        # Convert the local PDF file
        print(f"Converting '{pdf_path}'...")
        result = md.convert_local(pdf_path)
        
        # Write the converted text content into a Markdown file
        with open(output_md_path, "w", encoding="utf-8") as f:
            f.write(result.text_content)
            
        print(f"Success! Markdown saved to '{output_md_path}'")
        
    except Exception as e:
        print(f"An error occurred during conversion: {e}")

# Example Usage
if __name__ == "__main__":
    
    source_path=r"U:\ALR DATA\00_Container\Extracted_Data\pdf_files"
    workspace_directory = r"U:\ALR DATA\00_Container\Extracted_Data\MarkDown_extraction"
    
    source_root = Path(source_path)
    
    # rglob("*.pdf") finds all PDFs in all subfolders
    for file_path in source_root.rglob("*.pdf"): 
        if os.path.exists(file_path):
            convert_pdf_to_markdown(file_path, workspace_directory)
        else:
            print(f"Error: Missing verification path file -> {file_path}")