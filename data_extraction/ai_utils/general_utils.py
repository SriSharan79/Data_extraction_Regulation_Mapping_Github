import shutil
import pandas as pd
import difflib
import re
from pathlib import Path
import time
import os
import hashlib
import datetime

def generate_unique_id(filename, existing_ids):
    """
    Generate hash-based ID, append suffix if collision detected.
    """
    base_hash = hashlib.md5(filename.encode()).hexdigest()[:8]
    candidate = base_hash
   
    return candidate

def clean_folder_path(raw_folder_path):
        raw = raw_folder_path.strip()
        # Remove surrounding quotes if the user pasted "..."
        if (raw.startswith('"') and raw.endswith('"')) or (raw.startswith("'") and raw.endswith("'")):
            raw = raw[1:-1].strip()

        # Don't force as_posix() for existence checks
        return str(Path(raw).expanduser())

# def generate_unique_id(filename, existing_ids):
#     """
#     Generate hash-based ID, append suffix if collision detected.
#     """
#     base_hash = hashlib.md5(filename.encode()).hexdigest()[:8]
#     candidate = base_hash
    
#     counter = 1
#     while candidate in existing_ids:
#         candidate = f"{base_hash}_{counter}"
#         counter += 1
    
#     return candidate

def add_hh_mm_ss(time_str1, time_str2):
    # 1. Convert HH:MM:SS strings to timedelta objects
    h1, m1, s1 = map(int, time_str1.split(':'))
    delta1 = datetime.timedelta(hours=h1, minutes=m1, seconds=s1)
    
    h2, m2, s2 = map(int, time_str2.split(':'))
    delta2 = datetime.timedelta(hours=h2, minutes=m2, seconds=s2)
    
    # 2. Add the timedeltas
    total_delta = delta1 + delta2
    
    # 3. Format back to HH:MM:SS
    total_seconds = int(total_delta.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    
    return f"{hours:02}:{minutes:02}:{seconds:02}"
def _as_path(p):
    """Convert str/pathlike/Path into Path safely."""
    if isinstance(p, Path):
        return 
    return Path(str(p))

def caluculate_time_taken(start_time,end_time):

    duration_seconds = end_time - start_time

    return time.strftime("%H:%M:%S", time.gmtime(duration_seconds))   

def print_two_column_table(items, Header, col_width=40):
    """
    Prints index--string pairs in two columns with table borders.
    """
    if not items:
        print("No data to display.")
        return

    n = len(items)
    mid = (n + 1) // 2  # split point

    left = items[:mid]
    right = items[mid:]

    def format_cell(idx, text):
        return f"{idx:>3} -- {text}"[:col_width].ljust(col_width)

    horizontal = "+" + "-" * (col_width + 2) + "+" + "-" * (col_width + 2) + "+"
    header = (
        "|" + f" {Header}".center(col_width * 2 + 2)+ "|"
    )

    print(horizontal)
    print(header)
    print(horizontal)

    for i in range(mid):
        left_cell = format_cell(i, left[i]) if i < len(left) else " " * col_width
        right_idx = i + mid
        right_cell = (
            format_cell(right_idx, right[i]) if i < len(right) else " " * col_width
        )

        print(f"| {left_cell} | {right_cell} |")

    print(horizontal)

def Proccess_string_to_list(api_response: str) -> list:
    """
    Cleans the raw AI string output and converts it into a 
    Python list of unique, non-empty keywords.
    """
    # Split by newline and strip whitespace/extra characters from each line
    keywords = [line.strip() for line in api_response.strip().split('\n')]
    
    # Filter out empty strings just in case of trailing newlines
    return [kw for kw in keywords if kw]

def is_similar(str1, str2, threshold=0.8):
    """
    Check if two strings are similar based on a given similarity threshold using word occurrence.
    Safely handles NaN/float/None inputs.
    """
    # Convert to strings safely, handling NaN/float/None
    str1 = str(str1).strip().lower() if not (pd.isna(str1) or isinstance(str1, (float, type(None)))) else ''
    str2 = str(str2).strip().lower() if not (pd.isna(str2) or isinstance(str2, (float, type(None)))) else ''
    
    if not str1 or not str2:
        return False
    
    # Use SequenceMatcher to compare word occurrence
    ratio = difflib.SequenceMatcher(None, str1.split(), str2.split()).ratio()
    return ratio >= threshold

# import difflib
# import pandas as pd

# def is_similar(str1, str2, threshold=0.8):
#     """
#     Check if two strings are similar based on a given similarity threshold using word occurrence.
#     Safely handles NaN/float/None inputs.
#     """
#     # Convert to strings safely, handling NaN/float/None
#     str1 = str(str1).strip().lower() if not (pd.isna(str1) or isinstance(str1, (float, type(None)))) else ''
#     str2 = str(str2).strip().lower() if not (pd.isna(str2) or isinstance(str2, (float, type(None)))) else ''
    
#     print(f"Comparing: '{str1}' with '{str2}'")  # Debugging line
#     if not str1 or not str2:
#         print("One of the strings is empty. Returning False.")  # Debugging line
#         return False
    
#     # Use SequenceMatcher to compare word occurrence
#     ratio = difflib.SequenceMatcher(None, str1.split(), str2.split()).ratio()
#     print(f"Similarity ratio: {ratio}")  # Debugging line
    
#     return ratio >= threshold


def is_similar_Length(str1, str2, threshold=0.7):
    """
    Check if two strings are similar based on a given similarity threshold.

    Args:
        str1 (str): The first string.
        str2 (str): The second string.
        threshold (float): The similarity threshold (default is 0.8).

    Returns:
        bool: True if the strings are similar, False otherwise.
    """
    # Simple similarity check using length-based comparison
    len_str1 = len(str1)
    len_str2 = len(str2)

    if len_str1 < len_str2:
        return len_str1 / len_str2 >= threshold
    else:
        return len_str2 / len_str1 >= threshold


def merge_lists(initial_list, new_list):
    """
    Merge two lists ensuring no duplicates or near-duplicates (strings matching 80% or more).

    Args:
        initial_list (list): The initial list of strings.
        new_list (list): The new list of strings to be merged.

    Returns:
        list: The merged list with no duplicates or near-duplicates.
    """
    # Use a list to store the final merged result
    merged_list = []

    for item in initial_list + new_list:
        is_duplicate = False

        # Check if the current item should not be added due to similarity or duplication
        for existing_item in merged_list:
            if item == existing_item or is_similar(item, existing_item):
                # print(f"{Fore.WHITE}Existing: {existing_item}")
                # print(f"{Fore.RED}Skipping duplicate/near-duplicate: {item}")
                is_duplicate = True
                break

        # If the item is not a duplicate or near-duplicate, add it to the merged list
        if not is_duplicate:
            merged_list.append(item)

    return merged_list

def find_in_first_n_chars(strings, target, n):
    """
    Returns a list of strings where `target` appears within
    the first `n` characters.
    """
    if not strings or not target or n <= 0:
        return []

    target_lower = target.lower()

    return [
        s for s in strings
        if target_lower in s[:n].lower()
    ]

def find_first_match_in_first_n_chars_old(strings, target, n):
    """
    Returns the first string where `target` appears within
    the first `n` characters. Returns None if no match is found.
    """
    if not strings or not target or n <= 0:
        return None

    for s in strings:
        for t in target:
            target_lower = t.lower()
            if target_lower in s[:n].lower():
                return s
            if target_lower in s.lower():
                return s

    return None

def find_first_match_in_first_n_chars(strings, target, n):
    """
    Returns a tuple of (prev_string, matched_string, next_string) for the first 
    string where `target` appears within the first `n` characters (or anywhere in the string 
    as per the original fallback logic). 
    
    Returns None if no match is found.
    """
    if not strings or not target or n <= 0:
        return None

    # Use enumerate to track the current index for slicing neighbors
    for i, s in enumerate(strings):
        for t in target:
            target_lower = t.lower()
            
            # Check if target is within first n characters OR anywhere in the string (fallback)
            if target_lower in s[:n].lower() or target_lower in s.lower():
                # Safely grab the previous string if it exists
                # prev_s = strings[i - 1] if i > 0 else None
                
                # Safely grab the next string if it exists
                next_s = strings[i + 1] if i < len(strings) - 1 else None
                
                return s+next_s

    return None
    
def find_missing_elements(list_1, list_2):
    return [item for item in list_1 if item not in list_2]

def remove_string_from_list(strings, target):
    return [string for string in strings if string != target]

def clean_response_json_text(text):
    # Step 1: Remove unwanted markers or anything before the curly braces
    # Look for content inside curly braces and remove extra text outside
    
    match = re.search(r'(\{.*\})', text, re.DOTALL) # Look for content between the first and last curly brace
    
    if match:
        cleaned_text = match.group(0)  # Get the content inside the curly braces
        return  cleaned_text
    else:
        print("Invalid JSON format detected!")
        raise ValueError("No valid JSON-like structure found!")


# def print_with_separator(input_string, separator="*", terminal_width=None):
#     """
#     Prints the input string with the separator before and after it, ensuring the separator covers the entire terminal width.
    
#     Args:
#     input_string (str): The string to be printed in the middle of the separator.
#     separator (str): The character(s) used for the separator. Default is "*".
#     terminal_width (int): The width of the terminal. If not provided, it will be detected automatically.
#     """
#     if terminal_width is None:
#         # Get the current terminal width
#         terminal_width = os.get_terminal_size().columns
    
#     # Calculate the space available for the input string
#     separator_length = (terminal_width - len(input_string) - 2) // len(separator)
    
#     # Create the separator string that will span the terminal width
#     # separator_string = separator * separator_length
#     separator_string = separator * terminal_width
    
#     # Print the result
#     print(f"\n{separator_string} \n"+f"{input_string}".center(terminal_width) +f"\n{separator_string}\n")



def print_with_separator(input_string, separator="*", terminal_width=None, default_width=120):
    """
    Safe in Jupyter, terminals, and PyInstaller:
    - avoids os.get_terminal_size() ioctl errors
    - falls back to default_width when width can't be detected
    """
    # Determine width safely
    if terminal_width is None:
        try:
            terminal_width = shutil.get_terminal_size(fallback=(default_width, 30)).columns
        except Exception:
            # Extra fallback: some environments set COLUMNS
            try:
                terminal_width = int(os.environ.get("COLUMNS", default_width))
            except Exception:
                terminal_width = default_width

    # Guardrails
    terminal_width = max(20, int(terminal_width))  # prevent tiny/invalid widths
    input_string = "" if input_string is None else str(input_string)
    separator = "*" if not separator else str(separator)

    separator_string = separator * terminal_width

    print(
        f"\n{separator_string}\n"
        f"{input_string.center(terminal_width)}\n"
        f"{separator_string}\n"
    )

