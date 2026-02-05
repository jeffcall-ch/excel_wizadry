import os
import csv
import shutil
import re
from pathlib import Path

def extract_years_from_dates(dates_str):
    """
    Extract only years (4-digit numbers like 2021) from date strings.
    Returns a sorted list of unique years.
    """
    if not dates_str or dates_str.strip() == "":
        return []
    
    # Pattern to find 4-digit years (1900-2099)
    year_pattern = r'\b(19\d{2}|20\d{2})\b'
    years = re.findall(year_pattern, dates_str)
    
    # Remove duplicates and sort
    unique_years = sorted(set(years))
    return unique_years

def create_folder_name(years):
    """
    Create folder name based on years.
    - No years: "no year"
    - One year: "2021"
    - Multiple years: "2021_2022_2023"
    """
    if not years:
        return "no year"
    return "_".join(years)

def organize_files():
    """
    Organize PDF files from CSV into categorized folders.
    """
    # Configuration
    base_folder = r"C:\temp\WERK\FERO CALCS"
    csv_file = "fero_pipe_class_files.csv"
    
    # Pipe class codes
    pipe_class_codes = [
        "EFDX", "EFDJ", "EHDX", "EEDX", "ECDE", "ECDM", "EHDQ", "EEDQ",
        "EHGN", "EHFD", "EEFD", "AHDX", "AEDX", "ACDE", "ACDM", "AHDQ",
        "AHFD", "AHGN"
    ]
    
    # Get script directory
    script_dir = Path(__file__).parent
    csv_path = script_dir / csv_file
    
    print(f"Reading CSV file: {csv_path}")
    print(f"Organizing files into: {base_folder}")
    print("=" * 80)
    
    # Check if CSV exists
    if not csv_path.exists():
        print(f"ERROR: CSV file not found: {csv_path}")
        return
    
    # Create base folder if it doesn't exist
    os.makedirs(base_folder, exist_ok=True)
    
    # Create subfolders for each pipe class code
    print("\nCreating code subfolders...")
    for code in pipe_class_codes:
        code_folder = os.path.join(base_folder, code)
        os.makedirs(code_folder, exist_ok=True)
        print(f"  Created: {code}")
    
    print("\n" + "=" * 80)
    print("Processing files...\n")
    
    # Read CSV and process files
    files_processed = 0
    files_copied = 0
    files_skipped = 0
    
    with open(csv_path, 'r', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        
        for row in reader:
            files_processed += 1
            
            file_path = row['file_path']
            found_codes = row['found_codes'].split(', ')
            dates_in_path = row['dates_in_path']
            
            # Extract years from path
            years = extract_years_from_dates(dates_in_path)
            folder_name = create_folder_name(years)
            
            print(f"[{files_processed}] {os.path.basename(file_path)}")
            print(f"     Codes: {', '.join(found_codes)}")
            print(f"     Years in path: {', '.join(years) if years else 'none'}")
            print(f"     Folder: {folder_name}")
            
            # Check if source file exists
            if not os.path.exists(file_path):
                print(f"     ✗ Source file not found!")
                files_skipped += 1
                continue
            
            # Copy file to each relevant code folder
            for code in found_codes:
                code = code.strip()
                if code in pipe_class_codes:
                    # Create destination folder
                    dest_folder = os.path.join(base_folder, code, folder_name)
                    os.makedirs(dest_folder, exist_ok=True)
                    
                    # Copy file
                    dest_file = os.path.join(dest_folder, os.path.basename(file_path))
                    
                    try:
                        shutil.copy2(file_path, dest_file)
                        print(f"     ✓ Copied to: {code}\\{folder_name}\\")
                        files_copied += 1
                    except Exception as e:
                        print(f"     ✗ Error copying to {code}: {str(e)}")
                        files_skipped += 1
            
            print()
    
    # Summary
    print("=" * 80)
    print(f"SUMMARY:")
    print(f"  Files processed: {files_processed}")
    print(f"  Files copied: {files_copied}")
    print(f"  Files skipped: {files_skipped}")
    print(f"  Output folder: {base_folder}")
    print("=" * 80)
    print("✓ Organization complete!")

if __name__ == "__main__":
    organize_files()
