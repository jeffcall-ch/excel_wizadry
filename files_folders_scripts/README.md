# files_folders_scripts

## Purpose
This folder contains scripts for analyzing and copying files based on system codes listed in CSV files.
- `copy_unique_short_system_code_files.py`: Copies files listed in a CSV to a target directory.
- `count_unique_system_codes.py`: Counts unique system codes in a CSV file and prints the result.

**Input:**
- CSV files (e.g., `pdf_file_list_unique_short_system_code.csv`, `pdf_file_list.csv`) with file paths and system codes.

**Output:**
- Copied files (to a specified directory).
- Printed count of unique system codes.

## Usage
1. Ensure the required CSV files are present in this folder.
2. Run the scripts from the command line:
   ```powershell
   python copy_unique_short_system_code_files.py
   python count_unique_system_codes.py
   ```

## Examples
```powershell
python copy_unique_short_system_code_files.py
python count_unique_system_codes.py
```

## Known Limitations
- File paths and destination directories are hardcoded; edit the script to change them.
- CSV format must match expected columns (e.g., 'full_path', 'System_code').
- No error handling for missing files or invalid paths.
- Large file lists may impact performance.
