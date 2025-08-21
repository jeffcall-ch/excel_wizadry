# KKS_search_in_excel_files

## Purpose
This folder contains a tool for searching KKS codes in Excel files and exporting the results to CSV.
- `KKS_search_in_excel_files.py`: Recursively searches Excel files in a directory for specified KKS codes and outputs matching rows to a CSV file.

**Input:**
- Directory containing Excel files (`.xlsx`, `.xlsm`, `.xls`).
- KKS code(s) to search for.

**Output:**
- CSV file listing matches (file, sheet, row, context).

## Usage
1. Place your Excel files in a target directory.
2. Run the script from the command line:
   ```powershell
   python KKS_search_in_excel_files.py --directory <path_to_excel_files> --kks <KKS_code>
   ```
   (See script for additional arguments.)

## Examples
```powershell
python KKS_search_in_excel_files.py --directory C:\data\excels --kks 2QFB94BR110
```

## Known Limitations
- Only supports `.xlsx`, `.xlsm`, `.xls` files.
- Assumes KKS codes are present in the first column or header row.
- Performance may degrade with large directories.
- Output CSV path may be hardcoded; check script for details.
