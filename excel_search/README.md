# Excel File Search Tool

This script searches for a given phrase in all Excel and CSV files within a specified directory (including subdirectories). It finds all partial, case-insensitive matches in all sheets (for Excel files) and all rows/columns (for CSV files), then generates an Excel report summarizing the results.

## Features
- Supports `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, and `.csv` files
- Searches all sheets in Excel files and all rows/columns in CSV files
- Case-insensitive, partial match search
- Outputs an Excel report listing: file name, sheet name, cell address (or row/column for CSV), match context, and hyperlinks to the files
- Includes a summary sheet in the report

## Requirements
- Python 3.8+
- Packages: `pandas`, `openpyxl`, `xlrd`

Install dependencies:
```
pip install pandas openpyxl xlrd
```

## Usage
Run the script from the command line:
```
python excel_file_search.py <search_phrase> <directory> [--output <output_file>]
```
- `<search_phrase>`: The text to search for (case-insensitive, partial matches)
- `<directory>`: The folder containing Excel/CSV files to search (searches recursively)
- `--output <output_file>`: (Optional) Path for the Excel report. Default is `excel_search_results_<date>.xlsx`.


### Virtual Environment
Before running the script, make sure to activate your Python virtual environment (venv):
```
Windows:
    .venv\Scripts\activate
Linux/macOS:
    source .venv/bin/activate
```

### Example
```
C:/Users/szil/Repos/excel_wizadry/.venv/Scripts/python.exe excel_search/excel_file_search.py "YOURSEARCHTERM" "C:\Users\szil\OneDrive - Kanadevia Inova\Desktop\Projects\03 AbuDhabi\01 Documents\05 BOM Lists\250804\KVI SK-KVI AG-01943" --output "C:\Users\szil\OneDrive - Kanadevia Inova\Desktop\valve_search_results.xlsx"

```

## Output
- The Excel report will contain:
  - File name
  - Sheet name (or 'CSV' for CSV files)
  - Cell address (or row/column for CSV)
  - Match context (cell value)
  - Hyperlink to the file
  - Summary sheet with search details

## Notes
- Hyperlinks in the report open the file, but may not jump to the exact cell/row.
- All matches are reported; there is no limit per file/sheet/cell.

## Troubleshooting
- If you encounter permission errors on Windows, ensure no Excel files are open during the search.
- For large folders, the search may take some time.

## License
MIT License
