## comos_extract_text_to_number
### `pipe_category_calculator.py`
Description:
- **Purpose:** Automatically determines the safety category for industrial pipes based on European standards. This helps ensure that pipes are correctly classified according to the potential hazard of the fluid they carry.
- **Inputs:** 
    - An Excel file named `Export_pipe_list_30.07.2025_withfluid_code.xlsx` containing the pipe data (e.g., diameter, pressure, fluid type). This file must be in the same folder as the script.
    - A CSV file named `pipe categories according to EN13480-1 Table 5.1-1.csv` that contains the classification rules.
- **Outputs:** 
    - An Excel file with the `_categorized` suffix (e.g., `Export_pipe_list_30.07.2025_withfluid_code_categorized.xlsx`). This file is a copy of the input Excel file but with two new columns showing the calculated safety category and the reasoning.
    - A CSV file with the `_debug` suffix for technical analysis.
- **How to Run:** Execute the script directly from the command line without any arguments: `python pipe_category_calculator.py`.
- **Additional Notes:** The script `excel_text_to_number_bck.py` is an old backup and can be deleted.

Path: `/root/repos/excel_wizadry/comos_extract_text_to_number/pipe_category_calculator.py`
Keywords: pipe category, EN 13480-1, safety classification, automation

## dxf_iso_bom_extraction
### `parallel_dxf_extraction.py`
**Description:** This script serves as a high-performance, parallel processor for extracting Bill of Materials (BOM) data from a directory of DXF files. It takes a folder path as a command-line argument and recursively finds all `.dxf` files. Using Python's `multiprocessing` library, it distributes the workload across multiple CPU cores to significantly speed up the extraction process. The core extraction logic is imported from `dxf_iso_bom_extraction.py`. The script aggregates all extracted materials and cut lengths into three separate CSV files: `parallel_all_materials.csv`, `parallel_all_cut_lengths.csv`, and `parallel_summary.csv`, which includes performance metrics for the operation.
**Path:** `/root/repos/excel_wizadry/dxf_iso_bom_extraction/parallel_dxf_extraction.py`
**Keywords:** DXF, BOM, data extraction, parallel processing, multiprocessing, CSV

## excel_search
### `excel_file_search.py`
**Description:** This is a command-line utility designed to recursively search for a specific text string within all Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`) and CSV files in a given directory. The search is case-insensitive and finds partial matches. The script takes the search text and directory path as arguments. It outputs the results into a new, timestamped Excel file, which includes a summary tab and a results tab. The results tab lists every cell where the text was found, its location (file, sheet, cell), and includes a hyperlink directly to the source file for easy access.
**Path:** `/root/repos/excel_wizadry/excel_search/excel_file_search.py`
**Keywords:** excel, csv, search, text search, data discovery, reporting