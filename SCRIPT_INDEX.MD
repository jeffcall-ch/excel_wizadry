## comos_extract_text_to_number
### pipe_category_calculator.py
Description:
- **Purpose:** Calculates the safety category for each pipe based on EN 13480-1 standards.
- **Inputs:** Excel file with pipe data.
- **Outputs:** Updated Excel file with safety categories assigned.
- **How to Run:** Run `python pipe_category_calculator.py` from the command line. No arguments needed.
- **Additional Notes:** The script `excel_text_to_number_bck.py` is an old backup and can be deleted.

Path: c:\Users\szil\Repos\excel_wizadry\comos_extract_text_to_number\pipe_category_calculator.py  
Keywords: pipe category, EN 13480-1, safety classification, automation

## dxf_iso_bom_extraction
### async_dxf_extraction.py
Description:
- **Purpose:** Processes DXF files in a folder to extract bill of materials information using asynchronous operations for speed.
- **Inputs:** Folder containing DXF files.
- **Outputs:** Summary Excel or CSV files with extracted material and cut length data.
- **How to Run:** Run `python async_dxf_extraction.py <directory> [--concurrent N] [--debug]` from the command line. Replace `<directory>` with the folder path. Optional: `--concurrent N` to set number of parallel tasks, `--debug` for detailed output.
- **Additional Notes:** Other scripts like `demo_parallel.py` and `parallel_dxf_extraction.py` are older or alternative versions.

Path: c:\Users\szil\Repos\excel_wizadry\dxf_iso_bom_extraction\async_dxf_extraction.py  
Keywords: DXF, bill of materials, async, extraction

## excel_search
### excel_file_search.py
Description:
- **Purpose:** Searches for a specific text string in all Excel and CSV files within a folder.
- **Inputs:** Search text and folder path.
- **Outputs:** Timestamped Excel file listing all matches, with hyperlinks to source files.
- **How to Run:** Run `python excel_file_search.py <search_text> <directory>` from the command line.
- **Additional Notes:** None.

Path: c:\Users\szil\Repos\excel_wizadry\excel_search\excel_file_search.py  
Keywords: excel, csv, search, text search, reporting

## line_list_check
### line_list_check_pipe_class.py
Description:
- **Purpose:** Checks and validates pipe class specifications in a line list Excel file against reference standards for industrial piping.
- **Inputs:** Excel file with line list data (`Export 17.06.2025_LS.xlsx`), reference pipe class summary (`PIPE CLASS SUMMARY_LS_06.06.2025_updated_column_names.xlsx`).
- **Outputs:** Two Excel files in the `output` folder: `Line_List_with_Matches.xlsx` (validation results) and `Pipe_Class_Summary.xlsx` (summary of compliance).
- **How to Run:** Run `python line_list_check\line_list_check_pipe_class.py` from the command line.
- **Additional Notes:** There is a refactored version (`line_list_check_pipe_class_refactored.py`) with improved readability, but the main logic is unchanged.

Path: c:\Users\szil\Repos\excel_wizadry\line_list_check\line_list_check_pipe_class.py  
Keywords: pipe class, validation, Excel, compliance

## line_list_compare
### line_list_compare.py
Description:
- **Purpose:** Compares two pipeline list Excel files (old and new) and highlights differences such as changes, additions, and deletions.
- **Inputs:** `pipeline_list_old.xlsx` (old file), `pipeline_list_new.xlsx` (new file), both with a 'Query' sheet.
- **Outputs:** Excel file `pipeline_list_new_COMPARE_WITH_PREV_REV.xlsx` with color-coded differences.
- **How to Run:** Run `python line_list_compare.py` from the command line. By default, uses files in the folder.
- **Additional Notes:** Automatically detects the KKS header and handles errors like missing files or duplicate KKS values.

Path: c:\Users\szil\Repos\excel_wizadry\line_list_compare\line_list_compare.py  
Keywords: pipeline, comparison, Excel, KKS

## pipe_class_summary
### pipe_class_summary_script_clean.py
Description:
- **Purpose:** Summarizes and analyzes pipe class data from multiple Excel files, using OpenAI for advanced processing.
- **Inputs:** Excel files in the `pipe_class_summary` folder, especially `Summary.xlsx`.
- **Outputs:** Updated summary Excel file and debug logs.
- **How to Run:** Run `python pipe_class_summary_script_clean.py` from the command line. Requires OpenAI API key in `.env`.
- **Additional Notes:** The script `pipe_class_summary_script.py` is an older version with similar functionality but less optimized.

Path: c:\Users\szil\Repos\excel_wizadry\pipe_class_summary_script_clean.py  
Keywords: pipe class, summary, OpenAI, Excel