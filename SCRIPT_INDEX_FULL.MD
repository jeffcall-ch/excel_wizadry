## Table of Contents

| Script Group | Script | Purpose |
|-------------|--------|--------|
| comos_extract_text_to_number | [pipe_category_calculator.py](#pipe_category_calculatorpy) | Calculates the safety category for each pipe based on EN 13480-1 standards. |
| dxf_iso_bom_extraction | [async_dxf_extraction.py](#async_dxf_extractionpy) | Processes DXF files in a folder to extract bill of materials information using asynchronous operations for speed. |
| excel_search | [excel_file_search.py](#excel_file_searchpy) | Searches for a specific text string in all Excel and CSV files within a folder. |
| line_list_check | [line_list_check_pipe_class.py](#line_list_check_pipe_classpy) | Checks and validates pipe class specifications in a line list Excel file against reference standards for industrial piping. |
| line_list_compare | [line_list_compare.py](#line_list_comparepy) | Compares two pipeline list Excel files (old and new) and highlights differences such as changes, additions, and deletions. |
| pipe_class_summary | [pipe_class_summary_script_clean.py](#pipe_class_summary_script_cleanpy) | Summarizes and analyzes pipe class data from multiple Excel files, using OpenAI for advanced processing. |
| window_frame | [main.py](#mainpy) | Creates a colored, resizable frame around a selected window on your desktop. |
| window_frame_csharp | [MainWindow.xaml.cs](#mainwindowxamlcs) | Windows desktop app (C#) to create a colored frame around a selected window, similar to the Python version. |
| files_folders_scripts | [pdf_file_list_generator.py](#pdf_file_list_generatorpy) | Scans a folder for PDF files and generates a CSV listing each file with extracted KKS and system codes. |
| KKS_search_in_excel_files | [KKS_search_in_excel_files.py](#kks_search_in_excel_filespy) | Searches for KKS codes in all Excel files in a folder and exports matching rows to a CSV. |
| Navisworks_xml_filter_gen | [navisworks_xml_filter_gen.py](#navisworks_xml_filter_genpy) | Generates a Navisworks XML filter file from a list of KKS codes in an Excel file. |
| pdf_search | [multi_pdf_full_text_search.py](#multi_pdf_full_text_searchpy) | Searches for text across multiple PDF files in a folder and generates an Excel report of all matches. |
| python_package_documentation_crawler | [documentation_crawler.py](#documentation_crawlerpy) | Crawls official documentation for Python packages and compiles it into Markdown files. |
| Sikla_article_number_excel_compare | [sikla_article_number_excel_compare.py](#sikla_article_number_excel_comparepy) | Compares Sikla article numbers and quantities between two Excel sheets and highlights differences. |
| Sikla_price_sheet_clean_up | [sikla_price_sheet_clean_uppy](#sikla_price_sheet_clean_uppy) | Cleans and formats Sikla price sheet Excel files, removing unwanted rows and applying formatting. |

---


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


--------------------


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


--------------------


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


--------------------


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


--------------------


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


--------------------


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


--------------------


## window_frame
### main.py
Description:
- **Purpose:** Creates a colored, resizable frame around a selected window on your desktop. The frame follows the window as it moves, resizes, or changes state.
- **Inputs:** No input files. The user selects a window and a color via the GUI.
- **Outputs:** Visual frame overlay (no file output).
- **How to Run:** Run `python main.py` from the command line. The app will prompt you to select a window and color.
- **Additional Notes:** The script works only on Windows and requires `pywin32` and `tkinter`.

Path: c:\Users\szil\Repos\excel_wizadry\window_frame\main.py  
Keywords: window frame, GUI, overlay, Windows


--------------------


## window_frame_csharp
### MainWindow.xaml.cs
Description:
- **Purpose:** Windows desktop app (C#) to create a colored frame around a selected window, similar to the Python version.
- **Inputs:** User selects a window and color via the GUI.
- **Outputs:** Visual frame overlay (no file output).
- **How to Run:** Build and run the C# project in Visual Studio.
- **Additional Notes:** For Windows only.

Path: c:\Users\szil\Repos\excel_wizadry\window_frame_csharp\MainWindow.xaml.cs  
Keywords: window frame, C#, overlay, Windows


--------------------


## files_folders_scripts
### pdf_file_list_generator.py
Description:
- **Purpose:** Scans a folder for PDF files and generates a CSV listing each file with extracted KKS and system codes.
- **Inputs:** Source directory containing PDF files.
- **Outputs:** CSV file with file details and codes.
- **How to Run:** Run `python pdf_file_list_generator.py` after setting the source directory in the script.
- **Additional Notes:** Other scripts in this folder are helpers for counting and copying files based on codes.

Path: c:\Users\szil\Repos\excel_wizadry\files_folders_scripts\pdf_file_list_generator.py  
Keywords: PDF, file search, KKS, CSV


--------------------


## KKS_search_in_excel_files
### KKS_search_in_excel_files.py
Description:
- **Purpose:** Searches for KKS codes in all Excel files in a folder and exports matching rows to a CSV.
- **Inputs:** Search term and directory path.
- **Outputs:** CSV file with results.
- **How to Run:** Run `python KKS_search_in_excel_files.py "SEARCH_TERM" "DIRECTORY_PATH"` from the command line.
- **Additional Notes:** Automatically detects headers and tracks source file/sheet.

Path: c:\Users\szil\Repos\excel_wizadry\KKS_search_in_excel_files\KKS_search_in_excel_files.py  
Keywords: KKS, Excel, search, CSV


--------------------


## Navisworks_xml_filter_gen
### navisworks_xml_filter_gen.py
Description:
- **Purpose:** Generates a Navisworks XML filter file from a list of KKS codes in an Excel file.
- **Inputs:** Excel file (`Navis_KKS_List.xlsx`) with KKS codes.
- **Outputs:** XML filter file for Navisworks.
- **How to Run:** Run `python navisworks_xml_filter_gen.py` or use the provided batch/PowerShell scripts.
- **Additional Notes:** Output file is named with today's date.

Path: c:\Users\szil\Repos\excel_wizadry\Navisworks_xml_filter_gen\navisworks_xml_filter_gen.py  
Keywords: Navisworks, XML, KKS, filter


--------------------


## pdf_search
### multi_pdf_full_text_search.py
Description:
- **Purpose:** Searches for text across multiple PDF files in a folder and generates an Excel report of all matches.
- **Inputs:** Search text and directory path.
- **Outputs:** Excel report with matches, file paths, page numbers, and context.
- **How to Run:** Run `python multi_pdf_full_text_search.py "search text" path/to/pdf/directory` from the command line.
- **Additional Notes:** Requires `PyPDF2`, `pandas`, and `openpyxl`.

Path: c:\Users\szil\Repos\excel_wizadry\pdf_search\multi_pdf_full_text_search.py  
Keywords: PDF, search, reporting, Excel


--------------------


## python_package_documentation_crawler
### documentation_crawler.py
Description:
- **Purpose:** Crawls official documentation for Python packages and compiles it into Markdown files.
- **Inputs:** List of Python package dependencies.
- **Outputs:** Markdown files with documentation for each package.
- **How to Run:** Run `python documentation_crawler.py` after configuring the package list.
- **Additional Notes:** Requires `requests`, `beautifulsoup4`, and `markdownify`.

Path: c:\Users\szil\Repos\excel_wizadry\python_package_documentation_crawler\documentation_crawler.py  
Keywords: Python, documentation, crawler, Markdown


--------------------


## Sikla_article_number_excel_compare
### sikla_article_number_excel_compare.py
Description:
- **Purpose:** Compares Sikla article numbers and quantities between two Excel sheets and highlights differences.
- **Inputs:** Excel file with two sheets for comparison.
- **Outputs:** DataFrames with comparison results (can be exported).
- **How to Run:** Run `python sikla_article_number_excel_compare.py` from the command line.
- **Additional Notes:** Requires `pandas`.

Path: c:\Users\szil\Repos\excel_wizadry\Sikla_article_number_excel_compare\sikla_article_number_excel_compare.py  
Keywords: Sikla, article number, Excel, comparison


--------------------


## Sikla_price_sheet_clean_up
### sikla_price_sheet_clean_up.py
Description:
- **Purpose:** Cleans and formats Sikla price sheet Excel files, removing unwanted rows and applying formatting.
- **Inputs:** Excel file path provided as a command-line argument.
- **Outputs:** Cleaned Excel file with `_cleaned_up_YYYYMMDD_HHMMSS` appended to the filename.
- **How to Run:** Run `python sikla_price_sheet_clean_up.py "C:\\path\\to\\your\\file.xlsm"` from the command line.
- **Additional Notes:** Requires `pandas`.

Path: c:\Users\szil\Repos\excel_wizadry\Sikla_price_sheet_clean_up\sikla_price_sheet_clean_up.py  
Keywords: Sikla, price sheet, clean up, Excel


--------------------
