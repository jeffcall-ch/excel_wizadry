# PDF Full Text Search Tool

## Overview

This tool allows you to search for specific text across multiple PDF documents within a directory and its subdirectories. It generates an Excel report listing all occurrences of the search term, including file paths, page numbers, and surrounding text context. The report includes direct hyperlinks to the PDF files for easy access.

## Features

- **Recursive PDF Search**: Scans all PDF files in the specified directory and all subdirectories
- **Context Capture**: Extracts text surrounding each match to provide context
- **Page Identification**: Records the exact page number where each match is found
- **Excel Report**: Generates a structured Excel report with the following:
  - File paths with hyperlinks to the original documents
  - Page numbers for quick navigation
  - Text context showing where the match appears
  - Summary statistics on search results
- **Case Sensitivity**: Optional case-sensitive search

## Requirements

The script requires the following Python packages:
```
PyPDF2>=3.0.0
pandas>=2.0.0
openpyxl>=3.1.0
```

Install these requirements using:
```powershell
pip install -r requirements.txt
```

## Usage

### Basic Command

```powershell
python multi_pdf_full_text_search.py "search text" path/to/pdf/directory
```

### Full Syntax

```powershell
python multi_pdf_full_text_search.py "search text" path/to/pdf/directory [--output path/to/output.xlsx] [--case-sensitive]
```

RUN LIKE THIS with .venv activated
C:/Users/szil/Repos/excel_wizadry/.venv/Scripts/python.exe pdf_search/multi_pdf_full_text_search.py "water pipe" "C:\Users\szil\MED contract" --output "C:\Users\szil\Desktop\search_results.xlsx"


### Parameters

- `search_text`: The text you want to find in the PDF documents (required)
- `directory`: Path to the directory containing PDF files to search (required)
- `--output` or `-o`: Custom output path for the Excel report (optional, default: pdf_search_results_[timestamp].xlsx)
- `--case-sensitive` or `-c`: Enable case-sensitive search (optional, default: case-insensitive)

## Examples

### Basic Search

## Known Limitations

- Only supports PDF files that are not encrypted or password-protected.
- Large directories or files may impact performance.
- Text extraction may fail for scanned or image-based PDFs.
Search for "procurement lot order" in all PDFs in the "ProjectDocuments" folder:

```powershell
python multi_pdf_full_text_search.py "procurement lot order" "C:\ProjectDocuments"
```

### With Custom Output Path

```powershell
python multi_pdf_full_text_search.py "procurement lot order" "C:\ProjectDocuments" --output "C:\Reports\search_results.xlsx"
```

### Case-Sensitive Search

```powershell
python multi_pdf_full_text_search.py "API Standard" "C:\Specifications" --case-sensitive
```

### Paths with Spaces

For paths containing spaces, make sure to enclose them in quotes:

```powershell
python multi_pdf_full_text_search.py "procurement lot order" "C:\Users\szil\MED contract\test"
```

## Output File Structure

The generated Excel file contains:

1. **Search Results** sheet:
   - File Path: Absolute path to the PDF file (with hyperlink)
   - Page Number: Page where the match was found
   - Match Context: Text surrounding the match (~50 characters before and after)

2. **Summary** sheet:
   - Total files searched
   - Files with matches
   - Total matches found
   - Search parameters used

## Notes

- Large PDFs may take some time to process
- Password-protected PDFs cannot be searched
- PDF files with scanned images (without OCR) will not return text matches
