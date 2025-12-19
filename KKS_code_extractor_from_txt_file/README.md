# KKS Code Extractor

## Overview
This script extracts KKS (Kraftwerk-Kennzeichen-System) codes from P&ID (Piping and Instrumentation Diagram) text files, handling various formatting patterns that may occur in such documents. The extracted codes are validated and exported to an Excel file.

## What are KKS Codes?
KKS codes are standardized identifiers used in power plants and industrial facilities to uniquely label equipment, systems, and components. The typical KKS code pattern handled by this script is:

```
[0-3][A-Z]{3}[0-9]{2}[A-Z]{2}[0-9]{3}
```
Example: `1HRK30BR210`

## Features
- **Handles multiple KKS code formats:**
  - Codes split across lines (e.g., `1 HTX68` on one line, `BR920` on the next)
  - Codes with spaces on a single line (e.g., `1 HTQ10 CL301`)
  - Codes without spaces (e.g., `1HTF14BZ010`)
- **Removes duplicates:** Uses a set to ensure only unique codes are exported.
- **Validates codes:** Ensures only codes matching the exact KKS pattern are included in the output.
- **Exports to Excel:** Outputs the list of valid KKS codes to an Excel file (no headers).

## Usage

### 1. Prerequisites
- Python 3.x
- Required packages:
  - `pandas`
  - `openpyxl`

Install dependencies with:
```
pip install pandas openpyxl
```

### 2. Input File
Prepare a text file containing the P&ID data from which you want to extract KKS codes. By default, the script expects the input file at:
```
C:\Users\szil\Repos\excel_wizadry\KKS_code_extractor_from_txt_file\softened_water_all_pdf_text.txt
```
You can change this path in the `main()` function.

### 3. Running the Script
Run the script from the command line:
```
python KKS_code_extractor_from_txt_file.py
```

### 4. Output
The script will generate an Excel file containing the extracted KKS codes at:
```
C:\Users\szil\Repos\excel_wizadry\KKS_code_extractor_from_txt_file\kks_codes_extracted.xlsx
```
You can change this output path in the `main()` function.

## How It Works
- **extract_kks_codes(file_path):**
  - Reads the input file line by line.
  - Uses regular expressions to find KKS code patterns, handling codes that may be split across lines or have spaces.
  - Collects all unique codes found.
- **validate_kks_code(code):**
  - Checks if a code matches the strict KKS pattern.
- **main():**
  - Sets input/output file paths.
  - Extracts and validates codes.
  - Saves the valid codes to an Excel file (no headers).

## Customization
- **Input/Output Paths:**
  - Edit the `input_file` and `output_file` variables in the `main()` function to use different files.
- **KKS Pattern:**
  - If your KKS codes use a different pattern, adjust the regular expressions in `extract_kks_codes()` and `validate_kks_code()`.

## Example
Suppose your input file contains:
```
1 HTX68
BR920
1 HTQ10 CL301
1HTF14BZ010
```
The script will extract:
- 1HTX68BR920
- 1HTQ10CL301
- 1HTF14BZ010

## License
This script is provided as-is, without warranty. Modify and use it as needed for your projects.

## Author
- [Your Name or Organization]

## Contact
For questions or suggestions, please open an issue or contact the maintainer.
