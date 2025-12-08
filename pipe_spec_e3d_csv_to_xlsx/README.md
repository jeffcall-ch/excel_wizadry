# E3D Pipe Specification CSV to XLSX Converter

## Description

This script converts E3D pipe specification CSV files to consolidated Excel (XLSX) format.

## Input File

The input is a CSV file created with **Reinhard's macro** from the **E3D WURE toolbar** using semicolon (`;`) as the separator.

### Input Format

The CSV file has a special structure:
- **First column**: `SPRE` (pipe specification reference)
- **Remaining columns**: Alternating column name and column value pairs separated by `;`

Example:
```
/EFDX/TU_Z-MGP0M0000CC_0012:F ;TYPE;TUBE;PBOR;12mm;SHOP;FALS; DTXR; MAPRESS Stainless steel tube...
```

## Output

The script creates a consolidated Excel (`.xlsx`) file with:
- Proper column headers
- All data organized in a tabular format
- Same filename as input with `.xlsx` extension

## Usage

```powershell
python convert_to_xlsx.py <input_file.txt> [output_file.xlsx]
```

### Examples

Using default output filename:
```powershell
python convert_to_xlsx.py TBY_all_pspecs_wure_macro_08.12.2025.txt
```

Specifying custom output filename:
```powershell
python convert_to_xlsx.py TBY_all_pspecs_wure_macro_08.12.2025.txt output.xlsx
```

## Requirements

- Python 3.x
- pandas
- openpyxl

Install dependencies:
```powershell
pip install pandas openpyxl
```

## Output Example

The script will parse all rows and create columns for all unique field names found in the CSV, such as:
- SPRE
- TYPE
- PBOR
- SHOP
- DTXR
- DTXS
- MTXX
- WEIGHT
- P1 CONN
- P2 CONN
- etc.
