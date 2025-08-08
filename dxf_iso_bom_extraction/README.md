# DXF Isometric BOM Extraction Script

This script extracts table data from piping isometric DXF files, specifically the "ERECTION MATERIALS" and "CUT PIPE LENGTH" tables, and outputs the results to CSV files for further analysis.

## Features
- Recursively traverses a directory and its subdirectories for `.dxf` files
- Extracts the "ERECTION MATERIALS" and "CUT PIPE LENGTH" tables from each DXF
- Preserves table rows and columns; each row in the CSV matches a row in the drawing
- Adds the Drawing-No. (from the drawing header) to each extracted row for traceability
- Aggregates all material rows into `all_materials.csv` and all cut length rows into `all_cut_lengths.csv`
- Generates a `summary.csv` listing for each DXF: file path, filename, Drawing-No., number of rows extracted for each table, missing tables, and any errors
- Skips missing tables or files with errors, but logs them in the summary
- Handles up to 20 columns and 100 rows per table
- Uses the `ezdxf` library for DXF parsing

## Requirements
- Python 3.8+
- Install dependencies:
```
pip install ezdxf
```

## Usage
Run the script from the command line:
```
python dxf_iso_bom_extraction.py <directory>
```
- `<directory>`: The root folder containing DXF files (recursively searched)

### Example
```
python dxf_iso_bom_extraction.py C:/Users/szil/Repos/excel_wizadry/dxf_iso_bom_extraction/drawings
```

## Output Files
- `all_materials.csv`: All rows from all "ERECTION MATERIALS" tables, with Drawing-No. included
- `all_cut_lengths.csv`: All rows from all "CUT PIPE LENGTH" tables, with Drawing-No. included
- `summary.csv`: For each DXF, lists file path, filename, Drawing-No., number of rows extracted, missing tables, and errors

## Notes
- If a table is missing or cannot be extracted, it is skipped and logged in the summary
- The script uses text entity positions to reconstruct tables; formatting may vary depending on drawing conventions
- For best results, ensure DXF files are text-based and tables are not embedded as images

## License
MIT License
