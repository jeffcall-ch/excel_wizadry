# DXF Isometric BOM Extraction Script

This script extracts table data from piping isometric DXF files, specifically the "ERECTION MATERIALS" and "CUT PIPE LENGTH" tables, and outputs the results to CSV files for further analysis.

## Features

### Core Functionality
- Recursively traverses a directory and its subdirectories for `.dxf` files
- Extracts the "ERECTION MATERIALS" and "CUT PIPE LENGTH" tables from each DXF
- Preserves table rows and columns; each row in the CSV matches a row in the drawing
- Adds the Drawing-No. (KKS code) to each extracted row for traceability
- Aggregates all material rows into `all_materials.csv` and all cut length rows into `all_cut_lengths.csv`
- Generates a `summary.csv` listing for each DXF: file path, filename, Drawing-No., number of rows extracted for each table, missing tables, and any errors

### Advanced Processing
- **Smart KKS Code Detection**: Automatically finds KKS codes (format: digit + 3 letters + 2 digits + "BR" + 3 digits) from drawing headers
- **Intelligent Column Validation**: For CUT PIPE LENGTH tables, validates and corrects column placement based on content type (piece numbers, numeric values, text remarks)
- **Category Header Processing**: Properly handles material category headers (PIPE, FITTINGS, VALVES, etc.) in ERECTION MATERIALS tables
- **Optimized Single-Row Format**: CUT PIPE LENGTH data is converted to one piece per row format for easier analysis

### Performance Optimizations
- **Pre-compiled Regex Patterns**: Faster text pattern matching
- **Efficient Entity Sorting**: Text entities sorted once for improved processing speed
- **Reduced Debug Output**: Configurable logging to minimize performance impact
- **Memory-Efficient Processing**: Optimized data structures and filtering algorithms

### Error Handling & Reliability
- Skips missing tables or files with errors, but logs them in the summary
- Robust error handling for malformed or incomplete table data
- Graceful handling of missing Drawing-No. information
- Comprehensive validation for piece numbers, cut lengths, and nominal sizes

## Requirements
- Python 3.8+
- Install dependencies:
```
pip install ezdxf
```

## Usage

### Command Line
Run the script from the command line:
```
python dxf_iso_bom_extraction.py <directory>
```
- `<directory>`: The root folder containing DXF files (recursively searched)

### Single File Testing
For testing a single DXF file:
```
python test_single_file.py
```

### Batch Processing with Performance Monitoring
For processing multiple files with timing information:
```
python test_dxf_iso_bom_extraction.py
```

### Example
```
python dxf_iso_bom_extraction.py C:/Users/szil/Repos/excel_wizadry/dxf_iso_bom_extraction/drawings
```

## Output Files

### all_materials.csv
All rows from all "ERECTION MATERIALS" tables with the following structure:
- PT NO: Part number
- COMPONENT DESCRIPTION: Component description
- N.S. (MM): Nominal size in millimeters
- QTY: Quantity
- WEIGHT: Weight value
- Drawing-No.: KKS code for traceability

### all_cut_lengths.csv (New Format)
All pieces from all "CUT PIPE LENGTH" tables in single-row format:
- PIECE NO: Piece identifier (e.g., `<1>`, `<2>`)
- CUT LENGTH: Length to cut in millimeters
- N.S. (MM): Nominal size in millimeters
- REMARKS: Special instructions (e.g., "PLD BEND")
- Drawing-No.: KKS code for traceability

**Note**: Each piece is now on its own row for easier data analysis and filtering.

### summary.csv
Processing summary for each DXF file:
- file_path: Full path to the DXF file
- filename: DXF filename
- drawing_no: Extracted KKS code
- mat_rows: Number of material rows extracted
- cut_rows: Number of cut length pieces extracted
- mat_missing: Boolean indicating if materials table was missing
- cut_missing: Boolean indicating if cut lengths table was missing
- error: Any error messages encountered

## Performance Information

The script includes comprehensive performance monitoring:
- **Per-file timing**: Shows processing time for each individual file
- **Progress tracking**: Displays "Processing file X/Y" during batch operations
- **Performance summary**: Shows total processing time and average time per file
- **Data summary**: Reports total rows extracted across all files

Typical performance: ~15 seconds per large DXF file with optimized processing.

## Validation Features

### CUT PIPE LENGTH Table Validation
The script automatically validates and corrects CUT PIPE LENGTH table data:
- **Piece Number Validation**: Ensures piece identifiers follow the `<n>` format
- **Numeric Value Validation**: Validates cut lengths and nominal sizes are numeric
- **Content-Based Column Correction**: Automatically places data in correct columns based on content type
- **Remark Detection**: Properly identifies and places text remarks like "PLD BEND"

### Data Integrity Checks
- Ensures each piece has minimum required data: piece number, cut length, and nominal size
- Validates KKS code format and placement
- Cross-references table titles and content for accuracy

## Technical Details

### Table Extraction Method
- Uses text entity X/Y coordinates to reconstruct table structure
- Handles merged header cells and complex table layouts
- Filters data based on spatial relationships to table titles
- Supports tables with varying column counts and layouts

### KKS Code Detection
- Pattern: `\d[A-Z]{3}\d{2}BR\d{3}` (e.g., "2QFB94BR140")
- Searches in bottom-right area of drawings relative to ERECTION MATERIALS table
- Fallback to manual Drawing-No. field detection if KKS pattern not found

## Troubleshooting

### Common Issues
1. **Missing Tables**: Check if table titles exactly match "ERECTION MATERIALS" and "CUT PIPE LENGTH"
2. **Incorrect KKS Codes**: Verify KKS codes follow the expected format and are positioned correctly
3. **Column Misalignment**: The script automatically corrects most alignment issues, but complex layouts may require manual review

### Debug Information
The script provides detailed debug output including:
- Table detection results
- Row extraction with coordinates
- Column validation and correction steps
- Performance timing for each processing stage

## Notes
- If a table is missing or cannot be extracted, it is skipped and logged in the summary
- The script uses text entity positions to reconstruct tables; formatting may vary depending on drawing conventions
- For best results, ensure DXF files are text-based and tables are not embedded as images
- CUT PIPE LENGTH data is automatically converted from 8-column (2 pieces per row) to 5-column (1 piece per row) format

## License
MIT License
