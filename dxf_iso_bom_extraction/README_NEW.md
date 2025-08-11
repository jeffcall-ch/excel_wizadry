# DXF Isometric BOM Extraction Tool

This tool extracts "ERECTION MATERIALS" and "CUT PIPE LENGTH" tables from DXF piping isometric drawings and exports them to CSV files. It provides traceability by including drawing numbers and pipe classes, supports batch processing with automatic multiprocessing for performance.

## Features

- **Table Extraction**: Extracts both "ERECTION MATERIALS" and "CUT PIPE LENGTH" tables
- **Traceability**: Includes drawing number (from filename) and pipe class (from DESIGN DATA section)
- **Batch Processing**: Processes all DXF files in a directory recursively
- **Multiprocessing**: Automatically uses parallel processing for multiple files
- **Performance Monitoring**: Detailed timing and efficiency metrics
- **Error Handling**: Robust error handling with detailed error reporting
- **Debug Mode**: Optional verbose output for troubleshooting
- **CSV Output**: Clean, structured CSV files with proper headers

## Output Files

The tool generates three CSV files:

1. **all_materials.csv**: Combined erection materials from all drawings
2. **all_cut_lengths.csv**: Combined cut pipe lengths from all drawings  
3. **summary.csv**: Processing summary with file statistics and errors

## Usage

### Basic Usage
```bash
python dxf_iso_bom_extraction.py <directory>
```

### With Debug Output
```bash
python dxf_iso_bom_extraction.py <directory> --debug
```

### With Custom Worker Count
```bash
python dxf_iso_bom_extraction.py <directory> --workers 4
```

### Arguments

- `directory`: Root folder containing DXF files (searched recursively)
- `--debug`: Enable detailed debug output (optional)
- `--workers`: Number of parallel workers (optional, auto-detected by default)

## Performance

The tool automatically detects when to use multiprocessing:
- **Single file**: Sequential processing
- **Multiple files**: Parallel processing using multiple CPU cores

### Worker Count Selection
- **Auto-detect (default)**: Uses min(file_count, cpu_count)
- **Manual**: Specify exact worker count with `--workers`
- **Sequential**: Use `--workers 1` to force sequential processing

## Example

```bash
# Process all DXF files in project folder with debug output
python dxf_iso_bom_extraction.py "C:\Projects\Piping\Isometrics" --debug

# Use 6 parallel workers for large batches
python dxf_iso_bom_extraction.py "C:\Projects\Piping\Isometrics" --workers 6
```

## Output Examples

### Summary Report
```
=== PROCESSING COMPLETE ===
Total files processed: 15
Successful: 15
Failed: 0
Workers used: 4
Total wall time: 12.34 seconds
Total processing time: 45.67 seconds
Average per file: 3.04 seconds
Speedup factor: 3.70x
Parallel efficiency: 92.5%
Time saved: 33.33 seconds
Total material rows: 1,234
Total cut length rows: 567
Output files written to: C:\Projects\Piping\Isometrics
```

## Requirements

- Python 3.8+
- ezdxf library
- Windows/Linux/macOS
