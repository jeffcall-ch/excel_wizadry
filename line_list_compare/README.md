# Pipeline List Comparison Tool

This tool compares two Excel pipeline lists (old and new versions) and generates a color-coded output file highlighting the differences between them.

## Features

- **Automatic KKS Header Detection**: Automatically locates the KKS column header in the input files regardless of its position
- **Color-Coded Differences**:
  - 🟡 **Yellow**: Changed values
  - 🟦 **Blue**: New rows
  - 🟥 **Red**: Deleted rows (placed at the end of the output file)
- **Smart Cell Comparison**: Identifies changes, additions, and deletions at the cell level
- **Professional Excel Formatting**:
  - Split and freeze panes at C2 for easy navigation
  - Auto-adjusted column widths
- **Robust Error Handling**:
  - Detects and reports duplicate KKS values
  - Handles missing files or invalid formats
  - Reports when KKS header is not found
- **Always Uses 'Query' Sheet**: Consistently reads from the 'Query' sheet in input files

## Usage

### Basic Usage

```python
from line_list_compare import compare_pipeline_lists

compare_pipeline_lists(
    old_file="path/to/old_pipeline_list.xlsx",
    new_file="path/to/new_pipeline_list.xlsx",
    output_file="path/to/output.xlsx"
)
```

### Command Line

```powershell
python line_list_compare.py
```

By default, the script looks for:
- `pipeline_list_old.xlsx` (old file)
- `pipeline_list_new.xlsx` (new file)
- Output: `pipeline_list_new_COMPARE_WITH_PREV_REV.xlsx`

## Requirements

- Python 3.6+
- pandas
- openpyxl

## Installation

```powershell
pip install pandas openpyxl
```

## Error Handling

The script performs the following validations:

1. **File Existence**: Checks if input files exist
2. **KKS Column**: Verifies that a KKS column exists in both files
3. **Uniqueness Check**: Ensures KKS values are unique in both files
4. **Sheet Validation**: Confirms the 'Query' sheet exists

If critical errors are encountered (like duplicate KKS values or missing KKS headers), the script will display an error message and exit with code 1.

## Testing

The script includes comprehensive test coverage (>95%) with pytest:

```powershell
python -m pytest _tests\ -v --cov=.
```

## Output Example

The generated Excel file will contain:
- All rows from the new file (with any changes highlighted)
- Any deleted rows from the old file (appended at the end)
- Headers at the top with frozen panes at C2
- Auto-adjusted column widths for better readability

## Use Case

This tool is particularly useful for:
- Tracking changes between pipeline revisions
- Quality control and verification of modifications
- Documenting updates for reporting purposes
- Visualizing additions and deletions in complex datasets
