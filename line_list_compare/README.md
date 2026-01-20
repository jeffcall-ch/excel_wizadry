# Excel Line List Comparator

A high-performance Python tool for comparing Excel piping line lists using KKS-based identification. Generates detailed change analysis with visual formatting in a single comparison sheet.

## 🚀 Quick Start

```bash
# Run with default file (C:\Users\szil\Repos\excel_wizadry\Line_List_Compare\compare.xlsx)
python line_list_compare.py

# Or specify input file
python line_list_compare.py compare.xlsx

# Or with custom output (not recommended - auto-generated names are better)
python line_list_compare.py compare.xlsx my_result.xlsx
```

## 📁 Core Files

- **`line_list_compare.py`** - Complete comparison engine with CLI
- **`compare.xlsx`** - Sample input file with multiple sheets
- **`original_VBA.vba`** - Original VBA reference code
- **`README.md`** - This documentation

## ✨ Key Features

### Comparison Logic
- **KKS-based identification** - Uses `KKS` column as unique identifier for row matching
- **NEW data order** - Results maintain the row order from the NEW sheet
- **1-to-1 duplicate matching** - Handles duplicate KKS values intelligently (matches in order)
- **Three-way change detection** - Identifies Added, Changed, and Deleted cells/rows

### Input Handling
- **Multi-sheet support** - Automatically uses the two rightmost sheets
  - Rightmost sheet = NEW data
  - Second-to-right sheet = OLD data
- **Legend row detection** - Automatically skips color-coding legend rows
- **Status column preservation** - Carries over PIPEN status columns from OLD to NEW
- **Leading zero preservation** - KKS values like "0100" stay as text, not numbers

### Output Generation
- **In-place sheet creation** - Adds new timestamped sheet to input file (e.g., `251105_compared`)
- **Auto-generated sheet names** - Pattern: `YYMMDD_compared` (with `_1`, `_2` suffix if needed)
- **Single comparison sheet** - All results in one organized worksheet
- **Automatic sheet activation** - New comparison sheet is selected and active when file opens

## 📊 Output Format

The tool generates a single comparison sheet with the following structure:

### Row 1: Legend
- Color-coded labels: **Changed** (Yellow), **Added** (Green), **Deleted** (Red)

### Row 2: Column Headers
- All column names from the data tables
- Bold formatting preserved from OLD sheet

### Row 3+: Main Data (NEW base)
- **NEW data rows** in their original order
- **Color coding per cell:**
  - 🟢 **Green** = Cell value added (was empty in OLD)
  - 🟡 **Yellow** = Cell value changed (different in OLD)
  - 🔴 **Red** = Cell value deleted (was present in OLD, now empty)
  - ⚪ **No color** = No change (same in both OLD and NEW)
- **Dark red font** = Duplicate KKS values (warning indicator)
- **Cell comments** = Show OLD value for changed cells

### Bottom Sections
- **"Added Rows"** section - Rows with KKS that only exist in NEW (all green)
- **"Deleted Rows"** section - Rows with KKS that only exist in OLD (all red)

### Change Marker Column
- Rightmost column shows: "Changed", "Added", "Deleted" or combinations

## 🎯 Example Workflow

**Input file:** `compare.xlsx` with sheets:
- Sheet 1: `old_1.0` (OLD data, 853 rows)
- Sheet 2: `251105` (NEW data, 920 rows)

**Run:**
```bash
python line_list_compare.py
```

**Output:** New sheet `251105_compared` added to `compare.xlsx` showing:
- 920 rows in NEW order with change highlighting
- 67 completely new rows (in main table with all-green cells)
- Deleted rows section at bottom (rows only in OLD)
- Status columns preserved from OLD sheet

## 📋 Requirements

### Input File Requirements
- Excel file (`.xlsx` format)
- At least 2 sheets
- Both sheets must have a `KKS` column for identification
- Column structure should be similar between OLD and NEW sheets

### Status Columns (Optional)
If present in OLD sheet, these will be copied to NEW:
- `PIPEN - Name\n[Family, Given]`
- `PIPEN - Comment\n[text]`
- `PIPEN - Date\n[dd/mm/yyyy]`

### Python Environment
- Python 3.8+
- Dependencies:
  - `pandas` - Data manipulation
  - `openpyxl` - Excel file handling
  - Standard library: `argparse`, `logging`, `datetime`, `pathlib`

Install dependencies:
```bash
pip install pandas openpyxl
```

## 🔧 Technical Details

### Algorithm Overview
1. **Extract sheets** - Get rightmost 2 sheets from input file
2. **Skip legend rows** - Auto-detect and remove color-coding legends
3. **Load data** - Read as DataFrames with KKS as string (preserve leading zeros)
4. **Clean data** - Remove comparison artifacts, empty trailing rows
5. **Enrich NEW** - Copy status columns from OLD if missing in NEW
6. **Create unique keys** - Use KKS for matching, detect duplicates
7. **Compare** - Cell-by-cell comparison with NEW as base
8. **Format** - Apply colors, comments, fonts based on change type
9. **Write** - Add new sheet to input file with comparison results

### Change Detection Logic

| Scenario | OLD Value | NEW Value | Result | Color |
|----------|-----------|-----------|--------|-------|
| No change | "ABC" | "ABC" | No change | None |
| Changed | "ABC" | "XYZ" | Changed | Yellow |
| Added | (empty) | "ABC" | Added | Green |
| Deleted | "ABC" | (empty) | Deleted | Red |
| New row | N/A | Any | Added row | All green |
| Deleted row | Any | N/A | Deleted row | All red (bottom section) |

### Duplicate Handling
- If multiple rows have the same KKS value, they are matched 1-to-1 in order
- Example: KKS "100" appears 3 times in OLD and 2 times in NEW
  - First OLD[100] matches First NEW[100]
  - Second OLD[100] matches Second NEW[100]
  - Third OLD[100] goes to "Deleted Rows" section
- Duplicate KKS values are marked with dark red font color as a warning

### Performance Optimization
- **Vectorized operations** - Uses pandas for O(n) comparisons
- **Batch processing** - Processes changes in groups
- **Memory efficient** - Streams data without loading entire file in memory
- **~1,600 rows**: 8-10 seconds (includes I/O)

## 🌍 Localization

The script auto-detects system language:
- **English systems** - Labels: "Changed", "Added", "Deleted"
- **German systems** - Labels: "Geändert", "Hinzugefügt", "Gelöscht"

## 📈 Output Order

**CRITICAL:** Results are in **NEW data order**, not OLD data order.

This means:
- ✅ Row positions match the NEW sheet
- ✅ If NEW has: KKS [A, B, C], output shows [A, B, C]
- ✅ Even if OLD had [C, A, B], output still shows [A, B, C]
- ✅ Deleted rows (only in OLD) appear in separate section at bottom

**Why NEW order?**
- The output is meant to be a "marked-up NEW list" showing what changed
- Engineers review the NEW data with annotations, not OLD data
- Makes it easier to update systems with NEW data while seeing changes

## 🐛 Error Handling

The script handles:
- Missing KKS columns
- Empty rows at bottom of sheets
- Previous comparison artifacts (removes "Changed" columns)
- Legend rows from prior comparisons
- Column mismatches between sheets
- Missing status columns
- File I/O errors with cleanup

## 📝 Logging

All operations are logged to `comparison.log`:
- Data loading details
- Row counts and column info
- Enrichment statistics
- Change detection summary
- Performance timing
- Error traces

---

**Note**: This tool is optimized for piping line list comparisons in engineering projects. It replaces manual Excel/VBA comparison workflows with a fast, automated Python solution.