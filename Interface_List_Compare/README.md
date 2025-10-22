# Interface List Comparator

A Python tool for comparing Excel interface lists, generating detailed change analysis with visual formatting in-place.

## 🚀 Quick Start

```bash
# Run with default file (compare.xlsx in script directory)
python Interface_list_excel_compare.py

# Or specify custom input file
python Interface_list_excel_compare.py path/to/your_file.xlsx
```

**Default File**: If no argument is provided, the script automatically uses:  
`C:\Users\szil\Repos\excel_wizadry\Interface_List_Compare\compare.xlsx`

## 📁 Core Files

- **`Interface_list_excel_compare.py`** - Main comparison script
- **`compare.xlsx`** - Default input file with interface list sheets
- **`README.md`** - This documentation

## ✨ Features

- **Interface No. based comparison** - Uses "Interface No." column as unique identifier
- **Positional sheet selection** - Rightmost sheet is NEW, left neighbor is OLD
- **In-place results** - Writes to input file, no new file created
- **Date-based naming** - Creates sheet named `YYMMDD_compared` (e.g., `251022_compared`)
- **Visual change tracking** - Color-coded rows with change markers
- **Formatting preservation** - Copies all formatting from OLD sheet
- **Intelligent cleanup** - Removes legend rows and previous comparison artifacts
- **Duplicate handling** - 1-to-1 in-order matching with dark red font
- **Localization support** - Handles English/German column names

## 📊 Sheet Selection Logic

The script automatically selects sheets based on position:
- **NEW sheet**: Rightmost sheet in the workbook
- **OLD sheet**: Direct left neighbor of the NEW sheet

Example workbook with sheets `[Sheet1, Sheet2, Sheet3, Sheet4]`:
- NEW = `Sheet4` (rightmost)
- OLD = `Sheet3` (left of Sheet4)

After comparison, the result sheet becomes the new active sheet.

## 🎨 Color Coding

| Color | Meaning | RGB |
|-------|---------|-----|
| 🟢 **Green** | Added row (exists in NEW, not in OLD) | #00FF00 |
| 🟡 **Yellow** | Changed row (modified values) | #FFFF00 |
| 🔴 **Red** | Deleted row (exists in OLD, not in NEW) | #FFC0C0 |
| 🔴 **Dark Red Font** | Duplicate Interface No. | #FF0000 |

## � Output Sheet Structure

The generated `YYMMDD_compared` sheet contains:

1. **Headers** (Row 1)
   - Bold, wrap text enabled
   - Auto-height
   - Includes "Changed" column (16 width)

2. **Data Rows** (Starting Row 2)
   - All rows from NEW table (green background if added)
   - Changed cells highlighted in yellow
   - Changed column shows modification summary
   - Comments indicate old values
   - Wrap text enabled, auto-height

3. **Deleted Rows Section** (Bottom)
   - Separator after last NEW row
   - All deleted rows in red background
   - Only rows with valid Interface No. (empty rows filtered out)

4. **Formatting**
   - Freeze panes at C3 (2 rows, 2 columns)
   - Column widths preserved from OLD sheet
   - Special widths: PIPEN - Name (20), comment (30), Date (20), Changed (16)
   - Borders, fonts, alignment copied from OLD sheet

## 🔧 Intelligent Cleanup

Before comparison, the script automatically removes:

1. **Legend Rows** - Color coding explanation rows at top (if present)
2. **"Changed" Column** - From previous comparison runs
3. **"Unnamed" Columns** - Empty columns from previous runs
4. **Added/Deleted Sections** - Sections from previous comparisons (detected by empty Interface No.)

This ensures clean comparisons even when running on previous output files.

## 🔑 Unique Key: Interface No.

The comparison uses **"Interface No."** as the unique identifier. The script supports:
- English: "Interface No.", "interface no.", "INTERFACE NO."
- German: "Schnittstellen-Nr.", "schnittstellen-nr.", etc.

### Duplicate Handling

If duplicate Interface No. values exist:
- 1-to-1 in-order matching between OLD and NEW
- First occurrence matched first
- Duplicate rows get **dark red font** color (#FF0000)
- Each duplicate matched only once (no cartesian product)

Example:
```
OLD has: [A001, A001, A002]
NEW has: [A001, A001, A003]
Result: 
  - 1st A001 OLD ↔ 1st A001 NEW (matched)
  - 2nd A001 OLD ↔ 2nd A001 NEW (matched)
  - A002 OLD = deleted
  - A003 NEW = added
```

## 💻 Usage Examples

### Example 1: Default File
```bash
python Interface_list_excel_compare.py
```
Uses `C:\Users\szil\Repos\excel_wizadry\Interface_List_Compare\compare.xlsx`

### Example 2: Custom File
```bash
python Interface_list_excel_compare.py "C:\Projects\interface_lists\my_list.xlsx"
```

### Example 3: Current Directory File
```bash
python Interface_list_excel_compare.py interfaces.xlsx
```

## � Comparison Logic

1. **Load Data**
   - Read rightmost sheet (NEW) and left neighbor (OLD)
   - Remove legend rows if present
   - Remove previous comparison artifacts

2. **Create Unique Keys**
   - Use "Interface No." column
   - Detect duplicates using value_counts()

3. **Perform Comparison**
   - NEW table is the base
   - For each NEW row:
     - Find matching OLD row by Interface No.
     - If duplicate: match 1-to-1 in order
     - Compare all columns
     - Mark changes in yellow, add comments

4. **Track Additions/Deletions**
   - Added rows: In NEW but not in OLD (green)
   - Deleted rows: In OLD but not in NEW (red, at bottom)
   - Filter out empty deleted rows

5. **Write Output**
   - Create new sheet `YYMMDD_compared`
   - Write headers with bold, wrap text
   - Write data rows with colors and comments
   - Write deleted rows section
   - Apply formatting from OLD sheet
   - Set freeze panes at C3
   - Set as active sheet

## � File Handling

- **Input**: Excel .xlsx file with at least 2 sheets
- **Output**: Same file, new sheet added with date-based name
- **Backup**: Consider backing up before running (no automatic backup)
- **Active Sheet**: Result sheet becomes active after comparison

## 🛠️ Requirements

```
Python 3.8+
pandas
openpyxl
collections (standard library)
```

Install dependencies:
```bash
pip install pandas openpyxl
```

## 🐛 Troubleshooting

### "Interface No. column not found"
- Ensure your sheets have a column named "Interface No." or "Schnittstellen-Nr."
- Check for typos or extra spaces in column names

### "Not enough sheets in workbook"
- Workbook must have at least 2 sheets
- Rightmost sheet is NEW, left neighbor is OLD

### "No data to compare"
- Sheets may be empty or have only headers
- Check that data starts at row 2

### Duplicate Interface No. issues
- Duplicates are handled with 1-to-1 matching
- All duplicates shown with dark red font
- Check your source data for unintended duplicates

### Formatting issues
- Formatting is copied from OLD sheet
- If OLD sheet has no formatting, defaults are used
- Specific widths: PIPEN - Name (20), comment (30), Date (20), Changed (16)

### Legend rows not removed
- Legend detection looks for "Deleted", "Added", "Changed" keywords
- Manual cleanup may be needed for custom legend formats

## 📈 Performance

- **Small datasets (<100 rows)**: <1 second
- **Medium datasets (100-1000 rows)**: 1-3 seconds
- **Large datasets (1000+ rows)**: 3-10 seconds
- **Memory usage**: Efficient for typical interface list sizes

## 🔄 Workflow Integration

Typical workflow:
1. Export OLD interface list to Excel
2. Export NEW interface list to Excel
3. Copy both into sheets of `compare.xlsx` (or custom file)
4. Position NEW sheet rightmost, OLD sheet to its left
5. Run script
6. Review `YYMMDD_compared` sheet with color-coded changes
7. Can run again on same file (previous results are cleaned up automatically)

## ⚠️ Limitations

- Only compares 2 sheets at a time (rightmost and its left neighbor)
- Requires "Interface No." column in both sheets
- No automatic backup (save manually before running)
- Date format in sheet name is fixed (YYMMDD_compared)
- Duplicate matching is in-order only (not all permutations)

## 📝 Version History

- **v1.0**: Initial release with KKS comparison
- **v2.0**: Changed to Interface No. comparison
- **v3.0**: Added positional sheet selection
- **v4.0**: In-place results with date-based naming
- **v5.0**: Intelligent cleanup of previous comparisons
- **v6.0**: Improved duplicate handling with 1-to-1 matching
- **v7.0**: Current version with dark red duplicate font

---

**Note**: This tool is designed for interface list management in engineering projects, providing visual change tracking for configuration control and review processes.