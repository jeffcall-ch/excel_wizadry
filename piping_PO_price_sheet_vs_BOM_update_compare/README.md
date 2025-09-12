PO vs BOM Comparison Script

Overview

This script compares a Purchase Order (PO) CSV file to a Bill of Materials (BOM) CSV file and produces a combined comparison report. It matches pipe components by the description/name, reports quantity and weight differences, marks matched/removed/new items, and writes a timestamped CSV report.

Requirements

- Python 3.8+ (script uses f-strings and pandas)
- pandas
- numpy

Install dependencies (recommended in a virtual environment):

```powershell
pip install -r requirements.txt
```

(Ensure `pandas` and `numpy` are available if not using the project's `requirements.txt`.)

Inputs

Place the following CSV files in the same folder as the script, or call the main function with explicit paths:

- `PO.csv` — expected columns include at least:
  - `DESCRIPTION/ Pipe Component` (string): the PO item description used to match BOM's `Pipe Component`
  - `PO QUANTITY\n (pcs./m)` (numeric or string): PO quantity
  - `PO WEIGHT\n [kg]` (numeric or string): PO weight
  - Additional PO columns are preserved in the output

- `BOM.csv` — expected columns include at least:
  - `Pipe Component` (string): the key used to match PO items
  - `Total weight [kg]` (numeric or empty): BOM total weight for the component
  - `QTY [pcs/m]` (numeric or empty): BOM quantity
  - `Category`, `Material`, `Type` (optional): used when adding new BOM items to the result

Behavior and outputs

- The script loads both CSVs into pandas DataFrames.
- It creates a copy of the PO DataFrame and augments it with new columns:
  - `BOM QTY [pcs/m]`
  - `BOM Total weight [kg]`
  - `QTY Difference [pcs/m]` (BOM - PO; shown only if non-zero)
  - `Weight Difference [kg]` (BOM - PO; shows `N/A (BOM weight missing)` if BOM weight missing)
  - `Status` with values: `Matched`, `Removed from BOM`, or `New in BOM`
- Matching is done by exact string equality between `DESCRIPTION/ Pipe Component` (PO) and `Pipe Component` (BOM). The script keeps only the first occurrence when duplicates exist in BOM and prints a warning about duplicates.
- For PO items not found in BOM, the script marks them as `Removed from BOM` and writes negative differences equal to the PO values.
- For BOM items not present in PO, the script appends new rows with `Status = New in BOM` and populates available columns (Category, MATERIAL, TYPE).
- The final comparison is saved as `PO_vs_BOM_Comparison_Report_<timestamp>.csv` in the script directory or an explicit output directory if provided.

Usage

Run the script directly from the command line (it expects `PO.csv` and `BOM.csv` in the same folder):

```powershell
python .\piping_PO_price_sheet_vs_BOM_update_compare.py
```

Or use the `compare_po_vs_bom` function from another Python module:

```python
from piping_PO_price_sheet_vs_BOM_update_compare import compare_po_vs_bom

output_path = compare_po_vs_bom(r"C:\path\to\PO.csv", r"C:\path\to\BOM.csv", output_dir=r"C:\path\to\out")
print(output_path)
```

Assumptions and notes

- Matching uses exact string equality. If your files contain minor differences (case, whitespace, punctuation), consider normalizing the keys before calling the function (e.g., lowercasing and stripping whitespace).
- The script keeps the first BOM occurrence on duplicate component names and ignores subsequent duplicates, printing a count of duplicates found.
- Numeric parsing uses pandas `to_numeric(..., errors='coerce')`, meaning malformed numbers are treated as NaN and then set to 0 for comparisons.
- Weight differences show a human-readable `N/A` if BOM weight is missing.
- The script preserves additional columns present in the PO file.

Edge cases

- If `PO.csv` or `BOM.csv` are missing, the script prints an error and exits.
- If BOM contains components not present in PO, they are appended as new rows.
- If BOM's `Total weight [kg]` is empty or non-numeric, weight differences for those BOM items are marked as `N/A (BOM weight missing)`.

Suggestions for improvement

- Add fuzzy/more robust matching (e.g., case-insensitive, stripped punctuation, or fuzzy matching) to reduce missed matches.
- Report all duplicates and optionally merge duplicate BOM lines rather than keeping only the first.
- Add unit tests for the core `compare_po_vs_bom` function.
- Allow configurable column names and include an argument schema for required columns.

License

No license included. Add one if you plan to share this repository publicly.
