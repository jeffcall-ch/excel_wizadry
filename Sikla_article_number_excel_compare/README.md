# Sikla Article Number Excel Compare

This script compares two sheets from an Excel file containing Sikla article numbers and purchase order quantities. It identifies differences in article quantities between a baseline and an updated sheet, and outputs a processed Excel file highlighting changes.

## Features

- Reads both sheets from `price_list_rev1_vs_sikla_bom_rev2.xlsx` in the script's folder.
- Strips whitespace from all cell entries for accurate comparison.
- Aggregates quantities by article number, preserving the first non-null description.
- Compares baseline and update sheets by "ARTICLE NUMBER" and "PO QUANTITY".
- Outputs a new Excel file listing differences:
  - Added, removed, or changed articles.
  - "REMOVED" rows are colored red.
  - Columns are auto-sized and center-aligned.
  - Top row is filtered; sheet is frozen at B2.

## How to Run

1. Place `price_list_rev1_vs_sikla_bom_rev2.xlsx` in the same directory as the script.
2. Ensure you have Python 3 and the required packages installed:
   - `pandas`
   - `numpy`
   - `xlsxwriter`
3. Run the script from the command line:
   ```
   python sikla_article_number_excel_compare.py
   ```
4. The output file will be saved in the same directory, named like:
   ```
   price_list_rev1_vs_sikla_bom_rev2_comparison_processed_YYYY-MM-DD_HH-MM-SS.xlsx
   ```

## Output

- The output Excel file contains a "Differences" sheet listing articles with changed, added, or removed quantities.
- Removed articles are highlighted in red for easy identification.
