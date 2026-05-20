# Price Sheet Sikla — BoM Aggregation Tool

Converts the Sikla `Total Qty` Excel sheet into a fully aggregated price-sheet-ready SQLite database and Excel workbook.

**Source file**: `CA100-KVI-50296373_0.0 - BoM General Hangers and support material.xlsm`

---

## Overview

The tool reads raw BoM rows, classifies each item into one of 10 material groups, and applies group-specific aggregation and spare calculation rules:

- Standard items are summed by article + coating with a percentage-based spare added.
- Threaded rods and beam sections are packed with a **First Fit Decreasing (FFD)** bin-packing algorithm to minimise bar waste.
- Glass fabric tape is aggregated by total cut length with a 50% spare factor.

A validation report (29 checks across 7 categories) is written to a dedicated sheet in the output xlsx every run.

See [CALCULATION_RULES.md](CALCULATION_RULES.md) for the full specification of all classification, aggregation, spare, and validation rules.

---

## Requirements

```
pip install -r requirements.txt
```

Key dependencies: `pandas`, `python-calamine`, `openpyxl`

Python 3.11+. Uses the project-level venv at `..\\.venv`.

---

## Usage

All commands share the same entry point:

```
python price_sheet_db_tool.py <command> [options]
```

### `test-run` — full pipeline in one shot (most common)

```
python price_sheet_db_tool.py test-run --excel "CA100-KVI-50296373_0.0 - BoM General Hangers and support material.xlsm"
```

Creates a timestamped `.sqlite` + `.xlsx` pair in `_test_runs\`.

If your source workbook does not use the default sheet name `Total Qty`, pass `--sheet` explicitly.

Example for the FGT workbook:

```
python price_sheet_db_tool.py test-run --excel "CA100-KVI-50296374_0.0 - BoM List Hangers and Supports Material FGT pipes.xlsm" --sheet "Total Qty FGT"
```

From this folder, run both source workbooks (PowerShell):

```
$py = "..\.venv\Scripts\python.exe"
& $py .\price_sheet_db_tool.py test-run --excel "CA100-KVI-50296373_0.0 - BoM General Hangers and support material.xlsm"
& $py .\price_sheet_db_tool.py test-run --excel "CA100-KVI-50296374_0.0 - BoM List Hangers and Supports Material FGT pipes.xlsm" --sheet "Total Qty FGT"
```

The xlsx contains:

| Sheet | Contents |
|-------|----------|
| `RAW_DATA_REV0` | All 1202 raw rows imported from the source Excel |
| `AGGREGATED_DATA_REV0` | 384 aggregated rows with all calculated columns |
| `PRICE_SHEET_REV0` | Supplier-facing summary with grouped totals and total price formulas |
| `VALIDATION` | 29 validation checks, colour-coded PASS / FAIL / WARN / INFO |

Optional: `--revision REV1` (default `REV0`) to tag the run.

---

### `import` — import raw data only

```
python price_sheet_db_tool.py import --excel "source.xlsm" [--revision REV0]
```

Reads the `Total Qty` sheet and writes it to `RAW_DATA_{revision}` in a new SQLite file alongside the Excel.

---

### `populate` — aggregate from an existing DB

```
python price_sheet_db_tool.py populate --db path\to\file.sqlite [--revision REV0] [--out output.xlsx]
```

Runs the aggregation + spare logic on `RAW_DATA_{revision}` and writes `AGGREGATED_DATA_{revision}`.

---

### `export` — export an existing DB to xlsx

```
python price_sheet_db_tool.py export --db path\to\file.sqlite [--revision REV0] [--out output.xlsx]
```

Writes the raw + aggregated tables and the validation sheet to Excel.

---

## Output Columns

### `AGGREGATED_DATA_REV0`

| Column | Description |
|--------|-------------|
| `Material_Type` | One of 10 material groups (01–10) |
| `Article_Number` | Sikla article number |
| `Description` | Item description (cleaned for bracket items) |
| `Qty` | Net quantity needed (no spare). For cut items: minimum whole bars/rolls. |
| `Total_Weight_kg` | Net weight of `Qty` items. For cut items: net cut weight only. |
| `Cut_Length_mm` | Cut length (NULL for aggregated cut items and standard items) |
| `Remarks` | Aggregated remarks; `"10m roll"` for tape |
| `Coating` | Coating specification (e.g. `HCP`, `C5L RAL7035`) |
| `Spare` | Extra units on top of `Qty`; NULL for items with no spare rule |
| `Order_Qty` | `Qty + Spare` — total to order |
| `Order_Weight_kg` | Weight of `Order_Qty` (scaled proportionally) |
| `Net_Cut_Length_m` | Total cut length needed in metres (FFD/tape items only) |
| `Net_Cut_Weight_kg` | Weight of net cut material (FFD/tape items only) |
| `Total_Order_Length_m` | Total bar/roll length ordered (FFD/tape items only) |
| `Utilisation_pct` | Bar/roll material utilisation % (FFD/tape items only) |
| `Spare_Calculation_Rule` | Human-readable explanation of how the spare was calculated |

---

## Material Groups

| Group | Label | Spare rule |
|-------|-------|------------|
| 01 | Primary Supports | 5% (non-rod); FFD + 5% buffer (threaded rods) |
| 02 | Brackets Consoles | 5% |
| 03 | Bolts Screws Nuts | 15% |
| 04 | Beam Sections | FFD + 5% buffer (beam section ms/tp); 5% (channel items) |
| 05 | Installation Material | 50% spare factor (glass fabric tape); 15% (end caps) |
| 06–10 | C5 variants | Same rules as 01–05 |

C5 remapping: if `"C5"` appears in the `Coating` value, the base group (01–05) is remapped to the corresponding C5 group (06–10).

---

## Project Structure

```
Price Sheet Sikla\
├── price_sheet_db_tool.py          # Main aggregation script
├── price_sheet_agg_compare.py      # Revision comparison script (see below)
├── price_sheet_agg_compare_test.py # Validation test suite for compare script
├── CALCULATION_RULES.md            # Full specification of all rules + validation checks
├── README.md                       # This file
├── _test_runs\                     # Auto-created; timestamped .sqlite + .xlsx outputs
│   ├── 20260520_110103_REV0.0.sqlite
│   ├── 20260520_110103_REV0.0.xlsx
│   ├── 20260520_110103_REV1.0.sqlite
│   ├── 20260520_110103_REV1.0.xlsx
│   └── 20260520_130715_AGG_COMPARE.xlsx
└── CA100-KVI-50296373_0.0 - BoM General Hangers and support material.xlsm   # Source
```

---

## Revision Comparison — `price_sheet_agg_compare.py`

Compares the `AGGREGATED_DATA` tables from two SQLite databases (e.g. REV0 vs REV1) and writes a colour-coded Excel workbook.

### Usage

```
python price_sheet_agg_compare.py --old _test_runs\REV0.sqlite --new _test_runs\REV1.sqlite
```

`--out` is optional; defaults to a timestamped `AGG_COMPARE.xlsx` in `_test_runs\`.

If `--old` / `--new` are omitted, the script uses the most recently created `.sqlite` files found in `_test_runs\` whose names contain `REV0` and `REV1` respectively.

### Output

The output sheet mirrors the `AGGREGATED_DATA` sheet layout from `price_sheet_db_tool.py`:

- Rows grouped by `Material_Type` with separator and per-group `SUBTOTAL` rows
- Grand total row
- Deleted rows section at the bottom, also grouped by `Material_Type`
- Two extra columns appended: **Status** and **Order_Qty_Change**

### Match key

`Article_Number + Coating`. When the same key appears multiple times, rows are matched positionally (first-to-first, second-to-second).

### Colour scheme (data rows only)

| Colour | Meaning |
|--------|---------|
| Green | New row, or a cell value was added (was absent in OLD) |
| Yellow | Value changed: increased, or non-numeric change |
| Light blue | Numeric value decreased (NEW < OLD) |
| Red | Cell value removed, or entire row deleted |

### Console output

```
Results : 4 new  |  96 changed  |  2 deleted
Output  : _test_runs\20260520_130715_AGG_COMPARE.xlsx
```

### Test suite

```
python price_sheet_agg_compare_test.py
```

Runs 51 checks across 15 scenarios (new row, deleted row, qty up/down, cell add/remove, key matching, column positions, etc.).  A clean run prints `ALL PASS — 51 passed 0 failed`.

---

## Validation

Every run produces a `VALIDATION` sheet with 29 automated checks. The console also prints a summary line:

```
Validation: 0 FAIL  3 WARN  22 PASS — see VALIDATION sheet
```

The 3 expected WARNs on the current source file are all source-data artefacts (one article with no weight in the source, two articles with unit-weight rounding). See the `Known Source Data Quirks` section in [CALCULATION_RULES.md](CALCULATION_RULES.md).

A run is considered clean if there are **0 FAIL** results.
