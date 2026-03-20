# Price Sheet Generator — Solution Progress

## Approach

Python script (`create_price_sheet_template.py`) that generates an `.xlsx` workbook using `openpyxl`. The workbook uses M365 dynamic array formulas (`SORT`, `UNIQUE`, `FILTER`, `VSTACK`, `IFS`, `CEILING.MATH`) — no VBA, no Power Query. User pastes input data into designated sheets; output sheets auto-calculate.

## What Exists

### Script: `create_price_sheet_template.py`

~1080 lines. Generates `Price_Sheet_Template.xlsx` with 17 sheets.

**Input sheets (9):**
- `IN_SS`, `IN_CS`, `IN_MAPRESS` — main material (paste "Total Piping Material" data starting row 3)
- `IN_Erect_SS`, `IN_Erect_CS`, `IN_Erect_MAP` — erection material
- `IN_Paint` — painting BoM ("PIPING PAINTING MATERIAL" data)
- `IN_CS_KKS`, `IN_SS_KKS` — KKS detail data for flange guard extraction

**Config sheet (1):**
- `CFG` — blind disk DN→thickness lookup, flange guard system keywords, spare rules reference text

**Calculation sheet (1):**
- `CALC_Paint` — painting TUBI pivot (UNIQUE+SUMPRODUCT by description+colour), CS TUBI for-order aggregation, fitting surface by colour, TUBI surface by colour

**Output sheets (6):**
- `1. Mapress`, `2. Stainless Steel`, `3. Carbon Steel` — dynamic array data (SORT+FILTER+CHOOSE) + spare/for-order/weight formulas
- `4. Erection Material` — VSTACK of 3 erection inputs + erection spare formulas
- `5. Painting +6. Add Part` — painting pipe data, paint cans, blind disks, flange guards, welding nozzles
- `Price summary sheet` — grand totals from all sections

### Key Design Decisions

- `MAX_DATA_ROWS = 500` — formulas pre-filled for 500 data rows per sheet
- `GRAND_TOTAL_ROW = 510` — fixed row for grand totals (leaves gap between data and total)
- Dynamic array in columns A–G (spill formula in row 8), row-by-row formulas in columns H–L
- CS TUBI spare overridden by painting for-order calculation via CALC_Paint sheet
- All formulas use `_xlfn.` / `_xlfn._xlpm.` prefixes for OOXML compatibility

## Known Issues & Incomplete Items

### Bug: Painting column indices are wrong

The actual painting input file has **17 columns** (no "E3D System" column). The template was originally written assuming 18 columns. A partial fix was attempted (replacing `$R$5000` → `$Q$5000` and shifting column indices 17→16, 15→14, 14→13, 5→4, 4→3), but the replacement via PowerShell corrupted some references due to `$` variable expansion. **The column references need a clean re-verification and fix.**

Correct painting column mapping (1-based):
| Col | Header |
|-----|--------|
| 1 | Name of System |
| 2 | Name of Pipe |
| 3 | Type |
| 4 | Description |
| 5 | Medium |
| 6 | DN 1 |
| 7 | DN 2 |
| 8 | Material |
| 9 | Pipe Class |
| 10 | Building Section |
| 11 | AIC |
| 12 | Quantity [pcs] |
| 13 | Quantity [m] |
| 14 | External Surface [m2] |
| 15 | Corrosion class |
| 16 | Painting colour (acc. BS4800) |
| 17 | Insulated |

So formula range should be `$Q$5000` (17 cols), and column indices for Type=3, Description=4, Qty_m=13, Surface=14, Colour=16.

### Missing: Orifice plates

Section 6A in the existing output has orifice plate rows (5 items). Template has no placeholder for these. Should add manual-entry rows.

### Missing: Coversheet

The existing output has a project info/revision coversheet. Template doesn't generate one.

### Untested with real data

The template was generated successfully but never tested by pasting actual rev0 input data and comparing output against the existing rev0 price sheet.

### Revision comparison not implemented

The input files have "QTY prev. rev." and "Difference" columns which are not surfaced on the output sheets. No highlighting or change tracking between revisions.

### Potential formula fragility

- Flange guard MMULT-based system matching (comparing system names against CFG keyword list) — complex and untested
- Paint cans primer logic (doubling for "basic only" colour) — needs verification
- Price summary references hardcoded row numbers (e.g. `H224`, `F200`) that depend on section positioning in the painting sheet
- Erection VSTACK filter condition relies on Type column being non-empty

### Grand total placement

Fixed at row 510 regardless of data size. The existing output places it right after the last data row. This means ~400+ blank rows in between for most cases.

## What Was Verified

- All spare rules extracted from VBA files and encoded as Excel formula logic
- CS↔painting feedback loop numerically verified: 100% match for all 10 pipe sizes in rev0 data
- Main material and erection input column headers match the template's expected layout
- KKS input column headers match
- Script runs and generates `.xlsx` without errors

## Files in Working Directory

```
Price_Sheet_New_Format/
  create_price_sheet_template.py   — the generator script (needs painting column fix)
  Price_Sheet_Template.xlsx        — generated output (has wrong painting refs)
  REQUIREMENTS.md                  — clean requirements for fresh implementation
  SOLUTION_PROGRESS.md             — this file
  _audit_refs.py                   — helper script used to count IN_Paint column references
  _fix_paint_cols.py               — attempted fix script (not successfully applied)
  rev0/input/                      — revision 0 BoM input files
  rev0/output_current_format/      — existing rev0 price sheet for reference
  rev1/input/                      — revision 1 BoM input files
  rev1/output_current_format/      — existing rev1 price sheet for reference
```
