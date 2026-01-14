# Spare Calculation Macros

Two VBA modules power the spare calculations used in this workbook:

- `CalculateFastenerSpares` lives in [Spare_calcs_for_price_sheet/spare_calcs_erection_material.VBA](Spare_calcs_for_price_sheet/spare_calcs_erection_material.VBA) and targets erection-material fasteners (nuts, bolts, washers, gaskets, and everything else).
- `CalculateSpares` lives in [Spare_calcs_for_price_sheet/spare_calcs_main_material.VBA](Spare_calcs_for_price_sheet/spare_calcs_main_material.VBA) and covers main-material assemblies where both piece counts and meter lengths exist.

Both macros assume the active worksheet already holds a filtered BOM-like table and will write the computed spares plus rule documentation directly onto that sheet.

## Worksheet Expectations

| Macro | Header row | Required columns |
| --- | --- | --- |
| `CalculateFastenerSpares` | Row 6 | C: Description, E: Quantity, F: Spare (created/overwritten) |
| `CalculateSpares` | Row 7 | C: Type (numeric portion parsed), D: Description, E: Qty pcs, F: Qty m, H: Spare (created/overwritten) |

General rules:

- Data must start immediately below the header row listed above.
- Quantities should be numeric; invalid entries are treated as zero.
- The macros clear existing Spare values and cell comments before writing new results.
- Results include both the calculated spare amount and a cell comment describing the rule that fired.

## Fastener Spare Logic (`CalculateFastenerSpares`)

Classification priority is Nut/Washer/Bolt (NWB) > Gasket > Standard. Keyword matching is case-insensitive (`nut, washer, bolt` for NWB; `gasket` for gaskets).

| Category | Quantity range | Spare logic |
| --- | --- | --- |
| NWB | `qty <= 30` | 100% of qty, rounded up to the next 10 |
| NWB | `30 < qty <= 100` | 50%, rounded up to the next 10 |
| NWB | `qty > 100` | 20%, rounded up to the next 10 |
| Gasket | `qty <= 30` | 100%, rounded up to the next integer |
| Gasket | `30 < qty <= 500` | 50%, rounded up to the next 10 |
| Gasket | `qty > 500` | 30%, rounded up to the next 10 |
| Standard | `qty < 400` | 35%, rounded up to the next integer |
| Standard | `qty >= 400` | 20%, rounded up to the next integer |

Additional behavior:

- Negative quantities are coerced to zero, which effectively skips the row.
- A timer plus calculation/screen updating toggles are used to speed processing and are restored on exit.
- At the end, the macro appends a human-readable rules section five rows beneath the last data row.
- A summary message box reports counts of each classification, processed rows, and skipped rows.

## Main-Material Spare Logic (`CalculateSpares`)

Classification priority is Seal > Critical > MAPRESS > Standard. Keywords are case-insensitive:

- Seal: `seal ring, seal, ring`
- Critical: `nipple, nipples, union, connection, connections, welding, adaptor`
- MAPRESS: description must contain both `mapress` and one of `cap, coupling, elbow, flange, tee`

Piece counts (column E) take precedence over meter lengths (column F); only if `Qty pcs` is zero will the macro evaluate `Qty m`.

### Piece-based rules

| Category | Range | Spare logic |
| --- | --- | --- |
| Seal | `qty <= 10` | Fixed minimum of 5 |
| Seal | `qty > 10` | Max(3, ceil(qty * 10%)) |
| Critical | `qty <= 10` | Max(3, ceil(qty * 50%)) |
| Critical | `qty > 10` | Max(3, ceil(qty * 10%)) |
| MAPRESS | `qty <= 5` | 0 |
| MAPRESS | `5 < qty <= 20` | 1 |
| MAPRESS | `qty > 20` | ceil(qty * 10%) |
| Standard | `qty <= 5` | 0 |
| Standard | `5 < qty <= 20` | 1 |
| Standard | `qty > 20` and `Type < 50` | ceil(qty * 10%) |
| Standard | `qty > 20` and `Type >= 50` | ceil(qty * 5%) |

### Meter-based rules

If `Qty pcs` is zero but `Qty m` is positive, the macro applies:

| Category | Spare logic |
| --- | --- |
| Seal | Same thresholds as piece mode (>=10 switches to max(3, ceil(qty * 10%))) |
| Critical | Same thresholds as piece mode (>=10 switches to max(3, ceil(qty * 10%))) |
| MAPRESS | `qty > 20` -> ceil(qty * 10%), else 0 |
| Standard | `Type < 50` -> ceil(qty * 15%), `Type >= 50` -> ceil(qty * 10%) |

Other behavior mirrors the fastener macro: screen/calculation settings are restored, comments describe the firing rule, documentation is appended below the table, and a summary MsgBox reports processed/typed counts.

## Running the Macros

1. Open the workbook in Excel and enable macros.
2. Navigate to the sheet containing the BOM table for the corresponding macro (fastener vs. main material) and ensure the layout matches the column expectations.
3. Press `Alt + F8`, choose either `CalculateFastenerSpares` or `CalculateSpares`, and click **Run**.
4. When execution finishes, review the Spare column for calculated values, inspect cell comments for the triggered rule, and scroll below the data for the auto-generated rule summary.

## Customizing the Rules

- All adjustable parameters (header row, column letters, thresholds, percentages, keyword lists) are centralized at the top of each module.
- Edit the constants to adapt to new spreadsheet layouts or spare policies. For example, change `STANDARD_THRESHOLD` or `TYPE_THRESHOLD` to shift when percentages change.
- After modifying code, re-run the macro so the new constants flow through; the documentation block will automatically update.

## Troubleshooting

- If the macro reports "No data found", confirm there is at least one non-empty row below the header in the required columns.
- A runtime error dialog will restore Excel settings and display the VBA error number/message; fix the underlying data or configuration and rerun.
- Should spare comments clutter the sheet, simply re-run the macro to regenerate them after adjusting inputs.
