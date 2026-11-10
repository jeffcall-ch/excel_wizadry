# PTS Highlighter

Uses pandas to classify schedule rows according to the approved rules in
[SKILL.md](SKILL.md). Uses openpyxl to preserve the workbook and apply Excel
formatting. Source workbooks are never overwritten.

## Run

From the repository root, using the existing virtual environment:

```powershell
.\.venv\Scripts\python.exe .\PTS_highlighter\pts_highlighter.py
```

Without file arguments, processes source XLSX/XLSM files in this directory,
excluding generated `_LS` workbooks and Excel lock files. For a specific file:

```powershell
.\.venv\Scripts\python.exe .\PTS_highlighter\pts_highlighter.py "C:\Schedules\schedule.xlsx"
```

Dependencies: `pandas` and `openpyxl`. These are available in the repository's
virtual environment. The system Python may not have them installed.

Options:

- `--inspect`: read-only workbook, manual-mark and colour inspection.
- `--sheet "Worksheet Name"`: explicitly select a visible schedule worksheet.
- `--overwrite`: replace existing generated outputs, never the source.

VS Code tasks: **PTS focused tests**, **PTS process schedules**, and
**PTS inspect reference workbooks**. The processing task replaces existing LS
outputs; command-line runs require `--overwrite` to do so.

## Output

- `<source stem>_LS.xlsx` (or `.xlsm` for an XLSM source).
- `<source stem>_LS.report.json`: counts, ambiguous activities, unresolved dates,
  missing required milestones, differences from existing manual classifications,
    near-match review activities with reasons and source/output rows, and validation status.

Only the latest visible worksheet is processed. If several sheets are visible,
their names must contain unambiguous dates such as `25Sep2026`; otherwise use
`--sheet`. Hidden and historical sheets are retained unchanged.

Immediately after Finish come **Proposed date LS** (blank, `DD.MM.YYYY`) and
**Remarks LS** (blank text cells with wrapping), followed by Piping, Hangers,
Valves and ALL. Existing LS columns are reused when processing an annotated
workbook; their entries move with the complete activity row. To retain LS entries,
process that annotated workbook rather than regenerating from the original source.
Only Activity ID, Activity Name, Start, Finish, the two LS columns and the
classification columns remain visible.
The saved filter displays `ALL = X`; dates use `DD.MM.YYYY`.
Rows with both Start and Finish blank (including whitespace-only cells) are
preserved but skipped for activity-ID validation and classification. This allows
repeated human-readable section headings without treating them as activities.
Skipped rows remain hidden by the activity filter and have no classification
lookup formulas; their count is recorded as `skipped_undated_rows` in the report.
Rows with either date populated are still processed, including invalid dates
which remain listed under `unresolved_dates`.
All data rows are sorted by Finish date, earliest first, then Activity Name
(case-insensitive) for the same day. Blank or unparseable Finish dates sort last.
Complete rows, including styles, classification marks and visibility, move
together. The header stays fixed. Reports retain both source and output row
references for ambiguous activities and manual differences.

The generated **PTS Status** worksheet lists all missing approved family and
milestone combinations at the top, including optional DTS milestones. Found
activities and near matches follow in a compact editable classification table,
with Activity IDs linking to the schedule rows. Optional missing items
appear here even though they are excluded from the report's required-only
`missing_milestones` list.

The classification table has just Activity Name, Status, Activity ID, Piping,
Hangers and Valves. Pending near matches come first, highlighted amber with
**NEEDS YOUR REVIEW** in the Status column. Entering X changes the live label to
`Manually approved`; clearing all three marks restores the warning. Unrelated
milestones, unrecognized families and family-only headings are omitted. Pending
near matches also appear under `excluded` in the JSON report. Validation checks
agreement with configured rules and saved marks, not completeness of the rules.

## Accept or Change Classifications

In **PTS Status**, enter `X` or leave blank in the **Piping**, **Hangers**
and **Valves** columns (D:F). Confirmed activities are prefilled; near matches
start blank. **ALL** calculates automatically on the main sheet only. Do not edit
the Activity IDs. Main-sheet marks use exact-match VLOOKUP formulas by Activity
ID, returning blank rather than zero for missing or unmarked activities.

Marks and family colours update in Excel without running Python. On the main
sheet, use **Data > Reapply** to refresh the `ALL = X` filter and reveal newly
accepted rows. Clearing marks reverses approval. Save the workbook to retain edits.
JSON counts reflect the last script run, not live edits. Review labels update
immediately. The status sheet has no freeze panes and uses content-fitted row heights.

Rerunning on the original source with `--overwrite` reads saved marks from its
existing `_LS` output before replacing it. Processing a generated workbook reads
its own classification table first. Both preserve added X marks and intentionally
cleared marks by Activity ID, without changing global matching rules. If a reviewed
workbook was renamed, pass that workbook as the input to retain its edits.
Duplicate Activity IDs on dated rows (case-insensitive) and invalid marks stop
processing.

A **WEEK DELTA CALCULATOR** appears below the complete schedule, outside the
filter range. Enter the first date in the upper yellow cell and the second date
in the yellow cell directly below it, both in the Finish column. The green
result cell calculates `(second date - first date) / 7` with two decimal places;
reverse dates produce negative weeks. It stays blank until both dates are
entered. Date input validation is enabled, and Excel recalculates on opening.
The report records the three cell addresses under `calculator`.

A **DATE SHIFT CALCULATOR** sits below it. Enter a starting date and a signed
number of calendar weeks, for example `+2`, `-3` or `0`. It adds `weeks * 7` days,
then moves a Saturday or Sunday result forward to the next Monday. Weekday
results remain unchanged; public holidays are not excluded. Fractional weeks
are rounded to the nearest whole day before the weekend adjustment. The result
accepts a copied real Excel date or typed/pasted text such as `03.07.2027` or
`03/07/2027`, interpreted as day-month-year. The starting-date cell is formatted
as text to prevent regional reinterpretation on direct entry. Impossible dates
show `Invalid date`; incomplete inputs stay blank. Results display as `DD.MM.YYYY`.
Its cell addresses and formula are under `calculator.date_shift` in the report.

Processing a generated workbook refreshes its status sheet and calculator
instead of duplicating them. Both calculators' inputs reset to blank.

Approved rows are conditionally coloured across the populated schedule width according to
their lot: pipe material green, engineering blue, hangers peach, valves yellow,
and Lot 5 erection lilac. The four non-valve colours reproduce the July manual
reference. Valve yellow is a new default. Previously applied matching palette
colours are cleared on unapproved rows within these families.

Matching is intentionally conservative. Case, spacing and punctuation are
normalized; one obvious typo in a word of at least four letters is tolerated.
Engineering `Purchase Order` receives both Piping and Hangers marks. The exact
Hangers variant `Purchase Order - Base on BOQ1` maps to `Purchase Order` through
a family-specific alias; this does not accept arbitrary suffixes or PO adjustments.
`Shipment` is intentionally excluded for Hangers, Pipe Material and Valves,
including any previously saved marks. `Ready for Shipment` and DTS remain approved.
Lot numbers, short milestone codes and DTS T/S remain distinct. Near matches
are reported instead of selected. Rules and colours are centralized in
`RULES` and `COLOURS` for future project amendments.

Schedules with merged cells, tables, user-created conditional formatting or original
non-classification formulas stop for review rather than risk corrupting
references when columns move. The script's own classification formulas and colour
rules are refreshed safely on reruns. Formula results are calculated by Excel;
openpyxl validates formulas and table marks but does not calculate formula caches.