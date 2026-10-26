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
  and validation status.

Only the latest visible worksheet is processed. If several sheets are visible,
their names must contain unambiguous dates such as `25Sep2026`; otherwise use
`--sheet`. Hidden and historical sheets are retained unchanged.

The classification columns are placed immediately after Finish. Only Activity
ID, Activity Name, Start, Finish and the classification columns remain visible.
The saved filter displays `ALL = X`; dates use `DD.MM.YYYY`.
All data rows are sorted by Finish date, earliest first, then Activity Name
(case-insensitive) for the same day. Blank or unparseable Finish dates sort last.
Complete rows, including styles, classification marks and visibility, move
together. The header stays fixed. Reports retain both source and output row
references for ambiguous activities and manual differences.

The generated **PTS Status** worksheet lists all missing approved family and
milestone combinations at the top, including optional DTS milestones. Found
activities follow in family and procurement/engineering/erection rule order,
with dates, Activity IDs and links to the schedule rows. Optional missing items
appear here even though they are excluded from the report's required-only
`missing_milestones` list.

A **WEEK DELTA CALCULATOR** appears below the complete schedule, outside the
filter range. Enter the first date in the upper yellow cell and the second date
in the yellow cell directly below it, both in the Finish column. The green
result cell calculates `(second date - first date) / 7` with two decimal places;
reverse dates produce negative weeks. It stays blank until both dates are
entered. Date input validation is enabled, and Excel recalculates on opening.
The report records the three cell addresses under `calculator`.

Processing a generated workbook refreshes its status sheet and calculator
instead of duplicating them. User-entered calculator dates reset to blank.

Approved rows are coloured across the populated schedule width according to
their lot: pipe material green, engineering blue, hangers peach, valves yellow,
and Lot 5 erection lilac. The four non-valve colours reproduce the July manual
reference. Valve yellow is a new default. Previously applied matching palette
colours are cleared on unapproved rows within these families.

Matching is intentionally conservative. Case, spacing and punctuation are
normalized; one obvious typo in a word of at least four letters is tolerated.
Lot numbers, short milestone codes and DTS T/S remain distinct. Near matches
are reported instead of selected. Rules and colours are centralized in
`RULES` and `COLOURS` for future project amendments.

Schedules with merged cells, tables, conditional formatting or original
non-classification formulas stop for review rather than risk corrupting
references when columns move. These features have not been tested on other
projects. The supplied XLSX files and focused regression tests pass.