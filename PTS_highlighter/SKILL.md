---
name: project-schedule-cleaner
description: Cleans uploaded XLSX project schedules for piping execution by keeping Activity ID, Activity Name, Start, and Finish visible; cleaning all Start and Finish dates; marking approved Piping, Hangers, and Valves milestones; filtering ALL to X; and returning a new workbook. Use when a user uploads a project schedule and asks to clean, simplify, filter, or extract relevant dates.
---

# Project Schedule Cleaner

## Goal
Create a practical piping execution view from an XLSX export. Preserve source data and remove planning flags from the selected Start and Finish dates.

## Inputs and priority
- Input: one uploaded `.xlsx` project schedule.
- Instructions given with the current workbook override the defaults only for that workbook.
- Do not change permanent defaults unless the user explicitly requests it.

## Procedure

### 1. Detect the schedule
- Inspect all worksheets and locate one header row containing the best combined matches for Activity ID, Activity Name, Start, and Finish.
- Do not assume the first sheet, first row, fixed column letters, or exact header spelling.
- Preserve every worksheet.

### 2. Detect required columns
Find these headers with tolerance for case, spaces, punctuation, harmless extra words, and small obvious typos:
- Activity ID
- Activity Name
- Start
- Finish

Prefer plain `Start` and `Finish`. Use `Current Start` or `Current Finish` only if the plain equivalent is absent and the meaning is clearly the current project date. Never select Baseline, BL Project, Forecast, Actual, Remaining Duration, Original Duration, or similar qualified columns unless the user explicitly requests them.

If any required column cannot be identified confidently, stop and report it as `not found`. Do not guess.

### 3. Clean all selected dates
Clean every populated schedule row in the detected Start and Finish columns before filtering. `All dates` means all values in these two columns, not dates in Baseline, Forecast, Actual, or other hidden columns.

For each value:
- Extract a valid day-month-year date embedded in the cell.
- Accept `.`, `/`, or `-` separators and interpret ambiguous numeric dates as day-month-year.
- Remove all other letters, asterisks, spaces, planning flags, and non-date characters around the valid date.
- Convert the result to a real Excel date and format it as `DD.MM.YYYY`.
- Examples: `22.09.2026 A`, `*22.09.2026`, and `22/09/2026*` all become `22.09.2026`; `26/10/2026` becomes `26.10.2026`.
- Keep blank cells blank.
- If no valid date can be extracted, preserve the original value and report the worksheet, cell, and original value. Never invent a date.

### 4. Set visible columns
Keep these original columns visible:
- Activity ID
- Activity Name
- Start
- Finish

Hide all other original columns, not delete them, unless the user asks to keep additional columns visible for the current workbook.

### 5. Add classification columns
Insert immediately to the right of Finish, shifting later columns right:
- Proposed date LS (initially blank, DD.MM.YYYY)
- Remarks LS (initially blank text cells, wrap text)
- Piping
- Hangers
- Valves
- ALL

If equivalent columns already exist, reuse and refresh them without creating duplicates.

### 6. Controlled activity matching
For every discipline, first match the approved activity family and then an approved milestone within that family.

Normalization may ignore capitalization, repeated spaces, punctuation, apostrophe differences, RFQ/RfQ, BOM/BoM, harmless extra words, and small obvious typos. Do not use unrestricted semantic matching. Do not classify a row from a family prefix or generic word alone.

- Confirmed match: place uppercase `X` in every applicable discipline column.
- No match: leave blank.
- Ambiguous: leave blank and report Activity ID and Activity Name.
- DTS (T) and DTS (S) are separate and optional. Match each independently. Do not report an absent DTS milestone as missing.

## Default activity rules

### Piping
Set `Piping = X` for confirmed matches to the following family and milestone combinations.

#### Family: Lot General Piping Pipe Material
- TSD / BANF RFQ
- Tender Due Date
- Technical & Commercial Comparison Pre-TCC
- TSD PO
- Technical & Commercial Comparison TCC
- Purchase Order
- Ready for Shipment
- DTS (T) [optional]
- DTS (S) [optional]

#### Family: Lot General Piping Engineering
- Purchase Order
- BOM / BOQ-2 - Issue for for PO General Piping Material
- BOM and ISO and arrangement Drawings for PO General Piping Erection

#### Family: Lot 5 - Erection General Piping
- BOQ 1 (Doc's for RfQ) - Issued from Engineering
- BOQ 2 (Doc's for PO) - Issued from Engineering
- Contractor On Site Mobilisation Completed

### Hangers
Set `Hangers = X` for confirmed matches to:
- Lot General Piping Engineering - Purchase Order
- Lot General Piping Engineering - BOM / BOQ-2 - Issue for for PO General Piping Material
- Lot General Piping Hangers - TSD / BANF RFQ
- Lot General Piping Engineering - BOM and ISO and arrangement Drawings for PO General Piping Erection
- Lot General Piping Hangers - Purchase Order
	- Exact approved alias: Purchase Order - Base on BOQ1. Do not generalize to arbitrary suffixes or PO adjustments.
- Lot General Piping Hangers - Ready for Shipment
- Lot General Piping Hangers - DTS (T) [optional]
- Lot General Piping Hangers - DTS (S) [optional]

### Valves
Set `Valves = X` for confirmed matches to:
- Lot General Piping Valves - TSD / BANF RFQ
- Lot General Piping Valves - Tender Due Date
- Lot General Piping Valves - Technical & Commercial Comparison Pre-TCC
- Lot General Piping Valves - TSD PO
- Lot General Piping Valves - Technical & Commercial Comparison TCC
- Lot General Piping Valves - Purchase Order
- Lot General Piping Valves - Ready for Shipment
- Lot General Piping Valves - DTS (T) [optional]
- Lot General Piping Valves - DTS (S) [optional]

### Shared classifications and guardrails
- The three listed Lot General Piping Engineering milestones receive X in both Piping and Hangers.
- `Lot 5 - Erection General Piping` is fixed. Do not generalize the lot number.
- The approved family prefix alone is insufficient.
- Do not mark unrelated activities from an approved family.
- If family and milestone are not both clear, leave blank and report the row as ambiguous.

### 7. Fill ALL
Set `ALL = X` if Piping, Hangers, or Valves contains X. Otherwise leave ALL blank.

Store confirmed and near-match activities in one editable status table keyed by
unique Activity ID. Prefill discipline marks for confirmed matches; leave review
marks blank. Let users enter X or clear discipline marks and calculate ALL with a
formula. Main schedule classification cells use exact-match VLOOKUP formulas by
Activity ID, returning blanks for missing or unmarked entries. Preserve saved
table marks before rebuilding generated views, including deliberately cleared
marks. Reject duplicate IDs and invalid marks rather than silently selecting one.
Use conditional formatting for family colours and automatic Excel calculation.
Users must reapply the schedule filter after changing marks to update visibility.

### 8. Final view
- Apply AutoFilter to the complete schedule range including the new columns.
- Save the filter with only `ALL = X` displayed.
- Freeze the row immediately below the detected header.
- Make the processed worksheet active.
- Preserve nonmatching rows; hide them only through the filter.

### 9. Save and validate
- Never overwrite the uploaded workbook.
- Save as `<original filename without extension>_Piping_Filtered.xlsx` unless the user requests another name.
- Reopen and validate the output.
- Verify that Proposed date LS, Remarks LS, Piping, Hangers, Valves, ALL appear exactly once in that order immediately after Finish. Reuse existing LS columns without losing their entries when processing an annotated workbook.
- Verify every discipline X produces ALL X, and ALL X never appears with all disciplines blank.
- Verify the filter covers the complete schedule and is saved with ALL = X selected.
- Verify the freeze pane is immediately below the header.
- Verify recognized dates are real Excel dates displayed as `DD.MM.YYYY`.
- Verify all worksheets, rows, and original columns remain and the source workbook is unchanged.

## Completion response
Return the created workbook and briefly report:
- processed worksheet and detected header row;
- Piping, Hangers, and Valves match counts;
- ambiguous matches;
- near matches to approved milestones within recognized families, with reasons and source/output row references; omit unrelated milestones, unrecognized families and family-only headings from Needs Review;
- unresolved date cells;
- nonoptional approved milestones not found;
- temporary overrides used.

Do not claim success unless the output workbook was created and validated.