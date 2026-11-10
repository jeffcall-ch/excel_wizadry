# Paste To Filtered Rows (Excel VBA)

This macro pastes copied Excel values only into the **visible cells** of the selected target range. It never removes or reapplies filters, including Excel Table filters.

## Files

- `excel_paste_to_filtered_rows.vba` – source code of the macro.
- `Test-PasteToFilteredRows.ps1` - integration tests using a separate Excel instance and synthetic workbooks.

## Install so it is available in **all Excel files**

Use the **Personal Macro Workbook** (`PERSONAL.XLSB`).

1. Open Excel.
2. If the **Developer** tab is hidden:
   - `File` → `Options` → `Customize Ribbon`
   - Enable `Developer`.
3. Open VBA editor: `Developer` → `Visual Basic` (or `Alt+F11`).
4. In Project Explorer, find `VBAProject (PERSONAL.XLSB)`.
   - If `PERSONAL.XLSB` does not exist yet:
     - `View` tab → `Macros` → `Record Macro...`
     - In **Store macro in**, choose **Personal Macro Workbook**.
     - Start and immediately stop recording.
     - Re-open `Alt+F11`.
5. In `PERSONAL.XLSB`, right-click `Modules` → `Insert` → `Module`.
6. Open `excel_paste_to_filtered_rows.vba`, copy all code, paste into the new module.
7. Save VBA project:
   - In VBA editor press `Ctrl+S`, or close Excel and choose **Save** when prompted for `PERSONAL.XLSB`.
8. Close and reopen Excel.

Now the macro is globally available in any workbook.

## Add button to the top Ribbon

1. `File` → `Options` → `Customize Ribbon`.
2. On the right side, create or select a custom group (for example under `Home` or `Data`).
3. On the left, set dropdown to `Macros`.
4. Select `PERSONAL.XLSB!PasteAndRestoreFilter` and click `Add >>`.
5. (Optional) Click `Rename...` to set a friendly label/icon.
6. Click `OK`.

You now have a Ribbon button at the top in every Excel session.

## Optional: assign a keyboard shortcut

1. `View` → `Macros` → `View Macros`.
2. Select `PERSONAL.XLSB!PasteAndRestoreFilter`.
3. Click `Options...` and assign a shortcut.

## Macro security notes

If macros are blocked:

- `File` → `Options` → `Trust Center` → `Trust Center Settings` → `Macro Settings`.
- Use your organization’s allowed setting (commonly **Disable with notification**).
- If needed, place files in a trusted location under `Trusted Locations`.

## Usage

1. Copy source cells (`Ctrl+C`).
2. Select either the starting target cell or the complete target range (can be filtered; visible cells only are used). A single selected cell expands automatically to match the copied rows and columns, skipping hidden target rows and columns.
3. Run `PasteAndRestoreFilter` from Ribbon.

For example, copy two visible cells in column H and select only G2: the values are written to G2 and the next visible row in column G. Hidden rows between them are untouched.

For an explicitly selected multi-cell target, copied count must match visible target count; otherwise the macro warns without writing. One copied cell and two explicitly selected visible target cells is still a mismatch; the macro does not repeat the source value.

## Behavior and limits

- Values only: formulas become their copied results. Numbers, Boolean values, Excel errors, blank cells and text remain their original value types. Text containing tabs, line breaks, leading zeros or a leading `=` is not parsed or evaluated.
- The target's number formats are retained; dates remain Excel serial values and display using the target format. Source formatting is not copied. Written cells are marked with red font, as in the original macro.
- Source values are read row by row, left to right. Visible target cells are sorted in the same row/column order, even for disjoint selections. A single-cell starting point preserves the copied row/column shape across visible cells. Explicit multi-cell selections may have a different shape when cell counts match.
- Automatic expansion from a cell inside a filter range or Excel Table stays within that range's rows and columns, excluding Table headers and totals. Insufficient visible space is rejected before writing. Outside a filter/Table range, expansion is limited by the worksheet edges.
- Filtered-out and manually hidden target rows and columns are skipped. The source follows Excel's native copy behavior: filtered-out source rows are excluded, but manually hidden source rows may still be copied.
- Filter criteria and current row visibility are left untouched. If values in a filtered column change, use Excel's **Reapply** command later if you want to recompute which rows match.
- Copy from Excel with `Ctrl+C`, in the same Excel application instance. Cut, external text/image clipboard data and separately launched Excel instances are not supported. Copy again after a failed attempt if Excel has left copy mode.
- Merged cells, legacy array formulas and dynamic spill ranges in the target are rejected before writing. Protected sheets are rejected, even when some cells are unlocked. Protected workbook structure is supported because no target sheet is added or deleted.
- At most 100,000 visible target cells are accepted. Select the actual data range rather than whole rows or columns.
- The macro snapshots target values, formulas, number formats and cell font colors before writing. On a write error it attempts to restore them and reports any rollback failure. This is not a durable transaction: Excel crashes, workbook side effects and per-character rich-text formatting are not covered.
- VBA changes do not provide normal Excel **Ctrl+Z** undo. Save a backup before using the macro on important data.
- An unsaved temporary workbook is used to read the native clipboard and is closed without saving. Screen updating, event handling and the target selection are restored after success or failure. Clipboard contents/copy mode are not guaranteed to survive.

When upgrading an installation, replace the **entire old module contents**, not just the first procedure. Existing buttons and shortcuts still use `PasteAndRestoreFilter`. `PasteVisibleValues` is the error-raising core used by the wrapper and tests; run the wrapper interactively.

## Reliability Review

The original implementation had these issues:

| Severity | Finding | Correction |
| --- | --- | --- |
| High | Splitting clipboard text on tabs/newlines confused cell contents with cell boundaries; quoted text was not decoded. | Native Excel values-only paste replaces text parsing. |
| High | Converting all values to text lost numeric/error types and could turn literal `=...` text into formulas. | Read `Value2`; temporarily use Text format when assigning strings, then restore the target format. |
| High | Traversing multi-area ranges in Excel's area order could scramble multi-column pasted data. | Explicit row/column ordering of all visible target cells. |
| High | `ShowAllData` removed filters before writing; write failures bypassed filter restoration. | No filter removal or reconstruction. |
| High | Filter capture/restore did not reliably cover Table filters, grouped dates, colors or icons, and restoration errors were silently ignored. | Existing filter objects and criteria are never modified. |
| High | Errors midway through writing left partially updated data and colors. | Preflight checks, complete snapshots before writing and best-effort rollback. |
| Medium | `UsedRange` did not reliably represent trailing blank copied cells. The fallback added/deleted a sheet in the user's workbook. | Read the actual range selected by native paste in a separate disposable workbook. |
| Medium | Empty/failed clipboard reads could bypass fallback; the legacy HTML clipboard object was an environment-dependent dependency. | One native Excel clipboard path with normal error reporting. |
| Medium | Unsafe targets and huge selections were not checked before writes. | Explicit protection, merged/array/spill, overlap and size checks. |

## Integration Tests

Requires desktop Excel for Windows and the Excel Trust Center option **Trust access to the VBA project object model** for the test harness. Normal macro use does not require that option. Follow your organization's policy; the test does not change security settings.

Run from the repository root in PowerShell:

```powershell
& .\Excel_paste_to_filtered_rows\Test-PasteToFilteredRows.ps1
```

The harness uses synthetic, unsaved workbooks in its own Excel instance and does not open or save user files. It overwrites the system clipboard. Expected failures are caught inside a test-only VBA wrapper so they do not produce runtime dialogs under the normal VBA error-trapping setting. **Break on All Errors** in the VBA editor will still interrupt tests; use **Break on Unhandled Errors** when running them.

Coverage includes text delimiters and formula-like text, native value types, trailing/all-blank cells, formula results, single-cell starting-point expansion (including the H2/H4 to G2 case), multi-column expansion, filter-boundary rejection, multi-area targets, hidden rows/columns, overlapping copy/paste, rejected selections, filtered sources, Table filters, cross-workbook copying, application settings and injected mid-write failures. Dynamic spill tests require an Excel version supporting `SEQUENCE`.
