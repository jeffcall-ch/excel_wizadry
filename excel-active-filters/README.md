# Active Filters Popup

VBA macros for PERSONAL.XLSB and Excel 365 desktop on Windows. The popup lists
active and remembered filters on the active worksheet, including off-screen
columns in filtered ranges and Excel Tables. Table names distinguish duplicate
headers. There is no persistent panel, add-in package, Ribbon XML, or Windows
API dependency in the VBA.

## Install or re-import

1. If PERSONAL.XLSB does not exist, record a temporary macro with "Store macro
   in: Personal Macro Workbook", stop recording, and remove it if desired.
2. Press Alt+F11 and select `VBAProject (PERSONAL.XLSB)` in Project Explorer.
3. Before importing or re-importing, back up and remove old modules belonging
   to this tool: `ActiveFilters` and `AF_FilterTarget` (or its old
   `CFilterTarget` class). Do not remove unrelated modules, including an
   existing `ClearAllFilters` module.
4. Use File > Import File to import `src/AF_FilterTarget.cls` and
   `src/ActiveFilters.bas`; do not paste export headers into the editor.
5. Use Debug > Compile VBAProject, then save PERSONAL.XLSB in the VBE. Accept
   Excel's prompt to save the Personal Macro Workbook on exit if shown.

The public entry points are `ShowActiveFiltersMenu`, `ActiveFilters_OnClick`,
`ActiveFilters_SelfTest`, and `ActiveFilters_SelfTestReport`. Resolve any name
collisions with unrelated code before importing.

## Ribbon button and menu

In Excel, open File > Options > Customize Ribbon. Under "Choose commands from",
select Macros, then `PERSONAL.XLSB!ShowActiveFiltersMenu`. Add it to a custom
group, optionally rename it and choose an icon, and click OK.

Click the button on a worksheet to open the popup at the cursor:

- A checkmark means the filter is active; click it to clear that column.
- `(off)` means a cleared filter can be re-applied; click it to re-apply.
- `(cannot re-apply)` is a disabled entry for a cleared filter whose criteria
  could not be saved or whose type is unsupported.
- **Clear all filters on this sheet** appears for at least two active listed
  filters. It also attempts active filters in tables with hidden headers or
  disabled filter arrows.
- **Re-apply all cleared filters** appears for at least two re-applicable
  remembered entries on the sheet.
- **Forget cleared filters** appears when the menu contains remembered
  entries; it forgets the sheet's entries without re-applying them.

When no listed column filter is active but `FilterMode` is on, **Advanced
filter active - clear all** is available. Advanced Filter is cleared as a whole,
not remembered for re-application. With no entries or clear-all action, a
disabled status item appears. Escape or clicking outside dismisses the popup.

## Re-application and limits

The macro snapshots criteria before clearing. The following operators are
supported when Excel exposes the needed criteria:

- Single text or number conditions, and `xlAnd`/`xlOr` conditions: the
  recorded criteria and operator can be passed back to AutoFilter.
- Value lists/arrays (`xlFilterValues`): the exposed value criteria can be
  passed back to AutoFilter.
- Dynamic filters (`xlFilterDynamic`), including above/below average and date
  periods: the exposed criteria and operator can be passed back to AutoFilter.

Unsupported filters can still be cleared, but are shown as disabled afterward:

- Colour and icon filters: their operators are not supported for re-application.
- Top/bottom items and percent filters: their operators are not supported for
  re-application.
- Date-tree selections whose criteria Excel does not expose: there is no
  criterion to save and pass back. Other snapshot errors also prevent re-apply.

Memory is per sheet, with at most one remembered entry per sheet column. It is
in VBA memory and is lost when Excel closes, VBA is reset, or the module is
re-imported. Successful re-application consumes the entry; a failed attempt
keeps it for retry. Entries are discarded when their target becomes invalid or
the column is filtered again. **Clearing a filter by macro wipes Excel's undo
history**; there is no custom undo.

PivotTable filters and slicers are not handled. Hidden-header or filter-arrow-
disabled tables are omitted from the column list. Targets are checked against
their workbook, sheet, source, field, column and header; sheet ranges must also
retain their address. Table growth can change the range address. Same-named
replacement objects with identical metadata cannot be distinguished.

Protected sheets are not unprotected; errors are reported. Clear all is not
transactional, so a failure can leave earlier filters cleared. Scroll position
is restored on a best-effort basis; cells are not selected or activated. Chart
sheets and grouped sheets are rejected.

## Manual self-test

In the VBE, open the `ActiveFilters` module under `VBAProject (PERSONAL.XLSB)`
so PERSONAL.XLSB is the active project. Press Ctrl+G and run:

```vb
ActiveFilters_SelfTest
```

Alternatively use `Application.Run "'PERSONAL.XLSB'!ActiveFilters_SelfTest"`.
The test creates a temporary workbook, prints PASS/FAIL lines and a summary,
and closes it without saving. It also reports cleanup failures. Other installed
code may receive workbook-open/close events. Set VBE Tools > Options > General >
Error Trapping to **Break on Unhandled Errors**. If execution pauses, F5 can
continue to the cleanup handler; Reset prevents automatic cleanup.

Use a disposable workbook for manual checks of cursor placement, dismissal,
Advanced Filter, long/multiline headers, grouped/chart sheets and protection.

## Automated runner

From the repository root in Windows PowerShell:

```powershell
& .\excel-active-filters\tools\Run-SelfTest.ps1                 # Both (default)
& .\excel-active-filters\tools\Run-SelfTest.ps1 -Mode Unsaved
& .\excel-active-filters\tools\Run-SelfTest.ps1 -Mode Saved
```

Excel's Trust Center must allow **Trust access to the VBA project object
model** (File > Options > Trust Center > Trust Center Settings > Macro
Settings); organization policy may block it. The runner does not change this
setting. Set VBE Error Trapping to **Break on Unhandled Errors**. Its preflight
checks the VBE setting and fails for a nonzero `BreakOnAllErrors` value; if
the key or value is absent it warns that the setting could not be verified.

The runner starts and verifies a new hidden Excel process and imports the
source into a temporary host workbook. Unsaved mode uses an unsaved host;
Saved mode saves it as `%TEMP%\AF_SelfTest.xlsb`, refuses to overwrite an
existing file there, and removes its own file during cleanup. It does not
attach to an existing Excel instance or open PERSONAL.XLSB or existing work
files. A five-minute watchdog covers the Excel Run call and can stop only the
runner-owned process. After cleanup, a separate 10-second exit wait can also
stop that process. Two hangs with unchanged sources prevent another attempt.

Reports are printed and saved in `tools/selftest-last-unsaved.txt` and
`tools/selftest-last-saved.txt`; `tools/selftest-last.txt` contains the combined
output. The runner also writes `tools/selftest-trace.txt` for diagnosing a
hang. Success requires one zero-FAIL summary matching the number of PASS lines,
no FAIL lines, and successful cleanup; the check count is not hardcoded.
The existing unsaved report ends with:

```text
Summary: 101 PASS, 0 FAIL
```

The existing saved report ends with:

```text
Summary: 101 PASS, 0 FAIL
```

For a saved host, the observed popup `OnAction` is a quoted full path followed
by the macro name, for example
`'C:\Users\szil\AppData\Local\Temp\AF_SelfTest.xlsb'!ActiveFilters_OnClick`.
The unsaved report shows a quoted workbook name instead. The callback builder
uses the workbook name; Excel can resolve the saved binding to its full path.

## VBA source encoding

The exported `.bas` and `.cls` sources must be ASCII, without a BOM, with CRLF
line endings; the runner checks these before importing. When writing ASCII
text from PowerShell, normalize line endings and use .NET's ASCII encoding
instead of a cmdlet encoding that may add a BOM:

```powershell
$text = $text -replace "`r`n|`r|`n", "`r`n"
[IO.File]::WriteAllText($path, $text, [Text.Encoding]::ASCII)
```

Ensure `$text` contains only ASCII characters before writing; ASCII encoding
cannot preserve non-ASCII characters.
