# Active Filters Popup

VBA macros for PERSONAL.XLSB, Excel 365 desktop on Windows (32-bit or 64-bit).
Open a popup at the mouse cursor and click a filtered column to clear only that
column. Normal filtered ranges and Excel Tables on the active sheet are listed,
including off-screen columns. Table names distinguish duplicate headers.

This replaces the earlier floating-window proposal. There is no persistent
panel, add-in package, Ribbon XML, or Windows API dependency.

## Install

1. If PERSONAL.XLSB does not exist, record a temporary macro with "Store macro
   in: Personal Macro Workbook", stop recording, and remove that temporary macro
   if desired. Do not change your organization's macro-security policy.
2. Press Alt+F11. Select `VBAProject (PERSONAL.XLSB)` in Project Explorer.
3. Before importing new ones, back up and remove old modules belonging to this
  tool: `ActiveFilters` and its old `CFilterTarget` class (or a previously
  imported `AF_FilterTarget` class). Choose Remove in Project Explorer and
  export a backup when prompted. Do not remove unrelated modules, including
  your existing `ClearAllFilters` module or similarly named classes you own.
4. Use File > Import File to import both `src/AF_FilterTarget.cls` and
   `src/ActiveFilters.bas`. Import the files rather than pasting their export
   headers into the code editor. The files are ASCII with CRLF line endings.
5. Use Debug > Compile VBAProject. Excel normally already references the
   Microsoft Office Object Library and Microsoft Excel Object Library.
6. Save PERSONAL.XLSB in the VBE. When Excel closes, accept its prompt to save
   the Personal Macro Workbook if shown.

Only `ShowActiveFiltersMenu`, `ActiveFilters_OnClick`, `ActiveFilters_SelfTest`,
and `ActiveFilters_SelfTestReport` are Public. All internal procedures are Private and
use the `AF_` prefix, including `AF_ClearAllFilters`, `AF_ClearFilterTarget`,
and `AF_GetActiveFilterTargets`. The class is named `AF_FilterTarget`.
The old public `ClearAllFilters` routine is no longer defined by this tool.
If unrelated code uses the same module, class, or one of the four public
macro names, resolve that collision before importing; do not overwrite it.

## Self-Test

First open the `ActiveFilters` module under `VBAProject (PERSONAL.XLSB)` so
PERSONAL.XLSB is the active VBE project. Then press Ctrl+G and run:

```vb
ActiveFilters_SelfTest
```

If a different project's module is active, an unqualified Immediate-window
call can report "Sub or Function not defined". Open the correct module first,
or use the workbook-qualified call:

```vb
Application.Run "'PERSONAL.XLSB'!ActiveFilters_SelfTest"
```

The macro creates its own temporary workbook, prints individual `PASS:` or
`FAIL:` lines, and closes it without saving. It does not write to existing
workbooks. Expected final output:

```text
Summary: 26 PASS, 0 FAIL
```

Checks cover five sorted targets (C, T, X, AB/TblA, AF/TblB), single-field
clearing, retained filters and arrows, unchanged scroll position, stale-header
rejection, and an additional range starting at column D. The first sheet uses
a value filter, a multi-value array filter, and a cell-colour filter. Error
messages are suppressed during the self-test; unexpected errors print FAIL.

Eleven additional checks build the popup without calling ShowPopup. They
verify control counts, column/header/table captions, escaped ampersands,
callback bindings, clear-all visibility for multiple versus single filters,
the disabled empty state, deletion and callback-state cleanup, and valid,
invalid, and stale-header click dispatch. All output is also collected in a
module-level string. `ActiveFilters_SelfTestReport()` runs the same complete
test and returns the report, with lines separated by vbCrLf.

All three test arrays are explicitly allocated with lower and upper bounds
before access: `data(1 To 200, 1 To 25)`, `tableData(1 To 21, 1 To 3)`, and
`valueCriteria(1 To 2)`. Array-driven loops use LBound/UBound; the worksheet
writes and table source ranges use Resize with those matching bounds.
Production filter handling uses Collections, not arrays.

The current step is recorded before each setup operation and each assertion's
arguments are evaluated. An unexpected error includes its number, description,
and step, for example:

```text
FAIL: runtime error 9 (Subscript out of range) at step: build test data
```

Under Tools > Options > General > Error Trapping, use "Break on Unhandled
Errors" for a complete test run. "Break on All Errors" pauses even on handled
errors, before the cleanup handler runs. If paused, continue execution with F5
to let it report the failure and close the temporary workbook without saving;
do not click Reset if you expect automatic cleanup.

The 26-check test has been run successfully through Excel COM on Windows with
Microsoft 365 desktop Excel. Only the installed Office bitness was exercised;
there are no bitness-specific VBA declarations.
If cleanup fails, a FAIL line reports it; close the temporary workbook without
saving. Normal workbook-open/close events from other installed code may still
run when the temporary workbook is created or closed.

## Automated Runner

From the repository root in Windows PowerShell, run:

```powershell
& .\excel-active-filters\tools\Run-SelfTest.ps1
```

Enable "Trust access to the VBA project object model" in Excel File > Options
> Trust Center > Trust Center Settings > Macro Settings before running. If
access is blocked, the runner prints this setting and the policy caveat. It
does not change Trust Center, registry, or macro-security settings.

The runner creates a new hidden Excel.Application COM instance, verifies its
Hwnd belongs to a PID not already running, and prints that PID. It imports
both source exports into a new unsaved workbook and calls its
`ActiveFilters_SelfTestReport`. Automation startup does not load PERSONAL.XLSB;
the runner never attaches to an existing Excel instance and never opens or
saves PERSONAL.XLSB or existing work files. The self-test creates and closes
a second temporary workbook without saving.

The report is printed and saved to `tools/selftest-last.txt` (gitignored).
Exit code 0 requires all 26 PASS lines and a matching zero-FAIL summary;
failures, missing/inconsistent summaries, and cleanup errors return 1.
Finally, the host workbook is closed without saving, Excel is quit, and COM
references are released. If that recorded process remains alive after a
10-second exit wait, only that original process is stopped. Other Excel
processes are never stopped.

For diagnosing a native Excel crash, optionally pass `-TraceSteps`. This adds
private step logging only to the temporary imported module and writes
`tools/selftest-trace.txt` (gitignored). Normal runs import the files unchanged;
the optional instrumentation never modifies the source files. DisplayAlerts
does not suppress every VBA compile/debug dialog: unexpected errors can still
pause if Error Trapping is set to "Break on All Errors". The runner's final
10-second wait applies after the Run call returns, not to a blocked Run call.

## Ribbon Button

1. In Excel open File > Options > Customize Ribbon.
2. Set "Choose commands from" to "Macros".
3. Select `PERSONAL.XLSB!ShowActiveFiltersMenu`.
4. Select or create your custom group on the desired tab, then click Add.
5. Click Rename to set the label (for example, "Active Filters") and pick an icon.
6. Click OK. On a worksheet, click the button to open the menu at your cursor.

Choose a column row to clear that filter. With two or more active columns, a
separated "Clear all filters on this sheet" item is available. With only an
Advanced Filter active, its item clears the entire Advanced Filter. With no
active filters, a disabled status item appears. Escape/click outside dismisses
the menu. The temporary popup is deleted before creation and after dismissal.
Chart sheets and grouped sheets are rejected without changing filters.

## Behavior and Known Limits

- Macro-based filter clearing wipes Excel's undo history. There is no custom undo.
- PivotTable filters and slicers are not handled. Advanced Filter can only be
  cleared as a whole, not column by column.
- No filter criteria are inspected. Only the `On` flag is read, so colour/icon
  criteria do not need to be decoded. Header captions replace line breaks with
  spaces, truncate to 40 characters, and escape ampersands.
- Tables with hidden headers or disabled filter arrows are omitted from the
  column list. Clear all also attempts any active filters in those tables.
- A target must still belong to the active workbook/sheet, with the same source,
  range address, field, sheet column, and header. Otherwise nothing is cleared
  and the user is asked to reopen the menu. Same-named replacement objects with
  identical metadata cannot be distinguished by this name-based snapshot.
- Protected sheets are never unprotected. Clearing is attempted and failures
  on a protected sheet show a short message. Clear all is not transactional:
  an error can leave earlier filters cleared and later filters unchanged.
- Clearing does not select or activate cells. Scroll row/column are saved and
  restored for single and all-filter clearing; restoration is best-effort if
  another macro closes or invalidates the active window during the operation.
- ScreenUpdating and application event/security settings are not changed.
- The popup contains only one invocation's targets. There is no polling,
  automatic refresh, or persistent floating window.

## Manual UI Checks

After the self-test, use a disposable workbook to verify cursor placement,
Escape dismissal, repeated opening, Advanced Filter clearing, and captions
containing multi-line/long headers.
Also verify grouped-sheet/chart-sheet rejection and protected-sheet handling.
Test both Office bitness versions where available; no bitness-specific code
is used. These UI and protection checks are not covered by the automated VBA
self-test.