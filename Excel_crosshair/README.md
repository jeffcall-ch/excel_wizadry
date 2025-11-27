# Excel Permanent Crosshair

A VBA solution that creates a **permanent yellow highlight** for the active cell's row and column in Excel. The crosshair automatically updates as you select different cells and works even with hidden rows/columns.

## Features

- ✅ Yellow background highlighting for active cell's row and column
- ✅ Updates automatically when you select different cells
- ✅ Works with hidden rows and columns
- ✅ Uses Excel's native conditional formatting (no shapes)
- ✅ Preserves original cell colors
- ✅ Persists when you save the workbook

## Installation

### Step 1: Add VBA Code

#### A. Add Main Module
1. Press `Alt+F11` to open VBA Editor
2. Go to **Insert → Module**
3. Copy all code from `excel_permanent_crosshair.VBA` and paste into the module
4. Name the module `CrosshairModule` (optional)

#### B. Add Workbook Events
1. In the VBA Editor, double-click **ThisWorkbook** in the Project Explorer
2. Paste the following code:

```vba
Private Sub Workbook_Open()
    SetupAllSheets
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Application.Calculate
End Sub
```

### Step 2: Save and Test
1. Save the workbook as **Excel Macro-Enabled Workbook (.xlsm)**
2. Close and reopen Excel
3. The crosshair should appear automatically - select different cells to see the yellow highlight move

## Usage

The crosshair will:
- Start automatically when you open the workbook (via `Workbook_Open` event)
- Highlight the entire row and column of the active cell with yellow background
- Update automatically as you select different cells
- Preserve the original colors when you move to a different cell

### Manual Control (Optional)

You can control it manually from VBA Immediate Window (`Ctrl+G`):

```vba
' Enable crosshair on active sheet only
StartCrosshair

' Enable crosshair on all sheets
SetupAllSheets

' Disable crosshair on active sheet
StopCrosshair

' Disable crosshair on all sheets
RemoveAllSheets
```

## How It Works

The solution uses **conditional formatting with formulas** to dynamically highlight cells:
- Formula for row highlighting: `=ROW()=CELL("row")`
- Formula for column highlighting: `=COLUMN()=CELL("col")`

The `CELL()` function returns the row/column of the active cell. When you select a new cell, Excel recalculates the conditional formatting (triggered by `SheetSelectionChange` event), and the yellow highlight moves automatically.

## Troubleshooting

### Crosshair doesn't appear
- Check if macros are enabled (File → Options → Trust Center → Macro Settings)
- Verify the workbook events are added to `ThisWorkbook`
- Try running `SetupAllSheets` manually from VBA

### Crosshair doesn't update when selecting cells
- Make sure the `Workbook_SheetSelectionChange` event is in `ThisWorkbook`
- The event calls `Application.Calculate` to trigger the conditional formatting update

### Want to remove the crosshair
- Run `RemoveAllSheets` to remove from all sheets
- Or manually delete the conditional formatting rules (Home → Conditional Formatting → Manage Rules)

## Customization

### Change Highlight Color
In the code, find `.Interior.Color` and change the RGB values:
```vba
.Interior.Color = RGB(255, 255, 0)  ' Yellow
' Change to:
.Interior.Color = RGB(255, 200, 200)  ' Light red
.Interior.Color = RGB(200, 255, 200)  ' Light green
.Interior.Color = RGB(200, 200, 255)  ' Light blue
```

### Change Coverage Area
In `SetupConditionalFormatting`, adjust the range expansion:
```vba
' Current: adds 100 rows and 26 columns to used range
Set usedRange = ws.Range(usedRange.Address).Resize(usedRange.Rows.count + 100, _
                                                     usedRange.Columns.count + 26)
' Change numbers to cover more or less area
```

### Use Borders Instead of Background
Replace the `.Interior.Color` lines with:
```vba
With .Borders(xlTop)
    .LineStyle = xlDash
    .Color = RGB(255, 0, 0)  ' Red
    .Weight = xlMedium
End With
```

## Limitations

- Only works in Excel (uses VBA and conditional formatting)
- Requires macros to be enabled
- The `CELL()` function requires calculation to update (handled by event handler)
- Conditional formatting rules are saved with the workbook

## Requirements

- Microsoft Excel (2010 or later)
- VBA macros enabled
- Workbook saved as .xlsm (macro-enabled format)

## License

Free to use and modify for personal or commercial purposes.
