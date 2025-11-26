# Mechanical Calculation VBA System

Excel VBA macros for mechanical engineering calculations with automated sorting and data lookup functionality.

## Overview

This system consists of two main components:
1. **Priority-based sorting system** - Filters and sorts mechanical components by priority
2. **Cascading dropdown lookup** - Retrieves piping support data from reference tables

## Files

- `Module1.VBA` - Main calculation and sorting logic
- `Module2.VBA` - Legacy sorting implementation (not used)
- `Module3.VBA` - Sorting algorithm implementation
- `Module4.VBA` - Cascading dropdown functionality
- `Calcs_Sheet_Code.VBA` - Worksheet event handler

## Features

### 1. Mechanical Component Sorting (`Module1.VBA`, `Module3.VBA`)

**Purpose:** Automatically filters and sorts mechanical components based on priority values.

**How it works:**
- Scans rows 8-57 in the "Calcs" sheet
- Filters items where:
  - Column H contains "X" (preselection marker)
  - Column R is not empty
  - Column R ≤ 5 (priority threshold)
- Collects data from:
  - Column I: Component identifier/description
  - Column Q: Numeric value (quantity/specification)
  - Column R: Priority/ranking value
- Sorts by priority in ascending order
- Outputs results to columns A-C (starting row 8) in descending order
- Shows status notification: "Calculation complete - X items processed and sorted"

**To run:**
1. Mark items for processing with "X" in column H
2. Press Alt+F8 → Select `szam` → Run
3. Or let it auto-run when cells B2, B3, or B4 change

**Configuration:**
- **Priority threshold:** Change `<= 5` in lines 18 and 30 to adjust filtering
- **Row range:** Change `1 To 50` in lines 17 and 28 to scan more/fewer rows
- **Preselection column:** Change `"H"` to use a different column for selection marker

### 2. Cascading Dropdown Lookup (`Module4.VBA`)

**Purpose:** Provides interactive dropdowns to look up piping support mass data from the reference sheet.

**Location:** Cells F2-F5 in "Calcs" sheet

**Dropdowns:**
- **F2 - Material:** CS (Carbon Steel), SS (Stainless Steel), or Mapress
- **F3 - DN [mm]:** Pipe diameter (filtered by selected material)
  - CS: 15-250mm
  - SS: 15-350mm
  - Mapress: 20-100mm
- **F4 - Mass Type:** Choose from 4 options:
  - Gas w/o insulation
  - Gas w insulation
  - Liquid w/o insulation
  - Liquid w insulation
- **F5 - Result [kg]:** Automatically displays the mass value from reference sheet

**Setup:**
1. Open VBA Editor (Alt+F11)
2. Import Module4.VBA into Modules folder
3. Copy contents of `Calcs_Sheet_Code.VBA` into the Calcs sheet code window
4. Run `InitializeDropdowns` macro (Alt+F8 → InitializeDropdowns → Run)

**Usage:**
- Select Material → DN dropdown updates automatically
- Select DN → Choose Mass Type → Result appears automatically
- Changes in F2 clear F3 and F5
- All selections required for result to display

## Worksheet Events (`Calcs_Sheet_Code.VBA`)

Handles automatic triggers:
- **B2, B3, B4 changes** → Runs `szam` macro (if numeric value entered)
- **F2 changes** → Updates DN dropdown options
- **F3, F4 changes** → Updates result value in F5

## Data Structure

### Calcs Sheet
- **Input parameters:** Rows 1-3 (console length, force, moment)
- **Data table:** Rows 8+ with columns for profiles, deflection, stress
- **Column H:** Preselection marker ("X")
- **Column I:** Component name
- **Column Q:** Numeric value
- **Column R:** Priority value
- **Columns A-C:** Sorted output results
- **Columns F2-F5:** Dropdown lookup interface

### Spans_with_Mapress_from TV046 Sheet
Reference table with piping support span data:
- **Column A:** Material (CS, SS, Mapress)
- **Column B:** DN [mm] - Pipe diameter
- **Column C:** WT [mm] - Wall thickness
- **Columns D-O:** Various span distances, line masses, and mass per support values for different conditions

## Troubleshooting

**Items not appearing in sorted results:**
- Check if column H has "X" marker
- Verify column R value is ≤ 5
- Ensure row is within 8-57 range

**Dropdown not updating:**
- Verify Material is selected first
- Check if DN exists for that material in reference sheet
- Run `InitializeDropdowns` again if dropdowns are missing

**"Not found" in F5:**
- Verify all three dropdowns (F2, F3, F4) have selections
- Check that the Material/DN combination exists in reference sheet
- Confirm column names in reference sheet match expected format

**Status bar notification not clearing:**
- Normal behavior - clears after 3 seconds automatically
- Can be cleared manually: `Application.StatusBar = False` in Immediate Window

## Customization

### Change priority threshold:
```vba
' In Module1.VBA, lines 18 and 30
' Change from <= 5 to your desired value
If (Cells(i + 7, "H") = "X") And (Cells(i + 7, "R") <> "") And (Cells(i + 7, "R") <= 3) Then
```

### Scan more rows:
```vba
' In Module1.VBA, lines 17 and 28
' Change from 50 to desired number
For i = 1 To 100
```

### Change dropdown locations:
```vba
' In Module4.VBA, update all references to F2, F3, F4, F5
' Also update in Calcs_Sheet_Code.VBA
```

### Modify mass type options:
```vba
' In Module4.VBA, InitializeDropdowns sub
' Edit the Formula1 string
Formula1:="Option1,Option2,Option3,Option4"
```

## Technical Notes

- **Sorting algorithm:** Insertion sort (ascending order by priority)
- **Output order:** Descending (highest priority at top)
- **Array handling:** Dynamic sizing based on filtered item count
- **Event handling:** Uses `Application.EnableEvents` to prevent recursion
- **Status bar:** Auto-clears via `Application.OnTime` scheduled event
- **Data validation:** Uses Excel's built-in validation for dropdown lists

## Requirements

- Excel with macro support (.xlsm file)
- Two sheets: "Calcs" and "Spans_with_Mapress_from TV046"
- Macros enabled in Excel Trust Center settings

## Version History

- Initial version: Basic sorting with MsgBox notification
- v1.1: Added preselection filter (column H)
- v1.2: Extended row range from 13 to 50
- v1.3: Added empty cell check to prevent zero-filled results
- v1.4: Changed MsgBox to status bar notification with auto-clear
- v2.0: Added cascading dropdown lookup system
