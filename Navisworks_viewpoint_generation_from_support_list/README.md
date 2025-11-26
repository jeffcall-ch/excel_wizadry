# Navisworks Viewpoint Generation & Excel Comment Replicator

This folder contains two automation tools for working with construction project data:

1. **Python Script**: Generates Navisworks XML viewpoints from Excel support list
2. **Excel VBA Macro**: Automatically replicates comments across matching rows and displays live statistics

---

## 1. Navisworks Viewpoint Generator (Python)

### Purpose
Automatically generates Navisworks viewpoint XML files from an Excel support list. Each viewpoint includes:
- Camera position and rotation (oblique angle view)
- 2m × 2m × 2m clipping box centered on support location
- Hierarchical folder organization based on zones and support references

### Features
- **4-Level Folder Hierarchy**: GP_SU → Parent Folder → Zone → Subfolder → Viewpoints
- **AUX Folder Merging**: `/0AUX-P/PW` and `/1AUX-P/PW` automatically grouped under parent `AUX-P/PW`
- **Deterministic GUIDs**: Same input always generates same viewpoint IDs (reproducible)
- **Coordinate Parsing**: Extracts X, Y, Z from "X 18040mm Y 94404.038mm Z 186.7mm" format
- **Consistent Camera Angles**: All viewpoints use same oblique viewing angle with quaternion rotation

### Requirements
```bash
pip install pandas openpyxl
```

### Input Data Structure
Excel file with columns:
- **Column D**: SU/STEEL ref (e.g., `/0GAC05BQ003/SU`) - First 7 chars used for subfolder name
- **Column H**: BQ Zone (e.g., `/0AUX-P/PW`) - Used for parent folder and zone organization
- **Column O**: POS WRT /* (e.g., `X 18040mm Y 94404.038mm Z 186.7mm`) - 3D coordinates in millimeters

### Usage
```bash
python create_view_points.py
```

### Configuration
Edit these constants in `create_view_points.py`:

```python
# Input/Output Files
EXCEL_FILE = "path/to/your/GP_BQ_25.11.2025.xls"
OUTPUT_XML = "generated_viewpoints.xml"

# Camera Settings
CAMERA_OFFSET_X = 3.4516658778  # meters from target point
CAMERA_OFFSET_Y = 2.9757158425
CAMERA_OFFSET_Z = 2.5844675346

# Clipping Box Size (half-widths)
BOX_HALF_WIDTH_X = 1.0  # 2m total width
BOX_HALF_WIDTH_Y = 1.0  # 2m total depth
BOX_HALF_WIDTH_Z = 1.0  # 2m total height
```

### Output
Generates `generated_viewpoints.xml` with structure:
```
📁 GP_SU
  📁 AUX-P/PW (Parent - numeric prefix stripped)
    📁 /0AUX-P/PW (Original zone)
      📁 /0GAC05 (First 7 chars from column D)
        📷 /0GAC05BQ003/SU (Viewpoint)
    📁 /1AUX-P/PW (Original zone)
      📁 /1GAC07 
        📷 /1GAC07BQ001/SU
```

### Example Output Stats
```
✓ Successfully generated 1158 viewpoints
✓ Organized into 26 parent folders
✓ Containing 39 zone folders
```

### How Folder Hierarchy Works

1. **Top Level**: Always `GP_SU` (fixed)
2. **Parent Folder**: Column H with `/[0-9]` prefix stripped
   - `/0AUX-P/PW` → `AUX-P/PW`
   - `/1AUX-P/PW` → `AUX-P/PW` (merges to same parent)
3. **Zone Folder**: Column H as-is (keeps `/0` prefix)
4. **Subfolder**: First 7 characters from Column D
5. **Viewpoint**: Full value from Column D

---

## 2. Excel VBA Comment Replicator

### Purpose
Automatically replicates column P (comments) to all rows with matching column D (SU/STEEL ref) values. Displays live statistics showing unique count of column D values in the status bar.

### Features
- ✅ **Auto-Replication**: Edit column P once, updates all matching rows instantly
- ✅ **Case-Insensitive Matching**: `/0GAC05` matches `/0gac05`
- ✅ **Filter-Aware Statistics**: Status bar reflects filtered/unfiltered views
- ✅ **Performance Optimized**: Smart caching prevents slowdown with 4000+ rows
- ✅ **Visual Feedback**: Shows replication success message for 20 seconds
- ✅ **Single File Solution**: No separate modules required

### Installation

#### Step 1: Enable Macros
1. Open your Excel file
2. Go to **File → Options → Trust Center → Trust Center Settings**
3. Select **Macro Settings**
4. Choose **Enable all macros** (or enable for this workbook only)
5. Click **OK**

#### Step 2: Install VBA Code
1. Open your Excel file
2. Press **ALT + F11** to open VBA Editor
3. In the **Project Explorer** (left panel):
   - Expand "VBAProject (your-filename.xlsm)"
   - Expand "Microsoft Excel Objects"
   - **Double-click your worksheet** (e.g., "Sheet1", "BQ", etc.)
4. **Copy ALL code** from `SU_based_comment_replicator.VBA`
5. **Paste** into the worksheet code window
6. Press **ALT + Q** to close VBA Editor
7. **Save file as .xlsm** (macro-enabled format)

### How It Works

#### Automatic Replication
1. Edit any cell in **column P** (comments)
2. Press **ENTER**
3. Macro finds all rows with same **column D** value (case-insensitive)
4. Copies your comment to all matching rows
5. Status bar shows: `✓ Replicated to X row(s) in 0.XXs` (displays for 20 seconds)

#### Live Statistics
Status bar continuously displays:
- **Unfiltered**: `Column D: 3194 rows | 1168 unique values`
- **Filtered**: `Column D: 150 visible rows | 80 unique values (filtered)`

Statistics update automatically when:
- You apply/remove filters
- You click on cells (restores if overwritten)
- Calculation events fire

#### Valid Column D Values
Replication only occurs when column D contains:
- ✅ Non-empty values
- ✅ Not "-" or "N/A"
- ✅ Not Excel errors (#N/A, #REF!, etc.)

### Performance Features

#### Smart Caching
- Filter state tracked (only recalculates on actual filter changes)
- Status message cached (restored without recalculation)
- No slowdown when navigating cells

#### Optimizations During Replication
- Events disabled to prevent cascading triggers
- Screen updating paused for smooth execution
- Manual calculation mode during operation
- Automatic restoration after completion

### Troubleshooting

#### Code Doesn't Run
**Problem**: Nothing happens when editing column P

**Solutions**:
1. Check macros are enabled (see Installation Step 1)
2. Verify code is in **worksheet module**, not a standard module:
   - VBA Editor → Project Explorer → Microsoft Excel Objects → Sheet1
   - NOT in "Modules" section
3. Check if Design Mode is enabled:
   - Developer tab → Design Mode should be **OFF**
4. Verify file saved as `.xlsm` (macro-enabled)

#### Status Bar Not Updating
**Problem**: Statistics don't update after filtering

**Solutions**:
1. Click any cell after applying filter
2. Close and reopen the worksheet
3. Press **ALT + F11** → Immediate Window (Ctrl+G) → Type:
   ```vba
   Application.EnableEvents = True
   ```

#### Slow Performance
**Problem**: Excel freezes when clicking cells

**Solutions**:
1. This should not happen with current optimized code
2. If it does, check for other add-ins or macros interfering
3. Try closing other Excel workbooks

#### Replication Not Working for Specific Rows
**Problem**: Some rows don't get updated

**Possible Causes**:
- Column D contains "-", "N/A", or is empty
- Column D contains Excel error values
- Trailing/leading spaces in column D (macro handles this automatically)

### Code Structure

```
📄 SU_based_comment_replicator.VBA
├── Module-Level Variables
│   ├── previousFilterState: Tracks filter changes
│   └── lastStatusMessage: Cached status text
│
├── Event Handlers
│   ├── Worksheet_Activate(): Initialize on sheet open
│   ├── Worksheet_SelectionChange(): Update stats on cell navigation
│   ├── Worksheet_Calculate(): Backup for filter change detection
│   └── Worksheet_Change(): Main replication logic (column P edits)
│
├── Helper Functions
│   ├── GetFilterState(): Serialize current filter configuration
│   └── UpdateUniqueCountStatus(): Count unique column D values
│
└── Replication Logic
    ├── Validate column D value
    ├── Find all matching rows (case-insensitive)
    ├── Copy column P value to matches
    ├── Show success message (20 seconds)
    └── Restore statistics display
```

### Customization

#### Change Display Duration
Edit line in `Worksheet_Change`:
```vba
Do While Timer < pauseTime + 20  ' Change 20 to desired seconds
```

#### Change Target Columns
Edit these lines:
```vba
If Target.Column <> 16 Then Exit Sub  ' Column P (change 16 for different column)
dValue = Me.Cells(changedCell.Row, 4).Value  ' Column D (change 4 for different column)
```

Column numbers: A=1, B=2, C=3, D=4, ... P=16

#### Exclude Different Values
Edit validation section:
```vba
If UCase(dValue) <> "-" And UCase(dValue) <> "N/A" And Len(dValue) > 0 Then
```

### Version History

- **v1.0** - Initial release with basic replication
- **v2.0** - Added status bar statistics
- **v3.0** - Filter-aware statistics with smart caching
- **v4.0** - Performance optimization (removed constant recalculation)
- **v5.0** - Current version with 20-second message display

---

## File Structure

```
Navisworks_viewpoint_generation_from_support_list/
├── README.md                              # This file
├── create_view_points.py                  # Python viewpoint generator
├── SU_based_comment_replicator.VBA        # Excel VBA macro
├── SU_based_comment_replicator_SIMPLE.VBA # Testing version (with popups)
├── GP_BQ_25.11.2025.xls                   # Input Excel file
└── generated_viewpoints.xml               # Output (auto-generated)
```

## Support

For issues or questions:
1. Check the Troubleshooting sections above
2. Review code comments for detailed explanations
3. Test with the SIMPLE.VBA version for diagnostics

## License

Internal tool for construction project automation.
