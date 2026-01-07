# PTS Checker Script Analysis Report

## Issue Summary
The script is **NOT finding** the mobilisation activity and other important activities because it's **reading the WRONG sheet** from the Excel file.

## Root Cause Analysis

### 1. **Wrong Sheet Being Read**
The script at line 342 reads:
```python
df_plan = pd.read_excel(project_plan_path)
```

**Problem**: When `sheet_name` parameter is NOT specified, `pd.read_excel()` defaults to reading **the FIRST sheet only**.

### 2. **File Structure Analysis**

**PTS_project_specific.xlsx** contains **5 sheets**:
1. **'Boiler Only'** ← **THIS IS BEING READ (271 rows, WRONG SHEET)**
2. 'Baseline R1.1' (9,819 rows, CORRECT SHEET ✓)
3. 'DD250924 Complete PTS (2)' (12,140 rows)
4. 'DD250724 Complete PTS' (12,241 rows)
5. 'DD250724 Complete PTS rA' (12,241 rows)

### 3. **Evidence**

**Sheet 1: 'Boiler Only' (Currently being read)**
- Total rows: 271
- Non-null Activity Names: 258
- Contains 'Mobilisation': **0 rows** ❌
- Contains 'Lot 5': **0 rows** ❌
- Contains 'General Piping': **0 rows** ❌
- Activities are only about: "Lot Boiler - Buckstays", "Lot Boiler - Front Wall", etc.

**Sheet 2: 'Baseline R1.1' (Should be read)**
- Total rows: 9,819
- Non-null Activity Names: 9,819
- Contains 'Mobilisation': **39 rows** ✓
- Contains 'Lot 5': **39 rows** ✓
- Contains 'General Piping': **145 rows** ✓
- **Excel Row 7294**: `ORBE1080A0 - Lot 5 - Erection General Piping - Contractor On Site Mobilisation` ✓

### 4. **Why It Worked Before**

The script worked for a different PTS file because:
- The previous file likely had **only ONE sheet**, OR
- The **first sheet** contained all the activities needed

## Detailed Findings

### Activities Found in 'Baseline R1.1' Sheet
The mobilisation activity you mentioned (row 7294) EXISTS in the **'Baseline R1.1'** sheet:
- **Excel Row 7294**: `ORBE1080A0 - Lot 5 - Erection General Piping - Contractor On Site Mobilisation`
- **Excel Row 7295**: `ORBE1090M0 - Lot 5 - Erection General Piping - Contractor On Site Mobilisation Completed`

### All Important Activities from General File
The general file is looking for these 24 activities:
1. General Piping - BOM (for RfQ Erection General Piping, for TSD PO General PIping Engineering )
2. Lot General Piping Engineering - Purchase Order
3. Lot General Piping Engineering - BOM / BOQ-1 (for RfQ GP Piping Material, Valves, Hangers...)
4. Lot General Piping Engineering - BOM / BOQ-2 (for PO General Piping Material)
5. Lot General Piping Valves - TSD / BANF RFQ
6. Lot General Piping Valves - TSD PO
7. Lot General Piping Hangers - TSD / BANF RFQ
8. Lot General Piping Hangers - Purchase Order
9. Lot General Piping Hangers - Ready for Shipment
10. Lot General Piping Additional Material - TSD / BANF RFQ
11. Lot General Piping Pipe Material - TSD / BANF RFQ
12. Lot General Piping Pipe Material - TSD PO
13. Lot General Piping Pipe Material - Purchase Order
14. Lot General Piping Pipe Material - Ready for Shipment
15. **Lot 5 - Erection General Piping - Contractor On Site Mobilisation** ← ANCHOR ACTIVITY
16. Lot General Piping Hangers - DTS (S)
17. Lot General Piping Pipe Material  - DTS (S)
18. Lot 5 - Erection General Piping - BOQ 1 (Doc's for RfQ) - Available from Engineering
19. Lot 5 - Erection General Piping - BOQ 2 (Doc's for PO) - Available from Engineering
20. Lot General Piping Additional Material - Purchase Order
21. Lot General Piping Additional Material - DTS (S)
22. Lot General Piping Valves - Purchase Order
23. Lot General Piping Valves - Ready for Shipment
24. Lot General Piping Valves - DTS (S)

**ALL of these activities are likely in the 'Baseline R1.1' sheet**, NOT in 'Boiler Only'.

## Solution Required

The script needs to be modified to:
1. **Specify the correct sheet name** when reading the Excel file
2. **OR**: Make it configurable (add parameter to specify sheet name)
3. **OR**: Read all sheets and search across all of them
4. **OR**: Detect which sheet has the most activities and use that one

## Recommended Fixes (in order of preference)

### Option 1: Make sheet name a parameter (BEST)
Add a parameter to the function to specify which sheet to read:
```python
def process_pts_file(project_plan_path, general_file_path, 
                     sheet_name='Baseline R1.1',  # Add this parameter
                     similarity_threshold=95):
    df_plan = pd.read_excel(project_plan_path, sheet_name=sheet_name)
```

### Option 2: Auto-detect largest sheet
```python
xl = pd.ExcelFile(project_plan_path)
# Find sheet with most rows
max_rows = 0
best_sheet = None
for sheet in xl.sheet_names:
    df_temp = pd.read_excel(project_plan_path, sheet_name=sheet)
    if len(df_temp) > max_rows:
        max_rows = len(df_temp)
        best_sheet = sheet
df_plan = pd.read_excel(project_plan_path, sheet_name=best_sheet)
```

### Option 3: Hardcode the correct sheet name
```python
df_plan = pd.read_excel(project_plan_path, sheet_name='Baseline R1.1')
```

## Impact Analysis

**Current Behavior:**
- Script reads 271 rows from 'Boiler Only'
- Finds 0 matches for important activities
- Generates empty output with no markers
- All date calculations fail (no anchor activity found)
- User sees warnings about missing activities

**Expected Behavior (after fix):**
- Script reads 9,819 rows from 'Baseline R1.1'
- Finds all 24 important activities
- Generates proper output with category markers
- Date calculations work correctly
- No warnings about missing activities

## Verification Steps

To verify the fix works:
1. Modify script to read 'Baseline R1.1' sheet
2. Run the script
3. Check output for:
   - Mobilisation activity found at row 7294
   - All 24 activities marked with category flags
   - Requested dates calculated correctly
   - No missing activity warnings
