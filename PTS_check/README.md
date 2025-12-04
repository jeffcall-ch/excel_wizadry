# PTS Activity Matcher & Date Calculator

## Overview

This script automates the process of matching important activities from a general requirements list against a detailed project plan, then calculates required dates for these activities based on dependency relationships and holiday constraints.

**Key Features:**
- Fuzzy string matching (90% similarity threshold) to find activities
- Bidirectional date calculation (backward from milestone, forward from dependencies)
- Holiday period detection with automatic duration adjustments
- Weekday validation (all dates fall Monday-Friday)
- Conflict detection when multiple dependencies produce different dates
- Excel output with all rows, filtered to important activities by default

---

## Input Files

### 1. Project Plan (`PTS_project_specific.xlsx`)
- **Columns Required:** `Activity ID`, `Activity Name`, `Start`, `Finish`
- **Format:** Standard MS Project export with Swiss date format (dd/mm/yyyy)
- **Size:** Typically 9,000-11,000 activities
- **Purpose:** Contains all project activities with current planned dates

### 2. General Input Data (`PTS_General_input_data.xlsx`)

#### Sheet 1: Activities
Defines the 24 important activities to track:
- **Columns:** `Activity`, `Important for GP`, `Piping`, `Hangers`, `Additional Material`, `Valves`
- **Format:** Activity name (text), category markers ('X' for applicable categories)
- **Example:**
  ```
  Activity: "Lot General Piping Engineering - BOM / BOQ-1 (for RfQ GP Piping Material, Valves, Hangers...)"
  Important for GP: X
  Piping: X
  Hangers: 
  Valves:
  ```

#### Sheet 2: Relationships
Defines dependencies between activities with duration constraints:
- **Columns:** `From`, `To`, `Duration weeks`, `Comments`
- **Format:** 
  - `From`: Activity name (predecessor)
  - `To`: Activity name (successor)
  - `Duration weeks`: Number of weeks between activities
  - `Comments`: Optional description
  
- **Example Relationships:**
  ```
  From: "Lot General Piping Engineering - BOM / BOQ-1"
  To: "Lot General Piping Pipe Material - TSD / BANF RFQ"
  Duration weeks: 4
  Comments: "Possible 3W to 5W"
  ```

**Interpretation:** Activity "TSD / BANF RFQ" needs 4 weeks AFTER "BOM / BOQ-1" completes (forward relationship), OR "BOM / BOQ-1" must finish 4 weeks BEFORE "TSD / BANF RFQ" starts (backward relationship).

#### Sheet 3: Holidays
Defines holiday periods that affect activity scheduling:
- **Columns:** `From`, `To`, `Length adder weeks`
- **Format:**
  - `From`: Text description (e.g., "middle of June")
  - `To`: Text description (e.g., "end of August")
  - `Length adder weeks`: Integer (weeks to add when activity overlaps holiday)

- **Example:**
  ```
  From: "middle of December"
  To: "end of December"
  Length adder weeks: 3
  ```

**Supported date descriptions:**
- "middle of [Month]" → 15th of month
- "end of [Month]" → Last day of month (28-31 depending on month)

---

## Date Calculation Logic

### Anchor Point

The entire calculation starts from a single **anchor milestone**:
- **Activity:** "Lot 5 - Erection General Piping - Contractor On Site Mobilisation"
- **Date Source:** Uses the **START date** from the project plan
- **Direction:** All other dates are calculated relative to this anchor

**Example:**
```
Anchor: Mobilisation START = 17/01/2028
```

### Backward Calculation (Before Anchor)

Calculates dates for activities that must complete BEFORE another activity.

#### Basic Logic
1. **Start with known date** (e.g., Mobilisation = 17/01/2028)
2. **Find relationship** "Activity X is 5W before Mobilisation"
3. **Calculate initial date:** 17/01/2028 - 5 weeks = 13/12/2027
4. **Check weekday:** Adjust if weekend (Saturday → Monday, Sunday → Monday)
5. **Check holiday overlap:** Does the execution period overlap with holidays?
6. **Apply adjustment if needed**
7. **Result:** Activity X required date

#### Execution Period for Backward Calculation

**Key Concept:** The duration represents how long the activity takes to execute.

If "Activity X is 5W before Anchor (17/01/2028)":
- Initial calculation: X = 17/01/2028 - 5W = 13/12/2027
- **Execution period:** From 13/12/2027 TO 17/01/2028 (5 weeks)
- This period represents the time needed for Activity X to complete before reaching the Anchor

#### Example 1: Backward with Holiday Adjustment

**Relationship:**
```
From: "Lot General Piping Pipe Material - Ready for Shipment"
To: "Lot 5 - Erection General Piping - Contractor On Site Mobilisation"
Duration weeks: 5
```

**Calculation Steps:**

1. **Anchor date:** 17/01/2028 (Thursday)

2. **Initial backward calculation:**
   - Ready for Shipment = 17/01/2028 - 5W = 13/12/2027

3. **Weekday check:**
   - 13/12/2027 = Wednesday ✓ (no adjustment needed)

4. **Holiday overlap check:**
   - Execution period: 13/12/2027 to 17/01/2028
   - Holiday period: middle of December (15/12/2027) to end of December (31/12/2027)
   - **Overlap detected!** ✓ (execution period overlaps with December holiday)
   - Adder: 3 weeks

5. **Apply holiday adjustment:**
   - New duration: 5W + 3W = 8W
   - Adjusted date: 17/01/2028 - 8W = 22/11/2027

6. **Final weekday check:**
   - 22/11/2027 = Monday ✓

7. **Result:**
   - Required date: 22/11/2027
   - Comment: "8W before 'Lot 5 - Erection General Piping - Contractor On Site Mobilisation'. START date of Lot 5 - Erection General Piping - Contractor On Site Mobilisation. Originally 13/12/2027 (based on 5W), shifted by 3W due to middle of December to end of December holiday period"

**Why the adjustment?** The shipment activity execution overlaps with the December holiday season, requiring more buffer time to ensure completion.

### Forward Calculation (After Activities)

Calculates dates for activities that must happen AFTER another activity completes.

#### Basic Logic
1. **Start with known date** (e.g., BOM/BOQ-1 = 19/02/2026)
2. **Find relationship** "Activity Y is 4W after BOM/BOQ-1"
3. **Calculate initial date:** 19/02/2026 + 4 weeks = 19/03/2026
4. **Check weekday:** Adjust if weekend
5. **Check holiday overlap:** Does the execution period overlap with holidays?
6. **Apply adjustment if needed**
7. **Result:** Activity Y required date

#### Execution Period for Forward Calculation

If "Activity Y is 4W after Activity X (19/02/2026)":
- Initial calculation: Y = 19/02/2026 + 4W = 19/03/2026
- **Execution period:** From 19/02/2026 TO 19/03/2026 (4 weeks)
- This period represents the time needed for Activity Y to execute after Activity X completes

#### Example 2: Forward with Holiday Adjustment

**Relationship:**
```
From: "Lot General Piping Hangers - TSD / BANF RFQ"
To: "Lot General Piping Hangers - Purchase Order"
Duration weeks: 7
Comments: "Possible 6W to 8W"
```

**Calculation Steps:**

1. **Known predecessor date:** TSD / BANF RFQ = 17/05/2027 (Monday)

2. **Initial forward calculation:**
   - Purchase Order = 17/05/2027 + 7W = 05/07/2027

3. **Weekday check:**
   - 05/07/2027 = Monday ✓

4. **Holiday overlap check:**
   - Execution period: 17/05/2027 to 05/07/2027
   - Holiday period: middle of June (15/06/2027) to end of August (31/08/2027)
   - Extended holiday: 15/06/2027 to 07/09/2027 (includes 1 week after)
   - **Overlap detected!** ✓ (execution period overlaps with summer holiday)
   - Adder: 4 weeks

5. **Apply holiday adjustment:**
   - New duration: 7W + 4W = 11W
   - Adjusted date: 17/05/2027 + 11W = 02/08/2027

6. **Final weekday check:**
   - 02/08/2027 = Monday ✓

7. **Result:**
   - Required date: 02/08/2027
   - Comment: "11W after 'Lot General Piping Hangers - TSD / BANF RFQ'. Possible 6W to 8W. Originally 05/07/2027 (based on 7W), shifted by 4W due to middle of June to end of August holiday period"

**Why the adjustment?** The purchase order processing period overlaps with the summer holiday season, requiring additional time for completion.

---

## Holiday Logic in Detail

### Holiday Period Detection

The script converts text descriptions to actual date ranges for each year (anchor year ± 1):

**Example Conversion:**
```
Input: "middle of June to end of August"
Year: 2027
Output: 15/06/2027 to 31/08/2027
```

### Extended Holiday Period

For detection purposes, holidays include **1 week after** the end date:
```
Base holiday: 15/06/2027 to 31/08/2027
Extended for checking: 15/06/2027 to 07/09/2027
```

This ensures activities starting shortly after holidays are also adjusted.

### Overlap Detection Algorithm

**An activity execution period overlaps with a holiday if:**
```
activity_start <= holiday_extended_end AND activity_end >= holiday_start
```

This catches:
- ✓ Activity entirely within holiday
- ✓ Activity starts before holiday, ends during holiday
- ✓ Activity starts during holiday, ends after holiday
- ✓ Activity starts before holiday, ends after holiday
- ✓ Activity starts/executes within 1 week after holiday ends

### When Holidays Are Applied

**Backward calculation example:**
```
Activity needs 5W before Anchor
Initial: 13/12/2027
Execution: 13/12/2027 → 17/01/2028
Holiday: 15/12/2027 → 31/12/2027
Overlap? YES (execution overlaps December 15-31)
Action: Add 3W to duration → 8W total
Result: Start earlier at 22/11/2027
```

**Forward calculation example:**
```
Activity needs 7W after predecessor
Initial: 05/07/2027
Execution: 17/05/2027 → 05/07/2027
Holiday: 15/06/2027 → 31/08/2027
Overlap? YES (execution overlaps June 15 - August 31)
Action: Add 4W to duration → 11W total
Result: End later at 02/08/2027
```

### Multiple Holiday Periods

The script checks ALL holiday periods for the relevant years:
- Anchor year - 1
- Anchor year
- Anchor year + 1

**Example with Anchor = 17/01/2028:**
- Generates holidays for 2027, 2028, 2029
- Each calculation checks all 6 holiday periods (2 periods × 3 years)
- Applies the FIRST matching holiday adder found

---

## Weekday Validation

All calculated dates must fall on weekdays (Monday-Friday).

### Adjustment Rules
```python
if date is Saturday (weekday = 5):
    date = date + 2 days  # Move to Monday
elif date is Sunday (weekday = 6):
    date = date + 1 day   # Move to Monday
```

### Application Points
1. **After initial calculation** (before holiday check)
2. **After holiday adjustment** (after adding adder weeks)
3. **Anchor date** is also validated

**Example:**
```
Calculated: 13/11/2027 (Saturday)
Adjusted: 15/11/2027 (Monday)
```

---

## Iterative Calculation Process

The script uses an iterative approach to resolve all dependencies:

### Algorithm
```
1. Start with anchor milestone (known date)
2. Initialize iteration counter (max 100 iterations)
3. Loop until no new dates calculated:
   a. For each relationship:
      - Check if "To" date known and "From" unknown → Calculate backward
      - Check if "From" date known and "To" unknown → Calculate forward
   b. Track if any new dates were calculated this iteration
   c. If no progress made → Stop (all calculable dates found)
4. Report activities with calculated dates
```

### Why Iterative?

Activities form dependency chains:
```
A → B (4W) → C (5W) → D (3W)
```

**Iteration 1:**
- Know: A = 01/01/2027
- Calculate: B = 29/01/2027

**Iteration 2:**
- Know: B = 29/01/2027
- Calculate: C = 05/03/2027

**Iteration 3:**
- Know: C = 05/03/2027
- Calculate: D = 26/03/2027

**Result:** After 3 iterations, all dates calculated.

### Example from Your Data

**Dependency Chain:**
```
BOM/BOQ-1 (19/02/2026)
  ↓ 4W after
TSD / BANF RFQ (19/03/2026)
  ↓ 7W after
Purchase Order (02/08/2027 - with holiday adjustment)
  ↓ 13W after
Ready for Shipment (29/11/2027 - with holiday adjustment)
  ↓ 4W before
Mobilisation (17/01/2028 - anchor)
```

Each step builds on the previous, requiring multiple iterations to complete the entire chain.

---

## Conflict Detection

### What is a Conflict?

A conflict occurs when **multiple dependency relationships** try to set **different dates** for the same activity.

### Example Conflict from Your Data

**Activity:** "Lot General Piping Hangers - Purchase Order"

**Path 1 (Backward from Ready for Shipment):**
```
From: Ready for Shipment (29/11/2027)
Relationship: 13W before Ready for Shipment
Holiday adjustment: +4W (summer holiday overlap)
Calculation: 29/11/2027 - (13W + 4W) = 02/08/2027
```

**Path 2 (Forward from TSD / BANF RFQ):**
```
From: TSD / BANF RFQ (17/05/2027)
Relationship: 7W after TSD / BANF RFQ
No holiday overlap
Calculation: 17/05/2027 + 7W = 05/07/2027
```

**Result:**
- Path 1 says: 02/08/2027
- Path 2 says: 05/07/2027
- **Conflict!** 4 weeks difference

### Why Conflicts Occur

1. **Inconsistent constraints:** Duration estimates conflict between relationships
2. **Holiday effects:** Different paths have different holiday overlaps
3. **Chain propagation:** Errors compound through long dependency chains
4. **Flexible ranges:** Comments like "Possible 12W to 14W" suggest uncertainty

### Conflict Resolution

The script **reports all conflicts** in a dedicated Excel sheet but uses **first calculated value**:
- First path to calculate a date "wins"
- Subsequent paths that give different dates are flagged as conflicts
- Review Conflicts sheet to identify which relationships need adjustment

### Example Conflict Output

```
CONFLICT: 'Lot General Piping Hangers - Purchase Order'
  Existing: 02/08/2027 from 17W before 'Lot General Piping Hangers - Ready for Shipment'. 
            Possible 12W to14W. Originally 30/08/2027 (based on 13W), shifted by 4W due 
            to middle of June to end of August holiday period
  New calc: 05/07/2027 from 7W after 'Lot General Piping Hangers - TSD / BANF RFQ'. 
            Possible 6W to 8W
```

---

## Complete Calculation Example

Let's trace one activity through the entire process:

### Activity: "Lot General Piping Engineering - BOM / BOQ-1"

**Step 1: Find in Project Plan**
```
Fuzzy match against all 9,194 activities
Best match: "Lot General Piping Engineering - BOM / BOQ-1 (for RfQ GP Piping Material, Valves, Hangers...)"
Similarity score: 94%
Current plan date: (varies)
```

**Step 2: Identify Relationships**
```
Relationship found:
  From: BOM/BOQ-1
  To: Lot General Piping Pipe Material - TSD / BANF RFQ
  Duration: 4W
```

**Step 3: Backward Calculation Chain**

The script knows the anchor (17/01/2028) and works backward through the chain:

```
Mobilisation (anchor) = 17/01/2028
  ↓ 5W before (with 3W holiday adjustment)
Ready for Shipment = 22/11/2027
  ↓ 16W before
Purchase Order = 19/07/2027
  ↓ 10W before
TSD PO = 10/05/2027
  ↓ 2W before
BOM/BOQ-1 = 19/04/2027 (conflicting path)
```

**Alternative Forward Path:**
```
(Some earlier activity) = XX/XX/XXXX
  ↓ YW after
BOM/BOQ-1 = 17/05/2027 (conflicting path)
```

**Step 4: Check for Holidays**
```
Execution period: Depends on relationship direction
Holiday periods checked: 2027, 2028, 2029 summers and winters
Adjustment: Applied if overlap detected
```

**Step 5: Validate Weekday**
```
Final date: Adjusted if Saturday/Sunday
Result: Confirmed weekday date
```

**Step 6: Generate Comment**
```
If no holiday adjustment:
  "2W before 'Lot General Piping Pipe Material - TSD PO'"

If holiday adjustment:
  "6W before 'Activity X'. Originally 15/07/2027 (based on 2W), 
   shifted by 4W due to middle of June to end of August holiday period"
```

---

## Output File Structure

### Main Sheet (All Activities)

**Columns (in order):**
1. `Activity ID` - Project activity code (e.g., "ETAB1060M0")
2. `Activity Name` - Full activity description
3. `Start` - Current planned start date (dd/mm/yyyy)
4. `Finish` - Current planned finish date (dd/mm/yyyy)
5. `Current length` - Actual duration between related activities with direction
6. `Requested date` - Calculated required date (dd/mm/yyyy)
7. `Comment` - Calculation explanation with holiday adjustments
8. `Similarity Score %` - Fuzzy match score (90-100)
9. `Important for GP` - X if important for general piping
10. `Piping` - X if piping category
11. `Hangers` - X if hangers category
12. `Additional Material` - X if additional material category
13. `Valves` - X if valves category

**Row Count:** All rows from project plan (typically 10,967)

**Active Filter:** "Important for GP" column filtered to show only "X" values (24 activities visible by default)

**Formatting:**
- Column A width: 20
- Column B width: 75
- All cells: Wrap text, vertical center alignment
- Category columns: Horizontal center alignment
- Yellow highlight: Activity Name cells with similarity score < 95%
- Frozen panes: At C2 (freezes headers and columns A-B)
- Autofilter: Applied to all columns

### Current Length Column Format

Shows actual duration between activities matching the Comment direction:

**Example entries:**
```
"14.7W (planned: 14W) before 'Lot General Piping Engineering - BOM / BOQ-1'"
"5.1W (planned: 4W) after 'Lot General Piping Engineering - BOM / BOQ-1'"
"11.3W (planned: 5W) before 'Lot 5 - Erection General Piping - Contractor On Site Mobilisation'"
```

**Multiple dependencies:**
```
"14.7W (planned: 14W) before 'Activity A'
5.1W (planned: 4W) after 'Activity B'"
```

### Conflicts Sheet

Created only if conflicts detected.

**Columns:**
- `Conflict Description` - Full conflict details (width: 100, wrap text)

**Content Format:**
```
CONFLICT: '[Activity Name]'
  Existing: [Date] from [Calculation path 1]
  New calc: [Date] from [Calculation path 2]
```

**Example:**
```
CONFLICT: 'Lot General Piping Hangers - Purchase Order'
  Existing: 02/08/2027 from 17W before 'Lot General Piping Hangers - Ready for 
            Shipment'. Possible 12W to14W. Originally 30/08/2027 (based on 13W), 
            shifted by 4W due to middle of June to end of August holiday period
  New calc: 05/07/2027 from 7W after 'Lot General Piping Hangers - TSD / BANF RFQ'. 
            Possible 6W to 8W
```

---

## Usage

### Prerequisites
```bash
# Python 3.12+ with virtual environment
pip install pandas openpyxl thefuzz python-Levenshtein
```

### Running the Script
```bash
cd PTS_check
..\\.venv\Scripts\python.exe pts_activity_matcher.py
```

### Expected Output
```
Loading project plan: C:\...\PTS_project_specific.xlsx
Loading general file: C:\...\PTS_General_input_data.xlsx

Project plan shape: (10967, 6)
Important activities: 24
Constraints/relationships: 16
Rightmost used column: Total Float (index 5)

Matching activities (threshold: 90%)...
Matched 24 activities out of 9194 total activities
Loaded 2 holiday period(s)

Calculating requested dates based on dependencies...
Anchor point: Lot 5 - Erection General Piping - Contractor On Site Mobilisation
  Start date: 17/01/2028
  Loaded 2 holiday period(s) for years 2027 to 2029

⚠️  CONFLICT: 'Activity Name'
  Existing: [date] from [path 1]
  New calc: [date] from [path 2]
[... more conflicts if any ...]

Calculated dates for 16 activities
Calculated current lengths for 16 activities

Exporting all 10967 rows with 24 important activities marked
Saving output to: C:\...\PTS_project_specific_process_2025-12-04.xlsx
Added 'Conflicts' sheet with 16 conflict(s)

[OK] Successfully created PTS_project_specific_process_2025-12-04.xlsx
  - Total rows exported: 10967
  - Important activities: 24 (filtered by default)
  - Activities with score < 95%: 2 (highlighted in yellow)
  - Activities with calculated dates: 10967
  - Conflicts detected: 16 (see 'Conflicts' sheet)
  - Applied autofilter (showing important activities), column widths, wrap text, 
    and freeze panes at C2

================================================================================
Processing complete!
================================================================================
```

### Output Files
- **Main file:** `PTS_project_specific_process_YYYY-MM-DD.xlsx`
  - Sheet 1: All activities with calculated dates
  - Sheet 2: Conflicts (if any detected)

---

## Troubleshooting

### No Dates Calculated
**Symptom:** "Calculated dates for 0 activities"
**Causes:**
- Anchor milestone not found in matched activities
- Relationship names don't match any matched activities
- All relationships form closed loops with no anchor connection

**Solution:** Verify activity names in relationships match exactly with general input file.

### Too Many Conflicts
**Symptom:** Conflicts sheet has many entries
**Causes:**
- Inconsistent duration estimates in relationships
- Holiday adjustments affecting different paths differently
- Bidirectional relationships with asymmetric durations

**Solution:** Review relationship table for consistency, adjust durations, or accept one path as authoritative.

### Activities Not Matched
**Symptom:** "Matched 0 activities" or low match count
**Causes:**
- Activity names differ significantly between general file and project plan
- Similarity threshold too strict (90%)
- Typos in activity names

**Solution:** Review activity names in both files, consider lowering threshold, or fix typos.

### Holiday Adjustments Not Applied
**Symptom:** Comments don't show "shifted by XW due to holiday period"
**Causes:**
- Holiday periods don't overlap with execution periods
- Holiday sheet not found or incorrectly formatted
- Activity execution periods outside holiday date ranges

**Solution:** Verify Holidays sheet exists, check date ranges cover relevant years, review holiday period definitions.

### Wrong Dates in Output
**Symptom:** Calculated dates seem incorrect
**Debugging Steps:**
1. Check anchor date is correct in project plan
2. Verify relationship directions (From/To) are correct
3. Review holiday adjustments in comments
4. Check for conflicts that might indicate the issue
5. Manually calculate one path to verify logic

---

## Limitations & Assumptions

### Assumptions
1. **Single anchor:** All dates derived from one mobilisation milestone
2. **Finish dates used:** All activities except anchor use Finish date for calculations
3. **First calculation wins:** When conflicts occur, first calculated date is kept
4. **Linear dependencies:** No parallel paths or alternative routing
5. **Fixed holiday adders:** Each holiday period has one adder applied to all activities

### Limitations
1. **No resource leveling:** Doesn't consider resource availability
2. **No critical path:** Doesn't identify critical path or float
3. **No calendar:** Uses simple week arithmetic, not work calendars
4. **No partial weeks:** All durations are whole weeks
5. **Text-based holidays:** Holiday definitions must match expected text patterns
6. **One-to-one matching:** Each important activity matches one project plan activity

### Known Issues
1. **Conflict resolution:** No automatic resolution, manual review required
2. **Circular dependencies:** Will stop at max iterations (100) without resolution
3. **Multiple anchors:** Cannot handle activities anchored to different milestones
4. **Holiday year range:** Limited to anchor year ± 1 (3 years total)

---

## Maintenance

### Updating Input Files

**To add new activities:**
1. Add row to General Input Data sheet 1 (Activities)
2. Add relationships to sheet 2 (Relationships)
3. Rerun script

**To modify holiday periods:**
1. Edit Holidays sheet in General Input Data
2. Rerun script
3. Review adjustments in comments

**To change anchor:**
1. Modify `mobilisation_name` variable in `calculate_requested_dates()` function
2. Update to use different activity's date
3. Rerun script

### Script Modifications

**Key functions:**
- `process_pts_file()` - Main orchestration
- `calculate_requested_dates()` - Date calculation engine
- `check_holiday_conflict()` - Holiday overlap detection
- `parse_holiday_periods()` - Holiday text parsing
- `ensure_weekday()` - Weekday validation

**Configuration variables:**
- `similarity_threshold` - Fuzzy match threshold (default: 90)
- `max_iterations` - Maximum calculation loops (default: 100)
- `category_columns` - List of category column names

---

## Version History

**Current Version:** 1.0
- Fuzzy matching with 90% threshold
- Bidirectional date calculation
- Holiday period support with execution overlap detection
- Weekday validation
- Conflict detection and reporting
- Excel export with autofilter on important activities
- Current length calculation showing actual vs planned durations
- Swiss date formatting (dd/mm/yyyy)

---

## Contact & Support

For issues, questions, or enhancement requests, please refer to the project repository or contact the development team.

**Script Location:** `C:\Users\szil\Repos\excel_wizadry\PTS_check\pts_activity_matcher.py`

**Input Files Location:** `C:\Users\szil\Repos\excel_wizadry\PTS_check\`
