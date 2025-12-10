# ADI Project - Weld Count Report

**Report Date:** December 9, 2025  
**Project:** ADI  
**Script:** weld_counter.py

---

## Executive Summary

**Total Welds Required: 17,463**

---

## Calculation Formula

```
Component Welds + BWD Branch Ends - Touching Pairs - Components at BWD Ends
   15,779      +     1,809       -      42       -         83          = 17,463
```

---

## Detailed Breakdown

### 1. Component Welds: 15,779

**Welded Components:** 11,098 (from 28,907 total components)

Components classified as "welded" based on Excel specification file:
- P1 CONN = 'BWD' (Port 1 has Butt Weld connection)
- P2 CONN = 'BWD' (Port 2 has Butt Weld connection)
- TYPE = 'OLET' (Olet-type component)

**Weld Count Logic:**
- P1 CONN = BWD only → 1 weld
- P2 CONN = BWD only → 1 weld
- Both P1 and P2 = BWD → 2 welds
- TYPE = OLET → 1 weld

---

### 2. BWD Branch Ends: +1,809

**Understanding HCON and TCON:**
- HCON = Head Connection Type (how the branch HEAD connects)
- TCON = Tail Connection Type (how the branch TAIL connects)
- BWD = Butt Weld (requires welding)

**Branch Ends Requiring Welds:**
- HCON = BWD: 1,374 branch heads
- TCON = BWD: 435 branch tails
- **Total:** 1,809 BWD branch ends

These branch ends connect to:
- Equipment nozzles
- Other pipe branches (173 branch-to-branch connections)
- Main piping systems
- Termination points

**Note:** The 1,809 total ALREADY INCLUDES the 173 branch-to-branch connections (where one branch's TCON=BWD connects to another's HCON=BWD). These are NOT added separately.

---

### 3. Touching Component Pairs: -42

When two components are installed back-to-back, they share a single weld. Since this weld is counted in both components' weld counts, we subtract to avoid duplication.

**Detection Method:**
- Calculate 3D distance between consecutive components
- Compare to expected distance based on component geometry
- Components are "touching" if: Distance ≤ Expected Distance × 110%

**Component Length Formulas:**
- ELBO (Elbow): FORM × PBOR
- TEE: 0.90 × PBOR
- FLAN (Flange): 0.40 × PBOR
- VALV (Valve): 0.50 × PBOR (or PBOR1 if PBOR unavailable)
- REDU (Reducer): 1.15 × PBOR1
- CAP: 0.0

**Results:**
- Total component pairs analyzed: 676
- Touching pairs identified: 42

---

### 4. Components at BWD Branch Ends: -83

When a component is placed directly at a BWD branch end, its weld is counted in BOTH the component welds AND the branch end. We subtract to avoid double-counting.

**Detection:**
- Components at HCON=BWD positions: 43
- Components at TCON=BWD positions: 40
- **Total to subtract:** 83

---

## Summary Statistics

| Category | Count |
|----------|------:|
| **Total Components** | 28,907 |
| Found in Excel | 20,353 |
| Not Found in Excel | 8,554 |
| **Welded Components** | 11,098 |
| **Component Welds** | 15,779 |
| **Total Branches** | 1,696 |
| Connected Branches | 173 |
| HCON = BWD | 1,374 |
| TCON = BWD | 435 |
| **BWD Branch Ends** | 1,809 |
| **Component Pairs Analyzed** | 676 |
| Touching Pairs | 42 |
| Near Pairs | 18 |
| Separated Pairs | 616 |
| **Components at Branch Ends** | 252 |
| At HEADS | 118 (43 at HCON=BWD) |
| At TAILS | 134 (40 at TCON=BWD) |
| **Components at BWD Ends** | 83 |
| **BWD Connections WITH Component** | 299 |
| **BWD Connections WITHOUT Component** | 1,510 |

---

## Final Calculation

```
Component Welds:                    15,779
+ BWD Branch Ends:                  +1,809
- Touching Component Pairs:         -   42
- Components at BWD Ends:           -   83
──────────────────────────────────────────
TOTAL WELDS:                        17,463
```

---

## Data Sources

**Input Files:**
1. `0AUX-P-INN_8_Dec_2025.TXT` (7,929 components, 356 branches)
2. `1AUX-P-INN_8_Dec_2025.TXT` (10,514 components, 670 branches)
3. `2AUX-P-INN_8_Dec_2025.TXT` (10,464 components, 670 branches)
4. `ADI_PIPE_SPECS_ALL.xlsx` (Component specifications)

**Output Files Generated:**
1. `components_with_welds.csv` (28,907 rows)
2. `branch_connections.csv` (1,696 rows)
3. `connected_branches.csv` (173 rows)
4. `component_adjacency.csv` (676 rows)
5. `components_at_branch_ends.csv` (951 rows)
6. `bwd_connections_report.csv` (1,809 rows)

---

## Quality Assurance

✓ **No double-counting of branch-to-branch connections**
  - 173 connected branches are already part of the 1,809 BWD branch ends
  
✓ **No double-counting of component-to-component welds**
  - 42 touching pairs share welds already counted in both components

✓ **No double-counting of components at BWD branch ends**
  - 83 components subtracted to avoid counting their welds twice

✓ **Accurate geometric calculations**
  - Component lengths use engineering formulas (FORM × PBOR for elbows, etc.)

✓ **Proper filtering of non-welded connections**
  - Only BWD (Butt Weld) connections counted

---

*End of Report*
