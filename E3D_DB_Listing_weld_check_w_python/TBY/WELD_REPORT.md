# TBY Weld Count Report
**Generated:** December 12, 2025  
**Project:** TBY Auxiliary Piping

---

## Executive Summary

**Total Estimated Welds: 2,147**

This report analyzes weld requirements for the TBY project based on E3D database listings and pipe specifications.

---

## Input Files Processed

### E3D Database Listings
- `TBY-0AUX-P.txt` - 7,633 components, 482 branches
- `TBY-1AUX-P.txt` - 5,903 components, 523 branches

### Specifications
- `TBY_all_pspecs_wure_macro_08.12.2025.xlsx` - Pipe component specifications

### Total Data
- **Components:** 13,536
- **Branches:** 1,005
- **Branch Positions:** 1,011

---

## Component Analysis

### Component Lookup Results
| Category | Count | Percentage |
|----------|-------|------------|
| Found in Excel | 9,533 | 70.4% |
| Not found in Excel | 4,003 | 29.6% |
| **Total Components** | **13,536** | **100%** |

### Welded Components
| Category | Count |
|----------|-------|
| Components with BWD connections or OLET type | 1,425 |
| Total welds from components | 2,234 |

**Note:** Some components have multiple weld connections (P1 CONN and/or P2 CONN = BWD)

---

## Branch Connection Analysis

### BWD Connection Summary
| Connection Type | Count |
|-----------------|-------|
| HCON BWD (Head connections) | 148 |
| TCON BWD (Tail connections) | 91 |
| **Total BWD connections** | **239** |

### Branch-to-Branch Connections
**Connected Branches:** 26 (coordinate-based detection)

#### Match Type Distribution
| Match Type | Count | Description |
|------------|-------|-------------|
| XY tight, Z loose | 24 | X,Y ≤5mm; Z ≤150mm |
| YZ tight, X loose | 2 | Y,Z ≤5mm; X ≤150mm |
| XZ tight, Y loose | 0 | X,Z ≤5mm; Y ≤150mm |

*Note: These 26 branch connections are already included in the 239 BWD connection count above*

---

## Components at Branch Ends

### Overall Detection
| Location | Total | At BWD Connections |
|----------|-------|-------------------|
| At branch HEADS | 132 | 35 |
| At branch TAILS | 135 | 32 |
| **Total at branch ends** | **267** | **67** |

### BWD Connection Coverage
| Category | Count | Percentage |
|----------|-------|------------|
| BWD connections WITH component at end | 77 | 32.2% |
| BWD connections WITHOUT component at end | 162 | 67.8% |
| **Total BWD connections** | **239** | **100%** |

**Analysis:** 67.8% of BWD connections do not have a welded component directly at the branch end, indicating field welds or connections to other branches.

---

## Component Adjacency Analysis

**Total Component Pairs Analyzed:** 862

### Relationship Classification
| Relationship | Count | Percentage | Description |
|--------------|-------|------------|-------------|
| **Touching** | 259 | 30.0% | Within PBOR-based threshold (components likely welded together) |
| Near | 58 | 6.7% | Close but not touching |
| Separated | 545 | 63.2% | Beyond threshold (separate components) |

**Touching Threshold Logic:**
- Based on component type and PBOR (Pipe Bore) values
- Accounts for component geometry and typical tolerances
- Excludes OLET types and FLANGE-FLANGE pairs from weld count

---

## Final Weld Count Calculation

### Calculation Breakdown

```
Base Welds (from components):                    2,234
  + BWD branch ends (require welds):              +239
  - Touching component pairs (already counted):   -259
  - Components at BWD ends (already counted):      -67
  ─────────────────────────────────────────────
  = TOTAL ESTIMATED WELDS:                       2,147
```

### Explanation

1. **Component Welds (2,234):**
   - Components with P1 CONN = BWD: +1 weld each
   - Components with P2 CONN = BWD: +1 weld each
   - Components with TYPE = OLET: +1 weld each
   - Some components have multiple connections

2. **BWD Branch Ends (+239):**
   - Each BWD connection point requires a weld
   - Includes branch-to-branch connections (26)
   - Includes branch-to-equipment connections

3. **Touching Component Pairs (-259):**
   - Adjacent components that are welded together
   - Already counted in component welds, so deducted to avoid double-counting
   - Excludes OLETs and FLANGE-FLANGE pairs

4. **Components at BWD Ends (-67):**
   - Welded components located directly at BWD connection points
   - Already counted in component welds, so deducted to avoid double-counting

---

## Output Files Generated

All output files are located in the `TBY` folder:

| File Name | Description | Records |
|-----------|-------------|---------|
| `components_with_welds.csv` | All components with weld information | 13,536 |
| `branch_connections.csv` | Branch connection details | 1,005 |
| `connected_branches.csv` | Branch-to-branch connections | 26 |
| `component_adjacency.csv` | Adjacent component pairs | 862 |
| `components_at_branch_ends.csv` | Components at branch ends | 354 |
| `bwd_connections_report.csv` | BWD connection analysis | 239 |

---

## Key Findings

### 1. Component Coverage
- 70.4% of components found in pipe specifications
- 29.6% not found (may require manual verification)

### 2. Weld Distribution
- **Component-to-Component Welds:** ~2,234 (from component connections)
- **Branch-to-Branch Welds:** ~26 (coordinate-based detection)
- **Field Welds:** ~162 (BWD connections without components)

### 3. Quality Indicators
- **Touching Components:** 259 pairs detected (30% of analyzed pairs)
- **Branch End Coverage:** 67 components at 239 BWD points (28%)
- **Branch Connections:** 26 detected via coordinate matching

### 4. Missing Data
- 4,003 components not found in Excel specification
- 162 BWD connections without identified components at ends
- Requires further investigation for complete weld count

---

## Recommendations

1. **Verify Missing Components:**
   - Review the 4,003 components not found in specifications
   - Update pipe specification database if necessary

2. **Field Weld Verification:**
   - Review 162 BWD connections without components
   - Confirm these are legitimate field welds or connections

3. **Branch Connection Review:**
   - Validate the 26 detected branch-to-branch connections
   - Check for any missing connections that don't match coordinate criteria

4. **Specification Updates:**
   - Ensure all component types are properly classified in Excel
   - Verify PBOR, PBOR1, and FORM values for accurate weld counting

---

## Notes

- **Weld Count Methodology:** Based on BWD (Butt Weld) connection indicators and OLET components
- **Component Types Included:** ELBOW, TEE, REDUCER, FLANGE, VALVE, CAP, and others
- **Exclusions:** WELD type components are excluded from weld counting
- **Coordinate Matching Tolerance:** 5mm tight (2 axes), 150mm loose (1 axis)
- **Component Touching Tolerance:** PBOR-based with 10-50% buffer

---

*Report generated by weld_counter.py analysis tool*
