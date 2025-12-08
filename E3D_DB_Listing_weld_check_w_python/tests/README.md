# Test Suite for weld_counter.py

## Overview

Comprehensive test suite achieving **82% code coverage** with **72 passing tests** covering all major functions and edge cases, including extensive pipe length calculation validation.

## Coverage Summary

```
Name              Stmts   Miss  Cover   Missing
-----------------------------------------------
weld_counter.py     332     60    82%   514-607
-----------------------------------------------
```

**Note:** The uncovered lines (514-607) are in the `if __name__ == '__main__':` block, which is the main execution script. This is standard practice - main blocks are typically not unit tested as they orchestrate other tested functions.

## Test Structure

### Test Classes

1. **TestExtractBranchPositions** (10 tests)
   - Tests branch position extraction (HPOS/TPOS coordinates)
   - Covers happy path, edge cases, and real data

2. **TestFindConnectedBranches** (11 tests)
   - Tests axis-combination matching algorithm
   - All three match types: XY_tight_Z_loose, XZ_tight_Y_loose, YZ_tight_X_loose
   - Accuracy calculation and tolerance handling

3. **TestExtractBranchConnections** (8 tests)
   - Tests branch connection data extraction
   - Component tracking (first/last)
   - Pipe length calculation
   - ATTACHMENT exclusion logic

4. **TestExtractComponentsFromBranches** (6 tests)
   - Tests component extraction with SPRE references
   - Excluded component types
   - Multiple branches and pipes

5. **TestLookupAndMergeWithExcel** (6 tests)
   - Tests Excel cross-reference functionality
   - Weld detection (BWD connections and OLET type)
   - NaN value handling

6. **TestSaveToCsv** (3 tests)
   - Tests CSV output generation
   - Special character handling
   - Empty data handling

7. **TestIntegrationRealData** (3 tests)
   - End-to-end tests with real TBY data files
   - Validates specific known branches
   - Full pipeline execution

8. **TestPipeLengthCalculations** (14 tests) ⭐ NEW
   - **Single component scenarios**: 1 component with pipe segments
   - **Multiple components**: 2, 5 components in various configurations
   - **Edge cases at boundaries**: Components at exact HPOS, TPOS, or both
   - **Zero scenarios**: No components, only excluded components
   - **ATTACHMENT handling**: Between valid components (should be ignored)
   - **3D calculations**: Diagonal pipe lengths in all dimensions
   - **Extreme values**: Very short (sub-mm), very long (100m+)
   - **Negative coordinates**: Proper distance calculation
   - **Multiple branches**: Different pipe lengths in same file
   - **Real data validation**: Known branches with expected values

9. **TestEdgeCases** (11 tests)
   - Nonexistent files
   - Malformed coordinates
   - Unicode characters
   - Very large coordinate values
   - Duplicate branch IDs
   - Complex multi-pipe files
   - All match types in one test
   - Components at exact boundary positions

## Running Tests

### Run all tests with coverage report:
```bash
pytest tests/test_weld_counter.py -v --cov=weld_counter --cov-report=term-missing --cov-report=html
```

### Run specific test class:
```bash
pytest tests/test_weld_counter.py::TestExtractBranchPositions -v
```

### Run specific test:
```bash
pytest tests/test_weld_counter.py::TestExtractBranchPositions::test_happy_path_single_branch -v
```

### View HTML coverage report:
After running tests, open `htmlcov/index.html` in a browser for detailed line-by-line coverage.

## Test Data

### Synthetic Data
Tests use `tmp_path` fixtures to create temporary test files with controlled data for:
- Single branches
- Multiple branches per pipe
- Multiple pipes per file
- Various coordinate patterns
- Edge cases (negative, decimal, very large values)

### Real Data
Tests validate against actual TBY data files:
- `TBY-0AUX-P.txt` (483 branches, 7,633 components)
- `TBY-1AUX-P.txt` (528 branches, 5,903 components)
- `TBY_all_pspecs_wure_macro_08.12.2025.xlsx` (4,938 rows)

### Expected Results (Real Data)
- Total components: 13,536
- Total branches: 1,005
- Total branch positions: 1,011
- Connected branches: 174
- Welded components: 1,425

## Test Categories

### Happy Path Tests ✓
- Single branch extraction
- Perfect coordinate matches
- Standard component detection
- Excel lookup with valid data
- CSV output generation

### Edge Case Tests ✓
- Empty files
- Missing position data
- Malformed coordinates
- Unicode characters
- Very large coordinate values
- Duplicate branch IDs
- Components at exact HPOS/TPOS
- Branches with only excluded components

### Breaking Tests ✓
- Nonexistent files (FileNotFoundError expected)
- Branches too far apart (no matches)
- Components not found in Excel
- Empty data for CSV save
- Missing positions in branch matching

### Integration Tests ✓
- Full pipeline with real data
- Multiple files processed together
- All functions working in concert
- Specific known branch validation

## Coverage by Function

| Function | Coverage | Tests | Key Scenarios |
|----------|----------|-------|---------------|
| `extract_branch_positions` | 100% | 10 | Single/multiple branches, real data, edge cases |
| `find_connected_branches` | 100% | 11 | All match types, accuracy, tolerances, real data |
| `extract_branch_connections` | 100% | 8 | Components, pipe lengths, ATTACHMENT exclusion |
| `extract_components_from_branches` | 100% | 6 | Multiple types, SPRE extraction, exclusions |
| `lookup_and_merge_with_excel` | 100% | 6 | Weld detection, NaN handling, missing data |
| `save_to_csv` | 100% | 3 | Standard save, empty data, special chars |
| Main block | 0% | N/A | Not tested (standard practice) |

## Pipe Length Calculation Test Coverage ⭐

The test suite includes **14 comprehensive tests** specifically for pipe length calculations covering:

### Component Count Scenarios
- ✅ **0 components**: Zero pipe length (empty branch)
- ✅ **1 component**: Pipe on head and tail
- ✅ **2 components**: Pipe segments between components
- ✅ **5 components**: Complex multi-component branch
- ✅ **Multiple branches**: Different configurations in same file

### Boundary Edge Cases
- ✅ **Component at HPOS**: Zero head pipe length
- ✅ **Component at TPOS**: Zero tail pipe length
- ✅ **Components at both boundaries**: Zero head and tail lengths
- ✅ **Only excluded components**: ATTACHMENT only → zero lengths

### Mathematical Edge Cases
- ✅ **3D diagonal distances**: All three axes changing
- ✅ **Very short segments**: Sub-millimeter precision (0.87mm)
- ✅ **Very long segments**: 100+ meters (100,000mm)
- ✅ **Negative coordinates**: Proper absolute distance calculation
- ✅ **ATTACHMENT between components**: Correctly ignored for first/last tracking

### Real Data Validation
- ✅ **Known branch 0QCA30BR220/B1**: Validates 569.97mm head, 0.0mm tail
- ✅ **Statistical validation**: Counts branches with non-zero pipe lengths

### Test Examples

**Example 1: Single component with pipe on both ends**
```
HPOS (1000, 2000, 3000) → ELBOW (2000, 3000, 4000) → TPOS (3000, 4000, 5000)
Head pipe = √(1000² + 1000² + 1000²) = 1732.05mm ✓
Tail pipe = √(1000² + 1000² + 1000²) = 1732.05mm ✓
```

**Example 2: Component at exact TPOS**
```
HPOS (1000, 2000, 3000) → FLANGE (3000, 4000, 5000) → ELBOW@TPOS (5000, 6000, 7000)
Head pipe = 3464.10mm ✓
Tail pipe = 0.0mm ✓
```

**Example 3: Five components**
```
HPOS (0,0,0) → FLANGE (500,0,0) → ELBOW (2000,0,0) → REDUCER (5000,0,0) 
→ TEE (7000,0,0) → VALVE (9500,0,0) → TPOS (10000,0,0)
Head pipe = 500mm ✓
Tail pipe = 500mm ✓
```

## Key Test Validations

### Coordinate Matching Algorithm
- **XY tight, Z loose**: 2 axes within 5mm, Z axis within 150mm
- **XZ tight, Y loose**: X and Z within 5mm, Y within 150mm  
- **YZ tight, X loose**: Y and Z within 5mm, X within 150mm
- Accuracy percentage: `100 * (1 - max_tight_offset / tolerance_tight)`

### Weld Detection Criteria
- `P1 CONN == 'BWD'` → Welded
- `P2 CONN == 'BWD'` → Welded
- `TYPE == 'OLET'` → Welded
- All other combinations → Not welded

### Component Exclusion
From first/last component tracking (but included in extraction):
- ATTACHMENT, BRANCH, PIPE, ZONE, SITE, STRUCTURE, SUBSTRUCTURE, CYLINDER, CTORUS

### Pipe Length Calculation
- **Head pipe length**: 3D distance from HPOS to first component POS
- **Tail pipe length**: 3D distance from last component POS to TPOS
- Formula: `√((x₁-x₂)² + (y₁-y₂)² + (z₁-z₂)²)`

## Known Limitations

1. **Main Block Not Tested**: The `if __name__ == '__main__':` block (18% of code) is not unit tested, which is standard practice
2. **Excel Column Requirements**: Function requires all expected columns (P1 CONN, P2 CONN, TYPE) to be present
3. **Missing Branches**: Real data shows 6 branches not captured (1,005 vs 1,011 expected) - likely parsing edge cases

## Test Maintenance

When adding new features:
1. Add corresponding test class or expand existing
2. Include both happy path and edge case tests
3. Run full test suite to ensure no regressions
4. Aim to maintain >80% coverage for all new functions
5. Add real data validation tests for critical functionality

## Dependencies

```python
pytest>=7.4.0
pytest-cov>=4.1.0
pandas
openpyxl
```

## Test Execution Time

Average execution time: **~24 seconds** (includes real data processing and 72 comprehensive tests)

## Continuous Integration

For CI/CD pipelines, use:
```bash
pytest tests/test_weld_counter.py --cov=weld_counter --cov-report=xml --cov-fail-under=80
```

This will fail the build if coverage drops below 80%.
