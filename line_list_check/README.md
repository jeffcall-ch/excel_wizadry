# Line List Check

## Overview

The Line List Check is an Excel processing utility designed for industrial piping engineers and designers. It validates pipe class specifications in line lists against reference data to ensure compliance with engineering standards. This tool helps identify and correct discrepancies in pipe specifications, ensuring that all piping components meet the required standards and specifications.

## Functionality

The script performs the following operations:

1. **Data Import and Validation**:
   - Reads a Line List Excel file containing pipe specifications
   - Imports reference Pipe Class Summary data with standard specifications
   - Validates each line entry against the reference standards

2. **Validation Checks**:
   - Pipe Class existence in reference data
   - Material compatibility and compliance
   - Pressure and temperature ratings
   - Nominal diameter (DN) and pressure nominal (PN) specifications
   - Medium compatibility with pipe class

3. **Results and Reporting**:
   - Generates a detailed pipe class summary report showing compliance status
   - Highlights compliant vs non-compliant values using color coding
   - Creates a modified Line List with validation results for each entry
   - Provides statistical overview of compliance issues

## Setup

### Activating Virtual Environment

To activate the virtual environment for this project:

```powershell
# Navigate to the project directory
cd c:\Users\szil\Repos\excel_wizadry

# Activate the virtual environment
.venv\Scripts\Activate.ps1
```

You should see `(.venv)` prefix in your terminal prompt indicating the virtual environment is active.

### Deactivating Virtual Environment

To deactivate the virtual environment:

```powershell
deactivate
```

## Usage

### Main Script

Run the pipe class checker script:

```powershell
python line_list_check\line_list_check_pipe_class.py
```

This script:
1. Reads the 'Query' tab from the Excel file `Export 17.06.2025_LS.xlsx`
2. Loads reference data from `PIPE CLASS SUMMARY_LS_06.06.2025_updated_column_names.xlsx`
3. Performs validation checks on all pipe specifications
4. Generates two output files in the `output` directory:
   - `Line_List_with_Matches.xlsx`: Original line list with validation results
   - `Pipe_Class_Summary.xlsx`: Summary of all pipe classes with compliance status

### Refactored Script

A refactored version of the script is also available with improved readability and organization:

```powershell
python line_list_check\line_list_check_pipe_class_refactored.py
```

The refactored script performs the same operations but with:

- Better function organization
- Clearer error handling
- Improved code readability
- Simplified logic and structure

Use the refactored version for better maintainability and easier understanding of the code.

## Testing

The test suite for this project is located in the `_tests` folder. See the README.md file in that folder for detailed information about running tests and test coverage.
