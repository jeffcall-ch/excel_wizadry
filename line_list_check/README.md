# Line List Check

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

Run the pipe class checker script:

```powershell
python line_list_check\line_list_check_pipe_class.py
```

This script reads the 'Query' tab from the Excel file `Export 17.06.2025_LS.xlsx` and displays basic information about the data.

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
