# Testing Guide for Line List Check Pipe Class

This directory contains comprehensive tests for the `line_list_check_pipe_class.py` script using pytest.

## Test Structure

- **conftest.py**: Contains fixtures and shared test data
- **test_setup_and_utilities.py**: Tests for path setup and utility functions
- **test_data_reading.py**: Tests for Excel file reading functions
- **test_comparison_functions.py**: Tests for data comparison functions
- **test_data_validation.py**: Tests for data validation logic
- **test_excel_functions.py**: Tests for Excel formatting and saving functions
- **test_worksheet_formatting.py**: Tests for worksheet formatting functions
- **test_main_function.py**: Tests for the main function and integration tests

## Running the Tests

### Via Batch File
Simply double-click the `run_tests.bat` file to run all tests with coverage reports.

### Via Command Line
From the `_tests` directory:
```bash
pytest -xvs --cov=.. --cov-report=term --cov-report=html
```

This will:
- Run all tests (-v for verbose output)
- Generate coverage reports for the parent directory (--cov=..)
- Display coverage in the terminal (--cov-report=term)
- Create an HTML coverage report (--cov-report=html)

### Test Only Specific Files
```bash
pytest -xvs test_data_reading.py
```

## Coverage Report

After running tests, a coverage report will be generated in the `htmlcov` directory. 
Open `htmlcov/index.html` in a browser to view the detailed coverage report.

## Test Dependencies

Required Python packages for testing:
- pytest
- pytest-cov
- pandas
- numpy
- xlsxwriter

Install with:
```bash
pip install pytest pytest-cov pandas numpy xlsxwriter
```
