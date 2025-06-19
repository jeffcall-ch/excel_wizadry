import os
import sys
import pytest
import pandas as pd
import numpy as np
from unittest.mock import patch, MagicMock, mock_open
import xlsxwriter

# Add parent directory to path to import the module being tested
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Import the module being tested
import line_list_check_pipe_class as lp

# Fixtures and test data
@pytest.fixture
def sample_dirs_files():
    """Fixture for sample directories and files paths."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(current_dir)
    
    dirs = {
        'input_dir': os.path.join(parent_dir, "input_file"),
        'pipe_class_dir': os.path.join(parent_dir, "pipe_class_summary_file"),
        'output_dir': os.path.join(parent_dir, "output")
    }
    
    files = {
        'line_list_file': os.path.join(dirs['input_dir'], "Export 17.06.2025_LS.xlsx"),
        'pipe_class_file': os.path.join(dirs['pipe_class_dir'], "PIPE CLASS SUMMARY_LS_06.06.2025_updated_column_names.xlsx"),
        'output_file': os.path.join(dirs['output_dir'], "Line_List_with_Matches.xlsx"),
        'summary_file': os.path.join(dirs['output_dir'], "Pipe_Class_Summary.xlsx")
    }
    
    return dirs, files

@pytest.fixture
def sample_line_list_df():
    """Fixture for sample line list dataframe."""
    data = {
        "F1": ["Line No.", "L-001", "L-002", "L-003"],
        "F2": ["KKS", "30ABC001", "30ABC002", "30ABC003"],
        "Medium": ["Medium", "Water", "Steam", "Gas"],
        "PS [bar(g)]": ["PS [bar(g)]", "10", "15", "20"],
        "TS [°C]": ["TS [°C]", "50", "100", "150"],
        "DN": ["DN", "DN 25", "DN 50", "DN 100"],
        "PN": ["PN", "PN 16", "PN 25", "PN 40"],
        "EN No. Material": ["EN No. Material", "1.0425", "1.0460", "1.4571"],
        "Pipe Class": ["Pipe Class", "A1", "B2", "C3"],
        "ComosSystemInfo": ["ComosSystemInfo", "Info1", "Info2", "Info3"]
    }
    
    return pd.DataFrame(data)

@pytest.fixture
def sample_pipe_class_df():
    """Fixture for sample pipe class dataframe."""
    data = {
        "Pipe Class": ["A1", "B2", "C3", "D4"],
        "Medium": ["Water", "Steam", "Gas, Air", "Oil"],
        "PN": [16, 25, 40, 63],
        "Min temperature (°C)": [0, 0, -10, -20],
        "Max temperature (°C)": [80, 200, 180, 120],
        "Diameter from [DN, NPS]": [15, 15, 15, 15],
        "Diameter to [DN, NPS]": [300, 400, 500, 600],
        "EN No. Material": ["1.0425", "1.0460", "1.4571", "1.7335"]
    }
    
    return pd.DataFrame(data)

@pytest.fixture
def sample_pipe_class_dict(sample_pipe_class_df):
    """Fixture for sample pipe class dictionary."""
    pipe_class_dict = {}
    for index, row in sample_pipe_class_df.iterrows():
        pipe_class = row['Pipe Class']
        pipe_class_dict[pipe_class] = {
            column: row[column] for column in sample_pipe_class_df.columns
        }
    
    return pipe_class_dict

@pytest.fixture
def sample_validated_df():
    """Fixture for sample validated dataframe with check columns."""
    data = {
        "Row Number": [1, 2, 3, 4, 5],
        "Line No.": ["L-001", "L-002", "L-003", "L-004", "L-005"],
        "KKS": ["ABC123", "DEF456", "GHI789", "JKL012", "MNO345"],
        "Medium": ["Water", "Water", "Steam", "Gas", "Oil"],
        "Medium_check": ["OK", "OK", "OK", "OK", "Medium is missing from the pipe class"],
        "PS [bar(g)]": [10, 10, 20, 30, 5],
        "PS [bar(g)]_check": ["OK", "OK", "NOK", "NOK", "OK"],
        "TS [°C]": [50, 50, 150, 200, 25],
        "TS [°C]_check": ["OK", "OK", "OK", "NOK", "OK"],
        "DN": ["DN 25", "DN 25", "DN 50", "DN 80", "DN 15"],
        "DN_check": ["OK", "OK", "OK", "OK", "OK"],
        "PN": ["PN 16", "PN 16", "PN 25", "PN 40", "PN 10"],
        "PN_check": ["OK", "OK", "OK", "OK", "OK"],
        "EN No. Material": ["1.0425", "1.0425", "1.0460", "1.4571", "1.0425"],
        "EN No. Material_check": ["OK", "OK", "OK", "OK", "OK"],
        "Pipe Class": ["A1", "A1", "B2", "C3", "X1"],
        "Pipe Class_check": ["OK", "OK", "OK", "OK", "Pipe class not found in summary"],
        "Pipe Class status check": [
            "OK", 
            "OK", 
            "PS [bar(g)]_check: NOK", 
            "PS [bar(g)]_check: NOK, TS [°C]_check: NOK", 
            "Pipe Class_check: Pipe class not found in summary, Medium_check: Medium is missing from the pipe class"
        ]
    }
    
    return pd.DataFrame(data)

@pytest.fixture
def mock_formats():
    """Fixture for mock Excel formats."""
    formats = {
        'center_align': MagicMock(),
        'ok': MagicMock(),
        'nok': MagicMock(),
        'nan': MagicMock(),
        'header_red': MagicMock(),
        'header_yellow': MagicMock(),
        'row_number_red': MagicMock(),
        'row_number_yellow': MagicMock(),
        'header_default': MagicMock(),
        'header_default_wrap': MagicMock(),
        'header_red_wrap': MagicMock(), 
        'header_yellow_wrap': MagicMock(),
        'ok_wrap': MagicMock(),
        'nok_wrap': MagicMock(),
        'nan_wrap': MagicMock()
    }
    return formats
