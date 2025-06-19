import os
import sys
import pytest
import pandas as pd
import numpy as np
from unittest.mock import patch, MagicMock

# Import the module being tested
import line_list_check_pipe_class as lp

class TestPipeClassSummary:
    """Tests for pipe class summary generation."""
    
    @patch('line_list_check_pipe_class.os.makedirs')
    @patch('line_list_check_pipe_class.pd.ExcelWriter')
    def test_pipe_class_summary_file_error(self, mock_excel_writer, mock_makedirs):
        """Test error handling in generate_pipe_class_summary."""        # Create test data
        mock_line_list_df = pd.DataFrame({
            'Pipe Class': ['A1', 'B2'],
            'Pipe Class_check': ['OK', 'OK'],
            'Medium': ['Water', 'Steam'],
            'Medium_check': ['OK', 'OK'],
            'PS [bar(g)]': [10, 20],
            'PS [bar(g)]_check': ['OK', 'OK'],
            'TS [°C]': [80, 200],
            'TS [°C]_check': ['OK', 'OK'],
            'DN': ['DN50', 'DN100'],
            'DN_check': ['OK', 'OK'],
            'PN': ['PN16', 'PN25'],
            'PN_check': ['OK', 'OK'],
            'EN No. Material': ['1.4301', '1.0577'],
            'EN No. Material_check': ['OK', 'OK']
        })
        
        mock_pipe_class_dict = {
            'A1': {'Pipe Class': 'A1'},
            'B2': {'Pipe Class': 'B2'}
        }
        
        # Setup the mock to raise an exception
        mock_makedirs.return_value = None
        mock_writer = MagicMock()
        mock_excel_writer.return_value.__enter__.side_effect = Exception("Test error")
        
        # Test file path
        test_output_path = os.path.join('test_output', 'summary.xlsx')
        
        # Call function and capture the expected exception
        result = lp.generate_pipe_class_summary(mock_line_list_df, mock_pipe_class_dict, test_output_path)
        
        # Verify the function returned False due to the error
        assert result is False
