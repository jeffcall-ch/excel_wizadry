import os
import sys
import pytest
import pandas as pd
import numpy as np
from unittest.mock import patch, MagicMock

# Import the module being tested
import line_list_check_pipe_class as lp

class TestWorksheetFormatting:
    """Tests for Excel worksheet formatting functions."""
    
    def setup_mocks(self):
        """Helper to set up common mocks for tests."""
        mock_worksheet = MagicMock()
        mock_formats = {
            'center_align': MagicMock(),
            'ok': MagicMock(),
            'nok': MagicMock(),
            'nan': MagicMock(),
            'header_red': MagicMock(),
            'header_yellow': MagicMock(),
            'row_number_red': MagicMock(),
            'row_number_yellow': MagicMock(),
            'ok_wrap': MagicMock(),
            'nok_wrap': MagicMock(),
            'nan_wrap': MagicMock(),
            'header_red_wrap': MagicMock(),
            'header_yellow_wrap': MagicMock(),
            'header_default_wrap': MagicMock()
        }
        return mock_worksheet, mock_formats
    
    def test_format_worksheet(self):
        """Test format_worksheet function applies proper formatting."""
        # Setup mocks
        mock_worksheet, mock_formats = self.setup_mocks()
        
        # Create test dataframe with check columns
        df = pd.DataFrame({
            'Row Number': [1, 2, 3],
            'Column1': ['Value1', 'Value2', 'Value3'],
            'Column1_check': ['OK', 'NOK', 'nan'],
            'Column2': [10, 20, 30],
            'Column2_check': ['OK', 'OK', 'NOK'],
            'Pipe Class status check': ['OK', 'Column1_check: NOK', 'Column2_check: NOK, Column1_check: nan']
        })
        
        # Call function
        lp.format_worksheet(mock_worksheet, df, mock_formats)
        
        # Verify worksheet.set_column was called for every column
        assert mock_worksheet.set_column.call_count >= len(df.columns)
        
        # Verify worksheet.freeze_panes was called
        mock_worksheet.freeze_panes.assert_called_once_with(1, 3)
        
        # Verify worksheet.write was called multiple times
        assert mock_worksheet.write.call_count > 0
        
        # Verify worksheet.autofilter was called
        mock_worksheet.autofilter.assert_called_once_with(0, 0, len(df), len(df.columns) - 1)
    
    def test_format_status_column(self):
        """Test format_status_column function applies proper formatting."""
        # Setup mocks
        mock_worksheet, mock_formats = self.setup_mocks()
        
        # Create test dataframe with status column
        df = pd.DataFrame({
            'Column1': [1, 2, 3],
            'Pipe Class status check': ['OK', 'Error message', 'Medium_check: nan']
        })
        
        # Call function
        lp.format_status_column(mock_worksheet, df, mock_formats, 'Pipe Class status check')
        
        # Verify worksheet.set_column was called for the status column
        mock_worksheet.set_column.assert_called_once_with(1, 1, 35)
        
        # Verify worksheet.write was called for each row plus header
        assert mock_worksheet.write.call_count == 4
        
        # Verify specific formats were used
        mock_worksheet.write.assert_any_call(1, 1, 'OK', mock_formats['ok_wrap'])
        mock_worksheet.write.assert_any_call(2, 1, 'Error message', mock_formats['nok_wrap'])
        # For nan values in the status, the check looks if it contains 'nan', but is still formatted as nok_wrap if not exclusively 'nan'
        mock_worksheet.write.assert_any_call(3, 1, 'Medium_check: nan', mock_formats['nok_wrap'])
    
    def test_adjust_columns_and_add_filter(self):
        """Test adjust_columns_and_add_filter function."""
        # Setup mocks
        mock_worksheet = MagicMock()
        
        # Create test dataframe
        df = pd.DataFrame({
            'ShortCol': ['A', 'B', 'C'],
            'LongColumnName': ['Value1', 'Value2', 'Value3'],
            'VeryLongColumnWithLongValues': ['This is a long value', 'Another long value', 'Yet another long value']
        })
        
        # Call function
        lp.adjust_columns_and_add_filter(mock_worksheet, df)
        
        # Verify worksheet.set_column was called for each column
        assert mock_worksheet.set_column.call_count == 3
        
        # Verify worksheet.set_row was called for each row plus header
        assert mock_worksheet.set_row.call_count == 4
        
        # Verify worksheet.autofilter was called
        mock_worksheet.autofilter.assert_called_once_with(0, 0, 3, 2)


    
    @patch('line_list_check_pipe_class.os.makedirs')
    @patch('line_list_check_pipe_class.pd.ExcelWriter')
    def test_generate_pipe_class_summary_with_exception(self, mock_excel_writer, mock_makedirs):
        """Test generate_pipe_class_summary function with an exception."""
        # Mock data
        mock_line_list_df = pd.DataFrame({
            'Pipe Class': ['A1'],
            'Medium': ['Water'],
            'Medium_check': ['OK'],
            'PS [bar(g)]': [10],
            'PS [bar(g)]_check': ['OK'],
            'TS [°C]': [80],
            'TS [°C]_check': ['OK'],
            'DN': ['DN50'],
            'DN_check': ['OK'],
            'PN': ['PN16'],
            'PN_check': ['OK'],
            'EN No. Material': ['1.4301'],
            'EN No. Material_check': ['OK']
        })
        mock_pipe_class_dict = {'A1': {'Medium': 'Water'}}
        
        # Configure mocks
        mock_makedirs.return_value = None
        
        # Configure mock to raise an exception
        mock_excel_writer.side_effect = Exception("Test exception")
        
        # Create test file path with directory
        test_output_path = os.path.join('test_dir', 'output.xlsx')
        
        # Call the function
        result = lp.generate_pipe_class_summary(mock_line_list_df, mock_pipe_class_dict, test_output_path)
        
        # Verify the result indicates failure
        assert result is False


    @patch('line_list_check_pipe_class.os.makedirs')
    @patch('line_list_check_pipe_class.pd.ExcelWriter')
    def test_pipe_class_summary_file_error(self, mock_excel_writer, mock_makedirs):
        """Test error handling in generate_pipe_class_summary."""
        # Create test data
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
