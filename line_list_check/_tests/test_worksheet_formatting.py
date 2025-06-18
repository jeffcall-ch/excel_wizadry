import pytest
import pandas as pd
import numpy as np
from unittest.mock import patch, MagicMock

import line_list_check_pipe_class as lp

class TestWorksheetFormatting:
    """Tests for worksheet formatting functions."""
    
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
        mock_worksheet.write.assert_any_call(3, 1, 'Medium_check: nan', mock_formats['nan_wrap'])
    
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
