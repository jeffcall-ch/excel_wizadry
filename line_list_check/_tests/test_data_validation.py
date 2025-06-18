import pytest
import pandas as pd
import numpy as np
from unittest.mock import patch, MagicMock

import line_list_check_pipe_class as lp

class TestDataValidation:
    """Tests for data validation and status summary functions."""
    
    def test_validate_pipe_data(self, sample_pipe_class_dict):
        """Test validate_pipe_data function with different scenarios."""
        # Create a test dataframe
        df = pd.DataFrame({
            'Pipe Class': ['A1', 'B2', 'X1', ''],
            'Medium': ['Water', 'Steam', 'Oil', 'Water'],
            'PS [bar(g)]': [10, 30, 15, 10],
            'TS [°C]': [50, 150, 200, 50],
            'DN': ['DN 25', 'DN 50', 'DN 20', 'DN 25'],
            'PN': ['PN 16', 'PN 16', 'PN 25', 'PN 16'],
            'EN No. Material': ['1.0425', '1.0500', '1.0425', '1.0425']
        })
        
        # Add check columns
        df = lp.add_check_columns(df)
        
        # Validate data
        result_df = lp.validate_pipe_data(df, sample_pipe_class_dict)
        
        # Check results for row 0 (A1 - all should be OK)
        assert result_df.at[0, 'Pipe Class_check'] == 'OK'
        assert result_df.at[0, 'Medium_check'] == 'OK'
        assert result_df.at[0, 'PS [bar(g)]_check'] == 'OK'
        assert result_df.at[0, 'TS [°C]_check'] == 'OK'
        assert result_df.at[0, 'DN_check'] == 'OK'
        assert result_df.at[0, 'PN_check'] == 'OK'
        assert result_df.at[0, 'EN No. Material_check'] == 'OK'
        
        # Check results for row 1 (B2 - pressure and material should be NOK)
        assert result_df.at[1, 'Pipe Class_check'] == 'OK'
        assert result_df.at[1, 'Medium_check'] == 'OK'
        assert result_df.at[1, 'PS [bar(g)]_check'] == 'NOK'  # 30 > 25
        assert result_df.at[1, 'TS [°C]_check'] == 'OK'
        assert result_df.at[1, 'DN_check'] == 'OK'
        assert result_df.at[1, 'PN_check'] == 'NOK'  # PN 16 != 25
        assert result_df.at[1, 'EN No. Material_check'] == 'NOK'  # Different material
        
        # Check results for row 2 (X1 - pipe class not found)
        assert result_df.at[2, 'Pipe Class_check'] == 'Pipe class not found in summary'
        
        # Check results for row 3 (empty pipe class)
        assert result_df.at[3, 'Pipe Class_check'] == 'No pipe class assigned'
    
    def test_add_status_summary(self):
        """Test add_status_summary function that consolidates check results."""
        # Create a test dataframe with check columns
        df = pd.DataFrame({
            'Medium': ['Water', 'Steam', 'Oil'],
            'Medium_check': ['OK', 'NOK', 'nan'],
            'PS [bar(g)]': [10, 15, 20],
            'PS [bar(g)]_check': ['OK', 'NOK', 'OK'],
            'TS [°C]': [50, 100, 150],
            'TS [°C]_check': ['OK', 'OK', 'nan'],
            'Pipe Class': ['A1', 'B2', 'C3'],
            'Pipe Class_check': ['OK', 'Pipe class not found in summary', 'No pipe class assigned']
        })
          # Add status summary column (modifies df in-place)
        lp.add_status_summary(df)
        
        # Verify the status summary column was added
        assert 'Pipe Class status check' in df.columns
          # Verify position of status check column (right after 'Pipe Class_check')
        pipe_class_check_idx = df.columns.get_loc('Pipe Class_check')
        status_check_idx = df.columns.get_loc('Pipe Class status check')
        assert status_check_idx == pipe_class_check_idx + 1
        
        # Verify status summary values
        assert df.at[0, 'Pipe Class status check'] == 'OK'
        
        # Row 1 has two NOK values
        assert "Medium_check: 'NOK'" in df.at[1, 'Pipe Class status check']
        assert "PS [bar(g)]_check: 'NOK'" in df.at[1, 'Pipe Class status check']
        assert "Pipe Class_check: 'Pipe class not found in summary'" in df.at[1, 'Pipe Class status check']
        
        # Row 2 has nan values        assert "Medium_check: 'nan'" in df.at[2, 'Pipe Class status check']
        assert "TS [°C]_check: 'nan'" in df.at[2, 'Pipe Class status check']
        assert "Pipe Class_check: 'No pipe class assigned'" in df.at[2, 'Pipe Class status check']
    
    def test_empty_check_columns(self):
        """Test handling of empty check columns in status summary."""
        # Create a test dataframe with empty check columns
        df = pd.DataFrame({
            'Medium': ['Water', 'Steam'],
            'Medium_check': ['', ''],
            'PS [bar(g)]': [10, 15],
            'PS [bar(g)]_check': ['', ''],
            'Pipe Class': ['A1', 'B2'],
            'Pipe Class_check': ['OK', '']
        })
        
        # Add status summary column (modifies df in-place)
        lp.add_status_summary(df)
        
        # Verify status summary values
        # Row 0 has one OK check and two empty checks
        assert df.at[0, 'Pipe Class status check'] == 'OK'
          # Row 1 has all empty checks which is considered 'OK' since there are no problems
        assert df.at[1, 'Pipe Class status check'] == 'OK'
