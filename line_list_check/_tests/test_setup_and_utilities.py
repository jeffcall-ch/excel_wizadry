import os
import pytest
import pandas as pd
import numpy as np
from unittest.mock import patch, MagicMock

import line_list_check_pipe_class as lp

class TestSetupAndUtilities:
    """Tests for setup and utility functions."""
    
    def test_setup_paths(self):
        """Test setup_paths function creates correct directory and file paths."""
        with patch('os.path.dirname') as mock_dirname:
            mock_dirname.return_value = "/fake/path"
            dirs, files = lp.setup_paths()
            
            # Check directories
            assert dirs['input_dir'] == os.path.join("/fake/path", "input_file")
            assert dirs['pipe_class_dir'] == os.path.join("/fake/path", "pipe_class_summary_file")
            assert dirs['output_dir'] == os.path.join("/fake/path", "output")
            
            # Check files
            assert files['line_list_file'] == os.path.join(dirs['input_dir'], "Export 17.06.2025_LS.xlsx")
            assert files['pipe_class_file'] == os.path.join(dirs['pipe_class_dir'], "PIPE CLASS SUMMARY_LS_06.06.2025_updated_column_names.xlsx")
            assert files['output_file'] == os.path.join(dirs['output_dir'], "Line_List_with_Matches.xlsx")
    
    def test_extract_numeric_part(self):
        """Test extract_numeric_part function with various inputs."""
        # Test with standard cases
        assert lp.extract_numeric_part("DN 25") == 25
        assert lp.extract_numeric_part("PN 16") == 16
        assert lp.extract_numeric_part("100") == 100
        
        # Test with non-standard cases
        assert lp.extract_numeric_part("text123text") == 123
        assert lp.extract_numeric_part("DN25") == 25
        assert lp.extract_numeric_part("PN-16") == 16
        
        # Test with edge cases
        assert lp.extract_numeric_part("") == "nan"
        assert lp.extract_numeric_part("no numbers") == "nan"
        assert lp.extract_numeric_part(None) == "nan"
        assert lp.extract_numeric_part(np.nan) == "nan"
        
    def test_add_check_columns(self):
        """Test add_check_columns function."""
        # Create sample dataframe
        df = pd.DataFrame({
            'Medium': ['Water', 'Steam'],
            'PS [bar(g)]': [10, 15],
            'TS [°C]': [50, 100],
            'DN': ['DN 25', 'DN 50'],
            'PN': ['PN 16', 'PN 25'],
            'EN No. Material': ['1.0425', '1.0460'],
            'Pipe Class': ['A1', 'B2'],
            'Other Column': ['Value1', 'Value2']
        })
        
        # Add check columns
        result_df = lp.add_check_columns(df)
        
        # Verify check columns were added
        check_columns = [
            'Medium_check', 'PS [bar(g)]_check', 'TS [°C]_check', 
            'DN_check', 'PN_check', 'EN No. Material_check', 'Pipe Class_check'
        ]
        
        for col in check_columns:
            assert col in result_df.columns
        
        # Verify check columns are in the correct position
        for base_col in ['Medium', 'PS [bar(g)]', 'TS [°C]', 'DN', 'PN', 'EN No. Material', 'Pipe Class']:
            base_idx = result_df.columns.get_loc(base_col)
            check_idx = result_df.columns.get_loc(f"{base_col}_check")
            assert check_idx == base_idx + 1
        
        # Verify original column not in the check list is unchanged
        assert 'Other Column' in result_df.columns
        assert 'Other Column_check' not in result_df.columns
