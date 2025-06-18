import pytest
import pandas as pd
import numpy as np
from unittest.mock import patch, MagicMock, mock_open

import line_list_check_pipe_class as lp

class TestDataReading:
    """Tests for data reading functions."""
    
    @patch('pandas.read_excel')
    def test_read_line_list_with_headers(self, mock_read_excel):
        """Test read_line_list function when first row contains column names."""
        # Setup mock dataframe with first row as headers
        # Create mock data with all required columns F1-F31
        mock_data = {
            "F1": ["Line No.", "L-001", "L-002"]
        }
        # Add F2 to F31
        for i in range(2, 32):
            if i == 2:  # Special case for F2
                mock_data[f"F{i}"] = ["KKS", "30ABC001", "30ABC002"]
            elif i == 7:  # Special case for Medium
                mock_data[f"F{i}"] = ["Medium", "Water", "Steam"]
            elif i == 8:  # Special case for PS [bar(g)]
                mock_data[f"F{i}"] = ["PS [bar(g)]", "10", "15"]
            else:
                mock_data[f"F{i}"] = [f"Header{i}", f"Value{i}-1", f"Value{i}-2"]
        
        # Add ComosSystemInfo column
        mock_data["ComosSystemInfo"] = ["ComosSystemInfo", "Info1", "Info2"]
        
        mock_df = pd.DataFrame(mock_data)
        mock_read_excel.return_value = mock_df
        
        # Call the function
        result_df = lp.read_line_list("fake_path.xlsx")
        
        # Verify read_excel was called correctly
        mock_read_excel.assert_called_once_with("fake_path.xlsx", sheet_name='Query')
        
        # Verify the result has proper column names and Row Number column
        assert "Line No." in result_df.columns
        assert "KKS" in result_df.columns
        assert "ComosSystemInfo" in result_df.columns
        assert "Row Number" in result_df.columns
          # Verify the result has proper data (first row is dropped, row numbers added)
        # Shape will be (2, 35) because we have:
        # - 31 original F columns renamed to their headers
        # - ComosSystemInfo
        # - Row Number
        # - 2 check columns (Medium_check and PS [bar(g)]_check)
        assert result_df.shape[0] == 2  # 2 rows
        assert result_df.shape[1] > 33  # At least 33 columns (31 + Row Number + ComosSystemInfo)
        assert result_df.iloc[0]["Row Number"] == 1
        assert result_df.iloc[1]["Row Number"] == 2
        
    @patch('pandas.read_excel')
    def test_read_line_list_without_headers(self, mock_read_excel):
        """Test read_line_list function when data doesn't have headers in first row."""
        # Setup mock dataframe without headers in first row
        mock_df = pd.DataFrame({
            "F1": ["Data1", "Data2"],
            "F2": ["Data3", "Data4"],
            "ComosSystemInfo": ["Info1", "Info2"]
        })
        mock_read_excel.return_value = mock_df
        
        # Call the function
        result_df = lp.read_line_list("fake_path.xlsx")
        
        # Verify read_excel was called correctly
        mock_read_excel.assert_called_once_with("fake_path.xlsx", sheet_name='Query')
        
        # Verify the result uses default column names
        assert "F1" in result_df.columns
        assert "F2" in result_df.columns
        assert "ComosSystemInfo" in result_df.columns
        
        # Verify check columns were added
        assert "Medium_check" not in result_df.columns  # Since "Medium" column doesn't exist
        
    @patch('pandas.read_excel')
    def test_read_pipe_class_summary(self, mock_read_excel, sample_pipe_class_df):
        """Test read_pipe_class_summary function."""
        # Setup mock
        mock_read_excel.return_value = sample_pipe_class_df
        
        # Call the function
        result_dict = lp.read_pipe_class_summary("fake_path.xlsx")
        
        # Verify read_excel was called correctly
        mock_read_excel.assert_called_once_with("fake_path.xlsx", sheet_name='Pipe Class Summary')
        
        # Verify the result is a dictionary with pipe classes as keys
        assert len(result_dict) == 4
        assert "A1" in result_dict
        assert "B2" in result_dict
        assert "C3" in result_dict
        assert "D4" in result_dict
        
        # Verify the values in the dictionary
        assert result_dict["A1"]["Medium"] == "Water"
        assert result_dict["B2"]["PN"] == 25
        assert result_dict["C3"]["EN No. Material"] == "1.4571"
        assert result_dict["D4"]["Max temperature (°C)"] == 120

    @patch('pandas.read_excel')
    def test_read_pipe_class_summary_with_missing_values(self, mock_read_excel):
        """Test read_pipe_class_summary with missing or NaN pipe class values."""
        # Setup mock dataframe with missing values
        mock_df = pd.DataFrame({
            "Pipe Class": ["A1", "", np.nan, "D4"],
            "Medium": ["Water", "Steam", "Gas", "Oil"],
            "PN": [16, 25, 40, 63]
        })
        mock_read_excel.return_value = mock_df
        
        # Call the function
        result_dict = lp.read_pipe_class_summary("fake_path.xlsx")
        
        # Verify only valid pipe classes are in the dictionary
        assert len(result_dict) == 2
        assert "A1" in result_dict
        assert "D4" in result_dict
        assert "" not in result_dict  # Empty string should be excluded
        
        # Verify the values in the dictionary
        assert result_dict["A1"]["Medium"] == "Water"
        assert result_dict["D4"]["PN"] == 63
