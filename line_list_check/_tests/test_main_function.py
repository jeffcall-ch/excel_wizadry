import pytest
import pandas as pd
import os
from unittest.mock import patch, MagicMock, call

import line_list_check_pipe_class as lp

class TestMainFunction:
    """Tests for the main function and integration tests."""
    
    @patch('line_list_check_pipe_class.generate_pipe_class_summary')
    @patch('line_list_check_pipe_class.save_to_excel')
    @patch('line_list_check_pipe_class.validate_pipe_data')
    @patch('line_list_check_pipe_class.read_pipe_class_summary')
    @patch('line_list_check_pipe_class.read_line_list')
    @patch('line_list_check_pipe_class.setup_paths')
    def test_main_successful_execution(self, mock_setup_paths, mock_read_line_list, 
                                     mock_read_pipe_class_summary, mock_validate_pipe_data,
                                     mock_save_to_excel, mock_generate_summary):
        """Test main function with successful execution."""
        # Setup mocks
        mock_dirs = {'input_dir': '/input', 'pipe_class_dir': '/pipe_class', 'output_dir': '/output'}
        mock_files = {
            'line_list_file': '/input/line_list.xlsx',
            'pipe_class_file': '/pipe_class/summary.xlsx',
            'output_file': '/output/results.xlsx',
            'summary_file': '/output/summary.xlsx'
        }
        mock_setup_paths.return_value = (mock_dirs, mock_files)
        
        mock_line_list_df = pd.DataFrame({'Column1': [1, 2]})
        mock_read_line_list.return_value = mock_line_list_df
        
        mock_pipe_class_dict = {'A1': {'Medium': 'Water'}}
        mock_read_pipe_class_summary.return_value = mock_pipe_class_dict
        
        mock_validated_df = pd.DataFrame({'Column1': [1, 2], 'Column1_check': ['OK', 'NOK']})
        mock_validate_pipe_data.return_value = mock_validated_df
        
        mock_save_to_excel.return_value = True
        mock_generate_summary.return_value = True
        
        # Call main function
        lp.main()
        
        # Verify all functions were called in sequence with correct parameters
        mock_setup_paths.assert_called_once()
        mock_read_line_list.assert_called_once_with('/input/line_list.xlsx')
        mock_read_pipe_class_summary.assert_called_once_with('/pipe_class/summary.xlsx')
        mock_validate_pipe_data.assert_called_once_with(mock_line_list_df, mock_pipe_class_dict)
        mock_save_to_excel.assert_called_once_with(mock_validated_df, '/output/results.xlsx')
        mock_generate_summary.assert_called_once_with(mock_validated_df, mock_pipe_class_dict, '/output/summary.xlsx')

    @patch('line_list_check_pipe_class.generate_pipe_class_summary')
    @patch('line_list_check_pipe_class.save_to_excel')
    @patch('line_list_check_pipe_class.validate_pipe_data')
    @patch('line_list_check_pipe_class.read_pipe_class_summary')
    @patch('line_list_check_pipe_class.read_line_list')
    @patch('line_list_check_pipe_class.setup_paths')
    def test_main_with_save_error(self, mock_setup_paths, mock_read_line_list, 
                                mock_read_pipe_class_summary, mock_validate_pipe_data,
                                mock_save_to_excel, mock_generate_summary):
        """Test main function with error in save_to_excel."""
        # Setup mocks
        mock_dirs = {'input_dir': '/input', 'pipe_class_dir': '/pipe_class', 'output_dir': '/output'}
        mock_files = {
            'line_list_file': '/input/line_list.xlsx',
            'pipe_class_file': '/pipe_class/summary.xlsx',
            'output_file': '/output/results.xlsx',
            'summary_file': '/output/summary.xlsx'
        }
        mock_setup_paths.return_value = (mock_dirs, mock_files)
        
        mock_line_list_df = pd.DataFrame({'Column1': [1, 2]})
        mock_read_line_list.return_value = mock_line_list_df
        
        mock_pipe_class_dict = {'A1': {'Medium': 'Water'}}
        mock_read_pipe_class_summary.return_value = mock_pipe_class_dict
        
        mock_validated_df = pd.DataFrame({'Column1': [1, 2], 'Column1_check': ['OK', 'NOK']})
        mock_validate_pipe_data.return_value = mock_validated_df
        
        mock_save_to_excel.return_value = False
        mock_generate_summary.return_value = True
        
        # Call main function
        lp.main()
        
        # Verify correct functions were called
        mock_save_to_excel.assert_called_once_with(mock_validated_df, '/output/results.xlsx')
        mock_generate_summary.assert_called_once_with(mock_validated_df, mock_pipe_class_dict, '/output/summary.xlsx')

    @patch('line_list_check_pipe_class.generate_pipe_class_summary')
    @patch('line_list_check_pipe_class.save_to_excel')
    @patch('line_list_check_pipe_class.validate_pipe_data')
    @patch('line_list_check_pipe_class.read_pipe_class_summary')
    @patch('line_list_check_pipe_class.read_line_list')
    @patch('line_list_check_pipe_class.setup_paths')
    def test_main_with_summary_error(self, mock_setup_paths, mock_read_line_list, 
                                  mock_read_pipe_class_summary, mock_validate_pipe_data,
                                  mock_save_to_excel, mock_generate_summary):
        """Test main function with error in generate_pipe_class_summary."""
        # Setup mocks
        mock_dirs = {'input_dir': '/input', 'pipe_class_dir': '/pipe_class', 'output_dir': '/output'}
        mock_files = {
            'line_list_file': '/input/line_list.xlsx',
            'pipe_class_file': '/pipe_class/summary.xlsx',
            'output_file': '/output/results.xlsx',
            'summary_file': '/output/summary.xlsx'
        }
        mock_setup_paths.return_value = (mock_dirs, mock_files)
        
        mock_line_list_df = pd.DataFrame({'Column1': [1, 2]})
        mock_read_line_list.return_value = mock_line_list_df
        
        mock_pipe_class_dict = {'A1': {'Medium': 'Water'}}
        mock_read_pipe_class_summary.return_value = mock_pipe_class_dict
        
        mock_validated_df = pd.DataFrame({'Column1': [1, 2], 'Column1_check': ['OK', 'NOK']})
        mock_validate_pipe_data.return_value = mock_validated_df
        
        mock_save_to_excel.return_value = True
        mock_generate_summary.return_value = False
        
        # Call main function
        lp.main()
        
        # Verify correct functions were called
        mock_save_to_excel.assert_called_once_with(mock_validated_df, '/output/results.xlsx')
        mock_generate_summary.assert_called_once_with(mock_validated_df, mock_pipe_class_dict, '/output/summary.xlsx')

    @patch('line_list_check_pipe_class.print')
    @patch('line_list_check_pipe_class.setup_paths')
    def test_main_with_exception(self, mock_setup_paths, mock_print):
        """Test main function with an exception being raised."""
        # Setup mocks
        mock_setup_paths.side_effect = Exception("Test error")
        
        # Call main function
        lp.main()
        
        # Verify exception was caught and printed
        mock_print.assert_any_call("Error in main process: Test error")
