import sys
import os
import pytest
import pandas as pd
import io
from unittest.mock import patch

# Ensure parent directory is in sys.path for module import
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from line_list_compare import compare_pipeline_lists

def test_check_kks_unique_with_duplicates(tmp_path, monkeypatch):
    """Test that the script detects and reports duplicate KKS values correctly."""
    import io
    from contextlib import redirect_stdout
    import sys
    
    # Create test files with duplicated KKS values in old file
    old_file = tmp_path / 'old_with_dups.xlsx'
    new_file = tmp_path / 'new.xlsx'
    
    # Create test data with duplicates in old file
    old_df = pd.DataFrame({
        'KKS': ['A1', 'B1', 'A1', 'C1'],  # A1 is duplicated
        'Col1': [1, 2, 3, 4]
    })
    new_df = pd.DataFrame({
        'KKS': ['D1', 'E1', 'F1'],  # Use different values to avoid comparison issues
        'Col1': [10, 20, 30]
    })
    
    # Write to Excel files
    with pd.ExcelWriter(old_file, engine='openpyxl') as writer:
        old_df.to_excel(writer, sheet_name='Query', index=False)
    with pd.ExcelWriter(new_file, engine='openpyxl') as writer:
        new_df.to_excel(writer, sheet_name='Query', index=False)
    
    # Mock sys.exit to prevent test from exiting
    original_exit = sys.exit
    exit_called = False
    
    def mock_exit(code=0):
        nonlocal exit_called
        exit_called = True
        assert code == 1
        # Instead of exiting, raise exception to stop execution
        raise RuntimeError("sys.exit called with code 1")
    
    monkeypatch.setattr('sys.exit', mock_exit)
    
    # Capture stdout
    f = io.StringIO()
    with redirect_stdout(f):
        try:
            compare_pipeline_lists(
                old_file=str(old_file),
                new_file=str(new_file),
                output_file=str(old_file).replace('.xlsx', '_output.xlsx')
            )
        except RuntimeError as e:
            # Expected to raise due to our mock_exit
            assert "sys.exit called" in str(e)
    
    # Check that sys.exit was called
    assert exit_called
    
    # Check output for duplicate message
    output = f.getvalue()
    assert "ERROR: Non-unique KKS values found in old file:" in output
    assert "KKS value 'A1'" in output

def test_check_kks_unique_with_duplicates_in_new(tmp_path, monkeypatch):
    """Test that the script detects and reports duplicate KKS values in new file correctly."""
    import io
    from contextlib import redirect_stdout
    import sys
    
    # Create test files with duplicated KKS values in new file
    old_file = tmp_path / 'old.xlsx'
    new_file = tmp_path / 'new_with_dups.xlsx'
    
    # Create test data with duplicates in new file
    old_df = pd.DataFrame({
        'KKS': ['D1', 'E1', 'F1'],  # Use different values to avoid Series comparison error
        'Col1': [1, 2, 4]
    })
    new_df = pd.DataFrame({
        'KKS': ['A1', 'B1', 'B1', 'C1'],  # B1 is duplicated
        'Col1': [10, 20, 25, 30]
    })
    
    # Write to Excel files
    with pd.ExcelWriter(old_file, engine='openpyxl') as writer:
        old_df.to_excel(writer, sheet_name='Query', index=False)
    with pd.ExcelWriter(new_file, engine='openpyxl') as writer:
        new_df.to_excel(writer, sheet_name='Query', index=False)
    
    # Mock sys.exit to prevent test from actually exiting
    exit_called = False
    def mock_exit(code=0):
        nonlocal exit_called
        exit_called = True
        assert code == 1
        # Instead of exiting, raise exception to stop execution
        raise RuntimeError("sys.exit called with code 1")
    
    monkeypatch.setattr('sys.exit', mock_exit)
    
    # Capture stdout
    f = io.StringIO()
    with redirect_stdout(f):
        try:
            compare_pipeline_lists(
                old_file=str(old_file),
                new_file=str(new_file),
                output_file=str(old_file).replace('.xlsx', '_output.xlsx')
            )
        except RuntimeError as e:
            # Expected to raise due to our mock_exit
            assert "sys.exit called" in str(e)
    
    # Check that sys.exit was called
    assert exit_called
    
    # Check output for duplicate message
    output = f.getvalue()
    assert "ERROR: Non-unique KKS values found in new file:" in output
    assert "KKS value 'B1'" in output

def test_missing_kks_header(tmp_path, monkeypatch):
    """Test that ValueError is raised when KKS header is not found."""
    import io
    from contextlib import redirect_stdout
    import sys
    
    # Create test file without KKS header
    test_file = tmp_path / 'no_kks_header.xlsx'
    df = pd.DataFrame({
        'ID': ['A1', 'B1', 'C1'],  # No KKS column
        'Col1': [1, 2, 3]
    })
    
    # Write to Excel file
    with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Query', index=False)
    
    # Create a new file that has KKS so we can test the specific file's error
    valid_file = tmp_path / 'valid.xlsx'
    pd.DataFrame({'KKS': ['A1']}).to_excel(valid_file, index=False, sheet_name='Query')
    
    # Mock sys.exit to prevent test from exiting
    exit_called = False
    def mock_exit(code=0):
        nonlocal exit_called
        exit_called = True
        assert code == 1
        # Instead of exiting, raise exception to stop execution
        raise RuntimeError("sys.exit called with code 1")
    
    monkeypatch.setattr('sys.exit', mock_exit)
    
    # Capture stdout
    f = io.StringIO()
    with redirect_stdout(f):
        # The function should call sys.exit when KKS header is missing
        with pytest.raises(RuntimeError, match="sys.exit called with code 1"):
            compare_pipeline_lists(old_file=str(test_file), new_file=str(valid_file))
    
    # Check that sys.exit was called
    assert exit_called
    
    # Check output for error message
    output = f.getvalue()
    assert "ERROR: No 'KKS' header found" in output

def test_file_not_found_handling(monkeypatch):
    """Test that the script gracefully handles file not found errors."""
    non_existent_file = "non_existent_file.xlsx"
    
    log_messages = []
    
    # Mock logging to capture errors
    def mock_error(msg):
        log_messages.append(msg)
        
    monkeypatch.setattr('logging.error', mock_error)
    
    # This should not raise an exception
    compare_pipeline_lists(old_file=non_existent_file)
    
    # Check that error was logged
    assert any("Failed to read old_file:" in msg for msg in log_messages)
    assert any("No such file or directory" in msg for msg in log_messages)

def test_save_error_handling(tmp_path, monkeypatch):
    """Test handling of errors when saving the output file."""
    # Create valid test files
    old_file = tmp_path / 'old.xlsx'
    new_file = tmp_path / 'new.xlsx'
    
    # Create test data
    old_df = pd.DataFrame({
        'KKS': ['A1', 'B1', 'C1'],
        'Col1': [1, 2, 3]
    })
    new_df = pd.DataFrame({
        'KKS': ['A1', 'B1', 'D1'],  # C1 removed, D1 added
        'Col1': [10, 20, 40]
    })
    
    # Write to Excel files
    with pd.ExcelWriter(old_file, engine='openpyxl') as writer:
        old_df.to_excel(writer, sheet_name='Query', index=False)
    with pd.ExcelWriter(new_file, engine='openpyxl') as writer:
        new_df.to_excel(writer, sheet_name='Query', index=False)
    
    # Create a bad output path to force save error
    bad_output = os.path.join("Z:\\nonexistent\\dir", "output.xlsx")  # Invalid path should cause save error
    
    log_messages = []
    
    # Mock logging to capture errors
    def mock_error(msg):
        log_messages.append(msg)
        
    monkeypatch.setattr('logging.error', mock_error)
    
    # This should not raise an exception despite save error
    compare_pipeline_lists(
        old_file=str(old_file),
        new_file=str(new_file),
        output_file=bad_output
    )
    
    # Check that error was logged
    assert any("Failed to save output file:" in msg for msg in log_messages)
