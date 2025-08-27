#!/usr/bin/env python3
"""
Test Script for Camelot Table Extraction
Tests only the Camelot extraction part with manual coordinate input.
"""

import os
import csv
import pandas as pd
import logging
import time
from datetime import datetime
from pathlib import Path
from typing import Tuple

# Import the Camelot extraction functions and exceptions
from extract_table_camelot import extract_table_with_camelot
from extract_table_camelot import CamelotExtractionError, InvalidTableBoundsError, PDFPageAccessError, EmptyTableError

def setup_test_output_files(output_dir: str = ".") -> Tuple[str, str]:
    """
    Create timestamped test output files.
    
    Returns:
        Tuple of (csv_file_path, log_file_path)
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"test_camelot_results_{timestamp}.csv"
    log_filename = f"test_camelot_log_{timestamp}.csv"
    
    csv_path = Path(output_dir) / csv_filename
    log_path = Path(output_dir) / log_filename
    
    return str(csv_path), str(log_path)

def initialize_test_log_file(log_file_path: str):
    """Initialize the test log CSV file with headers."""
    with open(log_file_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'timestamp', 
            'pdf_path', 
            'page_num', 
            'input_coordinates',
            'status', 
            'details', 
            'processing_time_seconds',
            'table_rows_extracted',
            'table_columns_extracted'
        ])

def log_test_event(log_file_path: str, pdf_path: str, page_num: int, 
                  input_coords: str, status: str, details: str, 
                  processing_time: float = 0.0, table_rows: int = 0, table_cols: int = 0):
    """Log a test event to the log CSV file."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    with open(log_file_path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            timestamp, pdf_path, page_num, input_coords, status, details, 
            round(processing_time, 3), table_rows, table_cols
        ])

def test_camelot_extraction(pdf_path: str, coordinates: Tuple[float, float, float, float], 
                           page_num: int, csv_file_path: str, log_file_path: str):
    """
    Test Camelot extraction with given coordinates.
    
    Args:
        pdf_path (str): Path to PDF file
        coordinates (tuple): PyMuPDF style coordinates (x0, y0, x1, y1) 
        page_num (int): Page number (1-based)
        csv_file_path (str): Output CSV file path
        log_file_path (str): Log CSV file path
    """
    pdf_filename = os.path.basename(pdf_path)
    full_path = os.path.abspath(pdf_path)
    coord_str = f"({coordinates[0]}, {coordinates[1]}, {coordinates[2]}, {coordinates[3]})"
    
    print(f"\nTesting Camelot extraction:")
    print(f"  PDF: {pdf_filename}")
    print(f"  Page: {page_num}")
    print(f"  Coordinates: {coord_str}")
    
    start_time = time.time()
    
    try:
        # Extract table using Camelot
        result_df = extract_table_with_camelot(pdf_path, coordinates, page_num)
        
        # Add metadata columns
        result_df.insert(0, 'full_path', full_path)
        result_df.insert(1, 'filename', pdf_filename)
        result_df.insert(2, 'page_number', page_num)
        result_df.insert(3, 'input_coordinates', coord_str)
        
        # Save to CSV
        file_exists = os.path.exists(csv_file_path)
        result_df.to_csv(csv_file_path, mode='a', header=not file_exists, 
                        index=False, encoding='utf-8')
        
        processing_time = time.time() - start_time
        rows_extracted = len(result_df)
        cols_extracted = len(result_df.columns)
        
        log_test_event(log_file_path, pdf_path, page_num, coord_str, 
                      'SUCCESS', 'Table extracted successfully', 
                      processing_time, rows_extracted, cols_extracted)
        
        print(f"  ✅ SUCCESS: Extracted {rows_extracted} rows, {cols_extracted} columns")
        print(f"  Processing time: {processing_time:.3f} seconds")
        
        # Display first few rows for verification
        if len(result_df) > 0:
            print(f"\nFirst 3 rows of extracted data:")
            print(result_df.head(3).to_string(index=False))
        
        return True
        
    except (CamelotExtractionError, EmptyTableError) as e:
        processing_time = time.time() - start_time
        log_test_event(log_file_path, pdf_path, page_num, coord_str,
                      'EXTRACTION_FAILED', str(e), processing_time)
        print(f"  ❌ EXTRACTION FAILED: {str(e)}")
        return False
        
    except (InvalidTableBoundsError, PDFPageAccessError) as e:
        processing_time = time.time() - start_time
        log_test_event(log_file_path, pdf_path, page_num, coord_str,
                      'INVALID_INPUT', str(e), processing_time)
        print(f"  ❌ INVALID INPUT: {str(e)}")
        return False
        
    except Exception as e:
        processing_time = time.time() - start_time
        log_test_event(log_file_path, pdf_path, page_num, coord_str,
                      'UNEXPECTED_ERROR', str(e), processing_time)
        print(f"  ❌ UNEXPECTED ERROR: {str(e)}")
        return False

def main():
    """Main test function."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Test Camelot table extraction with manual coordinates",
        epilog="Example: python camelot_test_script.py document.pdf 100 200 400 600 --page 1"
    )
    parser.add_argument("pdf_path", help="Path to PDF file")
    parser.add_argument("x0", type=float, help="Left boundary (PyMuPDF style)")
    parser.add_argument("y0", type=float, help="Top boundary (PyMuPDF style)")  
    parser.add_argument("x1", type=float, help="Right boundary (PyMuPDF style)")
    parser.add_argument("y1", type=float, help="Bottom boundary (PyMuPDF style)")
    parser.add_argument("--page", type=int, default=1, help="Page number (1-based, default: 1)")
    parser.add_argument("--output-dir", default=".", help="Output directory (default: current directory)")
    
    args = parser.parse_args()
    
    # Configure logging for test
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
    
    # Validate input file
    if not Path(args.pdf_path).is_file():
        print(f"❌ Error: PDF file not found: {args.pdf_path}")
        return
    
    # Validate coordinates
    coordinates = (args.x0, args.y0, args.x1, args.y1)
    if args.x0 >= args.x1 or args.y0 >= args.y1:
        print(f"❌ Error: Invalid coordinates. x0 must be < x1 and y0 must be < y1")
        print(f"  Given: x0={args.x0}, y0={args.y0}, x1={args.x1}, y1={args.y1}")
        return
    
    # Setup output files
    csv_file_path, log_file_path = setup_test_output_files(args.output_dir)
    initialize_test_log_file(log_file_path)
    
    print(f"Camelot Extraction Test")
    print(f"=" * 50)
    print(f"Output files:")
    print(f"  Results CSV: {csv_file_path}")
    print(f"  Test Log: {log_file_path}")
    
    # Run the test
    success = test_camelot_extraction(
        args.pdf_path, 
        coordinates, 
        args.page,
        csv_file_path, 
        log_file_path
    )
    
    if success:
        print(f"\n✅ Test completed successfully!")
        print(f"Check the output files for detailed results.")
    else:
        print(f"\n❌ Test failed. Check the log file for details.")
        print(f"Log file: {log_file_path}")

if __name__ == "__main__":
    main()