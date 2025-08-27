#!/usr/bin/env python3
"""
Main PDF Table Extraction Pipeline
Recursively processes PDF files, extracts table boundaries, and saves results to CSV.
"""

import os
import csv
import pandas as pd
from pathlib import Path
from datetime import datetime
import logging
import time
from typing import Optional, Tuple, List
logging.getLogger('camelot').setLevel(logging.WARNING)
logging.getLogger('pdfminer').setLevel(logging.WARNING)
logging.getLogger('pdfplumber').setLevel(logging.WARNING)

# Import our custom modules
from table_boundary_finder import get_table_boundaries_for_page, get_total_pages
from table_boundary_finder import TableBoundaryError, AnchorTextNotFoundError, TableStructureError, PDFProcessingError, PageNotFoundError
from extract_table_camelot import extract_table_with_camelot
from extract_table_camelot import CamelotExtractionError, InvalidTableBoundsError, PDFPageAccessError, EmptyTableError

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def setup_output_files(output_dir: str = ".") -> Tuple[str, str]:
    """
    Create timestamped output files for extracted tables and processing log.
    
    Returns:
        Tuple of (csv_file_path, log_file_path)
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = f"extracted_tables_{timestamp}.csv"
    log_filename = f"processing_log_{timestamp}.csv"
    
    csv_path = Path(output_dir) / csv_filename
    log_path = Path(output_dir) / log_filename
    
    return str(csv_path), str(log_path)

def initialize_log_file(log_file_path: str):
    """Initialize the processing log CSV file with headers."""
    with open(log_file_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'timestamp', 
            'pdf_path', 
            'page_num', 
            'status', 
            'details', 
            'processing_time_seconds',
            'table_rows_extracted',
            'table_columns_extracted'
        ])

def log_processing_event(log_file_path: str, pdf_path: str, page_num: int, 
                        status: str, details: str, processing_time: float = 0.0,
                        table_rows: int = 0, table_cols: int = 0):
    """Log a processing event to the log CSV file."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    with open(log_file_path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            timestamp, pdf_path, page_num, status, details, 
            round(processing_time, 3), table_rows, table_cols
        ])

def add_empty_row_to_csv(csv_file_path: str, full_path: str, filename: str, page_num: int):
    """Add an empty row to CSV for pages with no data."""
    # Create a DataFrame with just metadata columns
    empty_row = pd.DataFrame({
        'full_path': [full_path],
        'filename': [filename], 
        'page_number': [page_num]
    })
    
    # Check if CSV exists to determine if we need headers
    file_exists = os.path.exists(csv_file_path)
    empty_row.to_csv(csv_file_path, mode='a', header=not file_exists, index=False, encoding='utf-8')

def find_pdf_files(input_dir: str) -> List[str]:
    """Recursively find all PDF files in the input directory."""
    pdf_files = []
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    return sorted(pdf_files)

def process_single_pdf(pdf_path: str, csv_file_path: str, log_file_path: str) -> dict:
    """
    Process a single PDF file and extract tables from all pages.
    
    Returns:
        Dictionary with processing statistics
    """
    pdf_filename = os.path.basename(pdf_path)
    full_path = os.path.abspath(pdf_path)
    
    stats = {
        'total_pages': 0,
        'pages_processed': 0,
        'pages_with_tables': 0,
        'pages_failed': 0,
        'pages_no_data': 0,
        'total_rows_extracted': 0
    }
    
    logging.info(f"Processing PDF: {pdf_filename}")
    
    try:
        # Get total page count first
        total_pages = get_total_pages(pdf_path)
        stats['total_pages'] = total_pages
        
        if total_pages == 0:
            log_processing_event(log_file_path, pdf_path, 0, 'EMPTY_PDF', 
                               'PDF file contains no pages', 0.0)
            return stats
        
        # Process each page
        for page_num in range(1, total_pages + 1):  # 1-based page numbering
            page_start_time = time.time()
            
            try:
                logging.debug(f"Processing page {page_num} of {pdf_filename}")
                
                # Step 1: Get table boundaries for this specific page
                try:
                    table_bounds = get_table_boundaries_for_page(pdf_path, page_num)
                except AnchorTextNotFoundError as e:
                    processing_time = time.time() - page_start_time
                    log_processing_event(log_file_path, pdf_path, page_num, 
                                       'NO_ANCHOR_FOUND', str(e), processing_time)
                    add_empty_row_to_csv(csv_file_path, full_path, pdf_filename, page_num)
                    stats['pages_processed'] += 1
                    stats['pages_no_data'] += 1
                    continue
                except TableStructureError as e:
                    processing_time = time.time() - page_start_time
                    log_processing_event(log_file_path, pdf_path, page_num, 
                                       'NO_TABLE_STRUCTURE', str(e), processing_time)
                    add_empty_row_to_csv(csv_file_path, full_path, pdf_filename, page_num)
                    stats['pages_processed'] += 1
                    stats['pages_no_data'] += 1
                    continue
                except (PageNotFoundError, PDFProcessingError) as e:
                    processing_time = time.time() - page_start_time
                    log_processing_event(log_file_path, pdf_path, page_num, 
                                       'BOUNDARY_ERROR', str(e), processing_time)
                    stats['pages_failed'] += 1
                    continue
                
                # Step 2: Extract table using Camelot
                try:
                    tables_df = extract_table_with_camelot(pdf_path, table_bounds, page_num)
                except (CamelotExtractionError, EmptyTableError) as e:
                    processing_time = time.time() - page_start_time
                    log_processing_event(log_file_path, pdf_path, page_num, 
                                       'EXTRACTION_FAILED', str(e), processing_time)
                    add_empty_row_to_csv(csv_file_path, full_path, pdf_filename, page_num)
                    stats['pages_processed'] += 1
                    stats['pages_no_data'] += 1
                    continue
                except (InvalidTableBoundsError, PDFPageAccessError) as e:
                    processing_time = time.time() - page_start_time
                    log_processing_event(log_file_path, pdf_path, page_num, 
                                       'CAMELOT_ERROR', str(e), processing_time)
                    stats['pages_failed'] += 1
                    continue
                
                # Step 3: Add metadata columns  
                tables_df.insert(0, 'full_path', full_path)
                tables_df.insert(1, 'filename', pdf_filename)
                tables_df.insert(2, 'page_number', page_num)
                
                # Step 4: Save to CSV
                file_exists = os.path.exists(csv_file_path)

                # If the file exists, drop the first 3 rows of the DataFrame
                df_to_save = tables_df.copy()
                if file_exists:
                    df_to_save = df_to_save.iloc[3:]

                df_to_save.to_csv(csv_file_path, mode='a', header=not file_exists, 
                                  index=False, encoding='utf-8')
                
                processing_time = time.time() - page_start_time
                rows_extracted = len(tables_df)
                cols_extracted = len(tables_df.columns)
                
                log_processing_event(log_file_path, pdf_path, page_num, 
                                   'SUCCESS', 
                                   f'Table extracted and saved successfully', 
                                   processing_time, rows_extracted, cols_extracted)
                
                stats['pages_processed'] += 1
                stats['pages_with_tables'] += 1
                stats['total_rows_extracted'] += rows_extracted
                
                logging.info(f"Page {page_num}: Extracted {rows_extracted} rows")
                
            except Exception as e:
                # Catch any unexpected errors
                processing_time = time.time() - page_start_time
                error_msg = f"Unexpected error processing page: {str(e)}"
                log_processing_event(log_file_path, pdf_path, page_num, 
                                   'UNEXPECTED_ERROR', error_msg, processing_time)
                stats['pages_failed'] += 1
                logging.error(f"Page {page_num}: {error_msg}")
                continue
                    
    except PDFProcessingError as e:
        error_msg = f"Cannot process PDF file: {str(e)}"
        log_processing_event(log_file_path, pdf_path, 0, 'PDF_ACCESS_ERROR', error_msg, 0.0)
        logging.error(f"Failed to process {pdf_filename}: {error_msg}")
    except Exception as e:
        error_msg = f"Critical error processing PDF: {str(e)}"
        log_processing_event(log_file_path, pdf_path, 0, 'CRITICAL_ERROR', error_msg, 0.0)
        logging.error(f"Critical failure for {pdf_filename}: {error_msg}")
    
    return stats

def get_table_boundaries_for_page(pdf_path: str, page_num: int) -> Optional[Tuple]:
    """
    DEPRECATED: This function is now handled directly by table_boundary_finder.py
    """
    from table_boundary_finder import get_table_boundaries_for_page as boundary_func
    return boundary_func(pdf_path, page_num)

def main():
    """Main processing function."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Process PDF files and extract tables to CSV")
    parser.add_argument("input_dir", help="Directory to search for PDF files")
    parser.add_argument("--output-dir", default=".", help="Output directory for CSV files (default: current directory)")
    
    args = parser.parse_args()
    
    # Validate input directory
    if not os.path.isdir(args.input_dir):
        print(f"Error: Input directory '{args.input_dir}' does not exist.")
        return
    
    # Setup output files
    csv_file_path, log_file_path = setup_output_files(args.output_dir)
    initialize_log_file(log_file_path)
    
    print(f"Output files:")
    print(f"  Tables CSV: {csv_file_path}")
    print(f"  Processing Log: {log_file_path}")
    
    # Find all PDF files
    pdf_files = find_pdf_files(args.input_dir)
    print(f"\nFound {len(pdf_files)} PDF files to process")
    
    if not pdf_files:
        print("No PDF files found in the specified directory.")
        return
    
    # Process each PDF file
    total_stats = {
        'pdfs_processed': 0,
        'pdfs_failed': 0,
        'total_pages': 0,
        'pages_with_tables': 0,
        'pages_no_data': 0,
        'pages_failed': 0,
        'total_rows_extracted': 0
    }
    
    start_time = time.time()
    
    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"\nProcessing {i}/{len(pdf_files)}: {os.path.basename(pdf_path)}")
        
        try:
            stats = process_single_pdf(pdf_path, csv_file_path, log_file_path)
            
            total_stats['pdfs_processed'] += 1
            total_stats['total_pages'] += stats['total_pages']
            total_stats['pages_with_tables'] += stats['pages_with_tables']
            total_stats['pages_no_data'] += stats['pages_no_data']
            total_stats['pages_failed'] += stats['pages_failed']
            total_stats['total_rows_extracted'] += stats['total_rows_extracted']
            
            print(f"  Pages processed: {stats['pages_processed']}/{stats['total_pages']}")
            print(f"  Tables found: {stats['pages_with_tables']}")
            print(f"  No data: {stats['pages_no_data']}")
            print(f"  Failed: {stats['pages_failed']}")
            print(f"  Rows extracted: {stats['total_rows_extracted']}")
            
        except Exception as e:
            total_stats['pdfs_failed'] += 1
            logging.error(f"Failed to process PDF {pdf_path}: {str(e)}")
            log_processing_event(log_file_path, pdf_path, 0, 'CRITICAL_ERROR', 
                               f"Critical error: {str(e)}", 0.0)
    
    # Final summary
    total_time = time.time() - start_time
    print(f"\n" + "="*50)
    print(f"PROCESSING COMPLETE")
    print(f"="*50)
    print(f"Total time: {total_time:.2f} seconds")
    print(f"PDFs processed: {total_stats['pdfs_processed']}")
    print(f"PDFs failed: {total_stats['pdfs_failed']}")
    print(f"Total pages: {total_stats['total_pages']}")
    print(f"Pages with tables: {total_stats['pages_with_tables']}")
    print(f"Pages with no data: {total_stats['pages_no_data']}")
    print(f"Pages failed: {total_stats['pages_failed']}")
    print(f"Total rows extracted: {total_stats['total_rows_extracted']}")
    print(f"\nOutput files:")
    print(f"  Tables: {csv_file_path}")
    print(f"  Log: {log_file_path}")

if __name__ == "__main__":
    main()