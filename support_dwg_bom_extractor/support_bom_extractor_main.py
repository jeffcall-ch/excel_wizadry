#!/usr/bin/env python3
"""
Main PDF Table Extraction Pipeline
Recursively processes PDF files, extracts table boundaries, and saves results to CSV.
Updated with improved exception handling and KKS warning system.
"""

import os
import csv
import pandas as pd
from pathlib import Path
from datetime import datetime
import logging
import time
import warnings
from typing import Optional, Tuple, List
import multiprocessing
from multiprocessing import Manager

logging.getLogger('camelot').setLevel(logging.WARNING)
logging.getLogger('pdfminer').setLevel(logging.WARNING)
logging.getLogger('pdfplumber').setLevel(logging.WARNING)

# Import our custom modules
from table_boundary_finder import (
    get_table_boundaries_for_page, 
    get_total_pages,
    # Improved exceptions
    TableBoundaryError,
    PDFFileError,
    PDFNotFoundError,
    PDFCorruptedError,
    PDFEmptyError,
    PageNavigationError,
    InvalidPageNumberError,
    TableDetectionError,
    AnchorTextNotFoundError,
    TableHeadersNotFoundError,
    WeightHeaderNotFoundError,
    TotalMarkerNotFoundError,
    TableBoundaryCalculationError,
    EmptyTableRegionError,
    # KKS warnings
    KKSCodeWarning,
    NoKKSCodesFoundWarning,
    NoKKSSUCodesFoundWarning,
    NoWorkingAreaCodesFoundWarning
)

from extract_table_camelot import (
    extract_table_with_camelot,
    # Improved exceptions
    CamelotExtractionError,
    PDFDimensionError,
    PDFPageNotAccessibleError,
    PDFPageCountMismatchError,
    TableBoundsError,
    InvalidTableBoundsFormatError,
    NonNumericTableBoundsError,
    ParserError,
    StreamParserFailedError,
    LatticeParserFailedError,
    AllParsersFailedError,
    TableDataError,
    NoTablesDetectedError,
    EmptyTableExtractedError,
    TableCleaningError,
    CSVExportError
)

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Configure warnings to be captured in logs
warnings.filterwarnings("default", category=KKSCodeWarning)

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
            'table_columns_extracted',
            'kks_codes_found',
            'kks_su_codes_found',
            'working_area_codes_found'
        ])

def log_processing_event(log_file_path: str, pdf_path: str, page_num: int, 
                        status: str, details: str, processing_time: float = 0.0,
                        table_rows: int = 0, table_cols: int = 0,
                        kks_found: bool = True, kks_su_found: bool = True, working_area_found: bool = True):
    """Log a processing event to the log CSV file."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    with open(log_file_path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            timestamp, pdf_path, page_num, status, details, 
            round(processing_time, 3), table_rows, table_cols,
            kks_found, kks_su_found, working_area_found
        ])

def add_empty_row_to_csv(csv_file_path: str, full_path: str, filename: str, page_num: int):
    """Add an empty row to CSV for pages with no data."""
    # Create a DataFrame with just metadata columns
    empty_row = pd.DataFrame({
        'full_path': [full_path],
        'filename': [filename], 
        'page_number': [page_num],
        'KKS': [[]],  # Empty list for KKS codes
        'KKS/SU': [[]],  # Empty list for KKS/SU codes
        'WORKING_AREA': [[]]  # Empty list for working area codes
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
        'total_rows_extracted': 0,
        'kks_warnings_count': 0,
        'kks_su_warnings_count': 0,
        'working_area_warnings_count': 0
    }
    
    logging.info(f"Processing PDF: {pdf_filename}")
    
    try:
        # Get total page count first
        total_pages = get_total_pages(pdf_path)
        stats['total_pages'] = total_pages
        
        if total_pages == 0:
            log_processing_event(log_file_path, pdf_path, 0, 'PDF_EMPTY', 
                               'PDF file contains no pages', 0.0)
            return stats
        
        # Process each page
        for page_num in range(1, total_pages + 1):  # 1-based page numbering
            page_start_time = time.time()
            kks_found = True
            kks_su_found = True
            working_area_found = True
            
            try:
                logging.debug(f"Processing page {page_num} of {pdf_filename}")
                
                # Step 1: Get table boundaries for this specific page
                # Capture KKS warnings using warning context
                with warnings.catch_warnings(record=True) as w:
                    warnings.simplefilter("always", category=KKSCodeWarning)
                    
                    try:
                        table_bounds, kks_codes_and_kks_su_codes_and_working_area_codes = get_table_boundaries_for_page(pdf_path, page_num)
                        
                        # Check for KKS warnings
                        for warning in w:
                            if issubclass(warning.category, NoKKSCodesFoundWarning):
                                kks_found = False
                                stats['kks_warnings_count'] += 1
                                logging.warning(f"Page {page_num}: {warning.message}")
                            elif issubclass(warning.category, NoKKSSUCodesFoundWarning):
                                kks_su_found = False  
                                stats['kks_su_warnings_count'] += 1
                                logging.warning(f"Page {page_num}: {warning.message}")
                            elif issubclass(warning.category, NoWorkingAreaCodesFoundWarning):
                                working_area_found = False
                                stats['working_area_warnings_count'] += 1
                                logging.warning(f"Page {page_num}: {warning.message}")
                        
                    except AnchorTextNotFoundError as e:
                        processing_time = time.time() - page_start_time
                        log_processing_event(log_file_path, pdf_path, page_num, 
                                           'ANCHOR_NOT_FOUND', str(e), processing_time,
                                           kks_found=False, kks_su_found=False, working_area_found=False)
                        add_empty_row_to_csv(csv_file_path, full_path, pdf_filename, page_num)
                        stats['pages_processed'] += 1
                        stats['pages_no_data'] += 1
                        continue
                        
                    except (TableHeadersNotFoundError, WeightHeaderNotFoundError, 
                           TotalMarkerNotFoundError) as e:
                        processing_time = time.time() - page_start_time
                        log_processing_event(log_file_path, pdf_path, page_num, 
                                           'TABLE_STRUCTURE_ERROR', str(e), processing_time,
                                           kks_found=False, kks_su_found=False, working_area_found=False)
                        add_empty_row_to_csv(csv_file_path, full_path, pdf_filename, page_num)
                        stats['pages_processed'] += 1
                        stats['pages_no_data'] += 1
                        continue
                        
                    except (TableBoundaryCalculationError, EmptyTableRegionError) as e:
                        processing_time = time.time() - page_start_time
                        log_processing_event(log_file_path, pdf_path, page_num, 
                                           'TABLE_BOUNDARY_ERROR', str(e), processing_time,
                                           kks_found=False, kks_su_found=False, working_area_found=False)
                        add_empty_row_to_csv(csv_file_path, full_path, pdf_filename, page_num)
                        stats['pages_processed'] += 1
                        stats['pages_no_data'] += 1
                        continue
                        
                    except (InvalidPageNumberError, PDFFileError) as e:
                        processing_time = time.time() - page_start_time
                        log_processing_event(log_file_path, pdf_path, page_num, 
                                           'PAGE_ACCESS_ERROR', str(e), processing_time,
                                           kks_found=False, kks_su_found=False, working_area_found=False)
                        stats['pages_failed'] += 1
                        continue
                
                # Step 2: Extract table using Camelot
                try:
                    tables_df = extract_table_with_camelot(pdf_path, table_bounds, page_num)
                    
                except (NoTablesDetectedError, EmptyTableExtractedError) as e:
                    processing_time = time.time() - page_start_time
                    log_processing_event(log_file_path, pdf_path, page_num, 
                                       'NO_TABLE_DATA', str(e), processing_time,
                                       kks_found=kks_found, kks_su_found=kks_su_found, working_area_found=working_area_found)
                    add_empty_row_to_csv(csv_file_path, full_path, pdf_filename, page_num)
                    stats['pages_processed'] += 1
                    stats['pages_no_data'] += 1
                    continue
                    
                except (AllParsersFailedError, StreamParserFailedError, LatticeParserFailedError) as e:
                    processing_time = time.time() - page_start_time
                    log_processing_event(log_file_path, pdf_path, page_num, 
                                       'PARSER_FAILED', str(e), processing_time,
                                       kks_found=kks_found, kks_su_found=kks_su_found, working_area_found=working_area_found)
                    add_empty_row_to_csv(csv_file_path, full_path, pdf_filename, page_num)
                    stats['pages_processed'] += 1
                    stats['pages_no_data'] += 1
                    continue
                    
                except (InvalidTableBoundsFormatError, NonNumericTableBoundsError, 
                       PDFPageNotAccessibleError, PDFPageCountMismatchError) as e:
                    processing_time = time.time() - page_start_time
                    log_processing_event(log_file_path, pdf_path, page_num, 
                                       'CAMELOT_SETUP_ERROR', str(e), processing_time,
                                       kks_found=kks_found, kks_su_found=kks_su_found, working_area_found=working_area_found)
                    stats['pages_failed'] += 1
                    continue
                    
                except TableCleaningError as e:
                    processing_time = time.time() - page_start_time
                    log_processing_event(log_file_path, pdf_path, page_num, 
                                       'TABLE_CLEANING_ERROR', str(e), processing_time,
                                       kks_found=kks_found, kks_su_found=kks_su_found, working_area_found=working_area_found)
                    stats['pages_failed'] += 1
                    continue
                
                # Step 3: Add metadata columns  
                tables_df.insert(0, 'full_path', full_path)
                tables_df.insert(1, 'filename', pdf_filename)
                tables_df.insert(2, 'page_number', page_num)

                # Add KKS, KKS/SU, and working area columns to the table
                kks_codes, kks_su_codes, working_area_codes = kks_codes_and_kks_su_codes_and_working_area_codes

                # Ensure the lists are represented as full lists in each cell
                tables_df['KKS'] = [kks_codes] * len(tables_df)
                tables_df['KKS/SU'] = [kks_su_codes] * len(tables_df)
                tables_df['WORKING_AREA'] = [working_area_codes] * len(tables_df)

                # Step 4: Save to CSV
                try:
                    file_exists = os.path.exists(csv_file_path)

                    # If the file exists, drop the first 3 rows of the DataFrame
                    df_to_save = tables_df.copy()
                    if file_exists:
                        df_to_save = df_to_save.iloc[3:]

                    df_to_save.to_csv(csv_file_path, mode='a', header=not file_exists, 
                                      index=False, encoding='utf-8')
                    
                except Exception as e:
                    processing_time = time.time() - page_start_time
                    log_processing_event(log_file_path, pdf_path, page_num, 
                                       'CSV_SAVE_ERROR', f"Failed to save to CSV: {str(e)}", processing_time,
                                       kks_found=kks_found, kks_su_found=kks_su_found, working_area_found=working_area_found)
                    stats['pages_failed'] += 1
                    continue
                
                processing_time = time.time() - page_start_time
                rows_extracted = len(tables_df)
                cols_extracted = len(tables_df.columns)
                
                log_processing_event(log_file_path, pdf_path, page_num, 
                                   'SUCCESS', 
                                   f'Table extracted and saved successfully', 
                                   processing_time, rows_extracted, cols_extracted,
                                   kks_found=kks_found, kks_su_found=kks_su_found, working_area_found=working_area_found)
                
                stats['pages_processed'] += 1
                stats['pages_with_tables'] += 1
                stats['total_rows_extracted'] += rows_extracted
                
                logging.info(f"Page {page_num}: Extracted {rows_extracted} rows")
                
            except Exception as e:
                # Catch any unexpected errors
                processing_time = time.time() - page_start_time
                error_msg = f"Unexpected error processing page: {str(e)}"
                log_processing_event(log_file_path, pdf_path, page_num, 
                                   'UNEXPECTED_ERROR', error_msg, processing_time,
                                   kks_found=False, kks_su_found=False, working_area_found=False)
                stats['pages_failed'] += 1
                logging.error(f"Page {page_num}: {error_msg}")
                continue
                    
    except (PDFNotFoundError, PDFCorruptedError, PDFEmptyError) as e:
        error_msg = f"PDF file error: {str(e)}"
        log_processing_event(log_file_path, pdf_path, 0, 'PDF_FILE_ERROR', error_msg, 0.0)
        logging.error(f"Failed to process {pdf_filename}: {error_msg}")
    except Exception as e:
        error_msg = f"Critical error processing PDF: {str(e)}"
        log_processing_event(log_file_path, pdf_path, 0, 'CRITICAL_ERROR', error_msg, 0.0)
        logging.error(f"Critical failure for {pdf_filename}: {error_msg}")
    
    return stats

def process_pdf_worker(pdf_path, csv_file_path, log_file_path, lock):
    """Worker function for processing a single PDF file."""
    try:
        stats = process_single_pdf(pdf_path, csv_file_path, log_file_path)
        return stats
    except Exception as e:
        logging.error(f"Error in worker for {pdf_path}: {e}")
        return {
            'pdfs_failed': 1,
            'total_pages': 0,
            'pages_processed': 0,
            'pages_with_tables': 0,
            'pages_failed': 0,
            'pages_no_data': 0,
            'total_rows_extracted': 0,
            'kks_warnings_count': 0,
            'kks_su_warnings_count': 0,
            'working_area_warnings_count': 0
        }

def main():
    """Main processing function."""
    import argparse

    parser = argparse.ArgumentParser(description="Process PDF files and extract tables to CSV")
    parser.add_argument("input_dir", help="Directory to search for PDF files")
    parser.add_argument("--output-dir", default=".", help="Output directory for CSV files (default: current directory)")
    parser.add_argument("--single-process", action="store_true", help="Run in single process mode (useful for debugging)")

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

    start_time = time.time()

    total_stats = {
        'pdfs_processed': 0,
        'pdfs_failed': 0,
        'total_pages': 0,
        'pages_with_tables': 0,
        'pages_no_data': 0,
        'pages_failed': 0,
        'total_rows_extracted': 0,
        'kks_warnings_count': 0,
        'kks_su_warnings_count': 0,
        'working_area_warnings_count': 0
    }

    if args.single_process:
        # Single process mode for debugging
        print("Running in single process mode...")
        for pdf_path in pdf_files:
            try:
                stats = process_single_pdf(pdf_path, csv_file_path, log_file_path)
                if stats:
                    total_stats['pdfs_processed'] += 1
                    for key in ['total_pages', 'pages_with_tables', 'pages_no_data', 
                               'pages_failed', 'total_rows_extracted', 'kks_warnings_count',
                               'kks_su_warnings_count', 'working_area_warnings_count']:
                        total_stats[key] += stats.get(key, 0)
                else:
                    total_stats['pdfs_failed'] += 1
            except Exception as e:
                logging.error(f"Error processing {pdf_path}: {e}")
                total_stats['pdfs_failed'] += 1
    else:
        # Multiprocessing mode
        num_workers = max(1, multiprocessing.cpu_count() - 2)  # Leave 2 cores free
        print(f"Using {num_workers} worker processes...")
        
        manager = Manager()
        lock = manager.Lock()

        with multiprocessing.Pool(processes=num_workers) as pool:
            results = [
                pool.apply_async(process_pdf_worker, args=(pdf_path, csv_file_path, log_file_path, lock))
                for pdf_path in pdf_files
            ]

            for result in results:
                try:
                    stats = result.get()
                    if stats and 'pdfs_failed' not in stats:
                        total_stats['pdfs_processed'] += 1
                        for key in ['total_pages', 'pages_with_tables', 'pages_no_data', 
                                   'pages_failed', 'total_rows_extracted', 'kks_warnings_count',
                                   'kks_su_warnings_count', 'working_area_warnings_count']:
                            total_stats[key] += stats.get(key, 0)
                    else:
                        total_stats['pdfs_failed'] += 1
                except Exception as e:
                    logging.error(f"Error retrieving result: {e}")
                    total_stats['pdfs_failed'] += 1

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
    print(f"KKS code warnings: {total_stats['kks_warnings_count']}")
    print(f"KKS/SU code warnings: {total_stats['kks_su_warnings_count']}")
    print(f"Working area code warnings: {total_stats['working_area_warnings_count']}")
    print(f"\nOutput files:")
    print(f"  Tables: {csv_file_path}")
    print(f"  Log: {log_file_path}")

if __name__ == "__main__":
    main()