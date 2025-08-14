import os
import argparse
from pathlib import Path
import pandas as pd
import PyPDF2
from datetime import datetime
import re
from multiprocessing import Pool, cpu_count
from functools import partial
import time

def search_pdf_for_text(pdf_info):
    """
    Search a PDF file for the specified text and return matches with page numbers.
    Modified to work with multiprocessing.
    
    Args:
        pdf_info (tuple): (pdf_path, search_text, case_sensitive)
        
    Returns:
        tuple: (pdf_path, matches_list)
    """
    pdf_path, search_text, case_sensitive = pdf_info
    matches = []
    
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            num_pages = len(reader.pages)
            
            for page_num in range(num_pages):
                page = reader.pages[page_num]
                text = page.extract_text()
                
                if text:
                    # Handle case sensitivity
                    if not case_sensitive:
                        search_pattern = search_text.lower()
                        page_text = text.lower()
                    else:
                        search_pattern = search_text
                        page_text = text
                    
                    # Check if the text exists on this page
                    if search_pattern in page_text:
                        # Find all occurrences and get surrounding context
                        for match in re.finditer(re.escape(search_pattern), page_text, re.IGNORECASE if not case_sensitive else 0):
                            start_pos = max(0, match.start() - 50)
                            end_pos = min(len(page_text), match.end() + 50)
                            
                            # Get context (text before and after the match)
                            context = "..." + page_text[start_pos:end_pos].replace('\n', ' ') + "..."
                            
                            # Store match information
                            matches.append({
                                'file_path': pdf_path,
                                'page_number': page_num + 1,  # +1 because page numbers start from 1, not 0
                                'match_context': context
                            })
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
    
    return pdf_path, matches

def search_directory_for_pdfs_parallel(directory, search_text, case_sensitive=False, num_workers=None):
    """
    Search all PDF files in the specified directory and its subdirectories using parallel processing.
    
    Args:
        directory (str): Directory to search
        search_text (str): Text to search for
        case_sensitive (bool): Whether the search should be case-sensitive
        num_workers (int): Number of worker processes (defaults to CPU count)
        
    Returns:
        list: List of dictionaries with match information
    """
    if num_workers is None:
        num_workers = cpu_count()
    
    print(f"Using {num_workers} worker processes for parallel PDF processing...")
    
    # Collect all PDF files first
    pdf_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                pdf_files.append(pdf_path)
    
    print(f"Found {len(pdf_files)} PDF files to process...")
    
    if not pdf_files:
        print("No PDF files found in the specified directory.")
        return []
    
    # Prepare data for parallel processing
    pdf_data = [(pdf_path, search_text, case_sensitive) for pdf_path in pdf_files]
    
    all_matches = []
    processed_count = 0
    match_count = 0
    
    start_time = time.time()
    
    # Process PDFs in parallel with progress tracking
    with Pool(processes=num_workers) as pool:
        # Use imap for better progress tracking
        results = pool.imap(search_pdf_for_text, pdf_data)
        
        for pdf_path, matches in results:
            processed_count += 1
            if matches:
                match_count += len(matches)
                all_matches.extend(matches)
            
            # Show progress every 10 files or at key intervals
            if processed_count % 10 == 0 or processed_count in [1, 5, len(pdf_files)]:
                elapsed = time.time() - start_time
                rate = processed_count / elapsed if elapsed > 0 else 0
                remaining = len(pdf_files) - processed_count
                eta = remaining / rate if rate > 0 else 0
                
                print(f"Progress: {processed_count}/{len(pdf_files)} files processed "
                      f"({processed_count/len(pdf_files)*100:.1f}%) - "
                      f"Rate: {rate:.1f} files/sec - "
                      f"ETA: {eta:.0f}s - "
                      f"Matches found so far: {match_count}")
    
    elapsed_total = time.time() - start_time
    print(f"\nSearch completed in {elapsed_total:.2f} seconds!")
    print(f"Found {match_count} matches across {len(pdf_files)} PDF files.")
    print(f"Average processing rate: {len(pdf_files)/elapsed_total:.1f} files/second")
    
    return all_matches

def create_excel_report(matches, output_file, search_text):
    """
    Create an Excel report with match results and hyperlinks to the original files.
    
    Args:
        matches (list): List of dictionaries with match information
        output_file (str): Path to save the Excel report
        search_text (str): The text that was searched for
    """
    if not matches:
        print("No matches found. No Excel report generated.")
        return
    
    # Create a DataFrame from the matches
    df = pd.DataFrame(matches)
    
    # Add a column for hyperlinks
    df['file_name'] = df['file_path'].apply(lambda x: os.path.basename(x))
    
    # Ensure the output directory exists
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Create Excel file with multiple sheets if needed
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Main results sheet
        df.to_excel(writer, sheet_name='Search Results', index=False)
        
        # Summary sheet
        summary_data = {
            'Search Term': [search_text],
            'Total Matches': [len(matches)],
            'Files with Matches': [df['file_path'].nunique()],
            'Search Date': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Files summary sheet
        file_summary = df.groupby('file_name').agg({
            'page_number': 'count',
            'file_path': 'first'
        }).rename(columns={'page_number': 'match_count'}).reset_index()
        file_summary.to_excel(writer, sheet_name='Files Summary', index=False)
    
    print(f"Excel report created: {output_file}")

def main():
    parser = argparse.ArgumentParser(description='Search for text in PDF files with parallel processing')
    parser.add_argument('search_text', help='Text to search for')
    parser.add_argument('directory', help='Directory to search for PDF files')
    parser.add_argument('-o', '--output', help='Output Excel file path (optional)')
    parser.add_argument('-c', '--case-sensitive', action='store_true', help='Case-sensitive search')
    parser.add_argument('-j', '--jobs', type=int, help=f'Number of parallel workers (default: {cpu_count()})')
    
    args = parser.parse_args()
    
    # Validate directory
    if not os.path.isdir(args.directory):
        print(f"Error: Directory '{args.directory}' does not exist.")
        return
    
    # Set number of workers
    num_workers = args.jobs if args.jobs else cpu_count()
    if num_workers < 1:
        num_workers = 1
    elif num_workers > cpu_count():
        print(f"Warning: Requested {num_workers} workers, but only {cpu_count()} CPU cores available.")
    
    print(f"Searching for '{args.search_text}' in {args.directory}...")
    print(f"Case sensitive: {args.case_sensitive}")
    print(f"Using {num_workers} parallel workers")
    
    # Search for matches
    matches = search_directory_for_pdfs_parallel(
        args.directory, 
        args.search_text, 
        args.case_sensitive,
        num_workers
    )
    
    # Generate output filename if not provided
    if args.output:
        output_file = args.output
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"pdf_search_results_{timestamp}.xlsx"
    
    # Create Excel report
    create_excel_report(matches, output_file, args.search_text)

if __name__ == "__main__":
    main()
