import os
import csv
from pathlib import Path
import PyPDF2
import re

def extract_dates(text):
    """
    Extract years and complete dates from text.
    Returns a list of unique dates found.
    """
    dates_found = []
    
    # Pattern for years (1900-2099)
    year_pattern = r'\b(19\d{2}|20\d{2})\b'
    years = re.findall(year_pattern, text)
    dates_found.extend(years)
    
    # Pattern for dates like: 2021-05-15, 2021/05/15, 2021.05.15
    date_pattern1 = r'\b(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})\b'
    dates1 = re.findall(date_pattern1, text)
    dates_found.extend(dates1)
    
    # Pattern for dates like: 15-05-2021, 15/05/2021, 15.05.2021
    date_pattern2 = r'\b(\d{1,2}[-/.]\d{1,2}[-/.]\d{4})\b'
    dates2 = re.findall(date_pattern2, text)
    dates_found.extend(dates2)
    
    # Pattern for dates like: May 15, 2021 or 15 May 2021
    date_pattern3 = r'\b((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4})\b'
    dates3 = re.findall(date_pattern3, text, re.IGNORECASE)
    dates_found.extend(dates3)
    
    # Pattern for dates like: 2021 May 15
    date_pattern4 = r'\b(\d{4}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2})\b'
    dates4 = re.findall(date_pattern4, text, re.IGNORECASE)
    dates_found.extend(dates4)
    
    # Remove duplicates while preserving order
    unique_dates = []
    for date in dates_found:
        if date not in unique_dates:
            unique_dates.append(date)
    
    return unique_dates

def search_pdfs_for_keywords():
    """
    Search all PDFs in C:\\temp\\WERK for FERO and specific pipe class codes.
    Save matching file paths to a CSV file.
    """
    # Configuration
    search_directory = r"C:\temp\WERK"
    output_csv = "fero_pipe_class_files.csv"
    
    # Keywords to search for
    required_word = "FERO"
    pipe_class_codes = [
        "EFDX", "EFDJ", "EHDX", "EEDX", "ECDE", "ECDM", "EHDQ", "EEDQ", 
        "EHGN", "EHFD", "EEFD", "AHDX", "AEDX", "ACDE", "ACDM", "AHDQ", 
        "AHFD", "AHGN"
    ]
    
    # Results storage
    matching_files = []
    
    # Get script directory for output
    script_dir = Path(__file__).parent
    output_path = script_dir / output_csv
    
    print(f"Searching for PDFs in: {search_directory}")
    print(f"Looking for '{required_word}' AND one of: {', '.join(pipe_class_codes)}")
    print("-" * 80)
    
    # Check if directory exists
    if not os.path.exists(search_directory):
        print(f"ERROR: Directory does not exist: {search_directory}")
        return
    
    # Find all PDF files
    pdf_files = []
    for root, dirs, files in os.walk(search_directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    print(f"Found {len(pdf_files)} PDF files to process\n")
    
    # Process each PDF
    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"[{i}/{len(pdf_files)}] Processing: {os.path.basename(pdf_path)}")
        
        try:
            # Read PDF content
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # Extract text from all pages
                text_content = ""
                for page in pdf_reader.pages:
                    text_content += page.extract_text()
                
                # Convert to uppercase for case-insensitive search
                text_upper = text_content.upper()
                
                # Check if FERO is present
                if required_word in text_upper:
                    # Check if any pipe class code is present
                    found_codes = [code for code in pipe_class_codes if code in text_upper]
                    
                    if found_codes:
                        # Extract dates from path, filename, and content
                        dates_in_path = extract_dates(os.path.dirname(pdf_path))
                        dates_in_filename = extract_dates(os.path.basename(pdf_path))
                        dates_in_content = extract_dates(text_content)
                        
                        matching_files.append({
                            'file_path': pdf_path,
                            'found_codes': ', '.join(found_codes),
                            'dates_in_path': ', '.join(dates_in_path) if dates_in_path else '',
                            'dates_in_filename': ', '.join(dates_in_filename) if dates_in_filename else '',
                            'dates_in_content': ', '.join(dates_in_content) if dates_in_content else ''
                        })
                        print(f"  ✓ MATCH FOUND - Contains: {', '.join(found_codes)}")
        
        except Exception as e:
            print(f"  ✗ ERROR reading file: {str(e)}")
    
    # Write results to CSV
    print("\n" + "=" * 80)
    print(f"Found {len(matching_files)} matching files")
    print(f"Writing results to: {output_path}")
    
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['file_path', 'found_codes', 'dates_in_path', 'dates_in_filename', 'dates_in_content']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for file_info in matching_files:
            writer.writerow(file_info)
    
    print(f"✓ CSV file created successfully!")
    print("\nMatching files:")
    for file_info in matching_files:
        print(f"  - {file_info['file_path']}")
        print(f"    Codes: {file_info['found_codes']}")
        if file_info['dates_in_path']:
            print(f"    Dates in path: {file_info['dates_in_path']}")
        if file_info['dates_in_filename']:
            print(f"    Dates in filename: {file_info['dates_in_filename']}")
        if file_info['dates_in_content']:
            print(f"    Dates in content: {file_info['dates_in_content']}")

if __name__ == "__main__":
    search_pdfs_for_keywords()
