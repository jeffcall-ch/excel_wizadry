import os
import argparse
from pathlib import Path
import pandas as pd
import PyPDF2
from datetime import datetime
import re

def search_pdf_for_text(pdf_path, search_text, case_sensitive=False):
    """
    Search a PDF file for the specified text and return matches with page numbers.
    
    Args:
        pdf_path (str): Path to the PDF file
        search_text (str): Text to search for
        case_sensitive (bool): Whether the search should be case-sensitive
        
    Returns:
        list: List of dictionaries with page_number, match_context, and file_path
    """
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
    
    return matches

def search_directory_for_pdfs(directory, search_text, case_sensitive=False):
    """
    Search all PDF files in the specified directory and its subdirectories.
    
    Args:
        directory (str): Directory path to search in
        search_text (str): Text to search for
        case_sensitive (bool): Whether the search should be case-sensitive
        
    Returns:
        list: List of dictionaries with match information
    """
    all_matches = []
    pdf_count = 0
    match_count = 0
    
    # Walk through all subdirectories and find PDF files
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                pdf_count += 1
                
                # Display progress
                print(f"Searching {pdf_path}...")
                
                # Search this PDF
                matches = search_pdf_for_text(pdf_path, search_text, case_sensitive)
                
                if matches:
                    match_count += len(matches)
                    all_matches.extend(matches)
    
    print(f"\nSearch completed. Found {match_count} matches across {pdf_count} PDF files.")
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
    
    # Create Excel writer with openpyxl engine to support hyperlinks
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write the DataFrame to an Excel sheet
        df.to_excel(writer, index=False, sheet_name='Search Results')
        
        # Get the workbook and the worksheet
        workbook = writer.book
        worksheet = writer.sheets['Search Results']
        
        # Format the worksheet
        # Add hyperlinks
        for row_idx, file_path in enumerate(df['file_path'], start=2):  # start=2 because Excel rows are 1-indexed and we have a header
            page_number = df.iloc[row_idx-2]['page_number']
            # Create a hyperlink formula that opens the PDF to the specific page
            # Note: Not all PDF readers support direct page links, this might only work with some readers
            cell = worksheet.cell(row=row_idx, column=df.columns.get_loc('file_name') + 1)
            cell.hyperlink = file_path
            cell.style = "Hyperlink"
        
        # Set column widths
        for idx, column in enumerate(df.columns):
            column_width = max(len(str(column)), df[column].astype(str).str.len().max())
            # Convert to Excel's width format
            column_letter = worksheet.cell(1, idx+1).column_letter
            worksheet.column_dimensions[column_letter].width = min(column_width + 2, 50)
        
        # Add a summary sheet
        summary_data = {
            'Date of Search': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'Search Term': search_text,
            'Total PDF Files Searched': len(set(df['file_path'])),
            'Total Matches Found': len(df),
            'Report Generated By': 'PDF Search Tool'
        }
        
        summary_df = pd.DataFrame(list(summary_data.items()), columns=['Item', 'Value'])
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

def main():
    """Main function to parse arguments and run the search"""
    parser = argparse.ArgumentParser(description='Search PDF files for specific text and generate Excel report')
    parser.add_argument('search_text', help='Text to search for in PDF files')
    parser.add_argument('directory', help='Directory containing PDF files to search')
    parser.add_argument('--output', '-o', help='Output file path for Excel report', 
                        default=f'pdf_search_results_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    parser.add_argument('--case-sensitive', '-c', action='store_true', help='Enable case-sensitive search')
    
    args = parser.parse_args()
    
    print(f"Searching for '{args.search_text}' in {args.directory}...")
    
    # Perform search
    matches = search_directory_for_pdfs(args.directory, args.search_text, args.case_sensitive)
    
    # Create Excel report
    create_excel_report(matches, args.output, args.search_text)
    
    if matches:
        print(f"Excel report created: {args.output}")

if __name__ == "__main__":
    main()
