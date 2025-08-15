import os
import argparse
from pathlib import Path
import pandas as pd
import PyPDF2
from datetime import datetime
import re
from typing import List, Dict, Tuple

def parse_boolean_query(query: str) -> List[Tuple[str, str]]:
    """
    Parse a boolean search query into tokens.
    
    Args:
        query (str): Boolean search query like "term1 AND term2 OR term3"
        
    Returns:
        List[Tuple[str, str]]: List of (token_type, value) tuples
                              token_type can be 'TERM', 'AND', 'OR', 'NOT', 'LPAREN', 'RPAREN'
    """
    # Replace operators with standardized versions and add spaces
    query = re.sub(r'\bAND\b', ' AND ', query)
    query = re.sub(r'\bOR\b', ' OR ', query)
    query = re.sub(r'\bNOT\b', ' NOT ', query)
    query = re.sub(r'\(', ' ( ', query)
    query = re.sub(r'\)', ' ) ', query)
    
    # Split into tokens and filter out empty strings
    tokens = [token.strip() for token in query.split() if token.strip()]
    
    parsed_tokens = []
    for token in tokens:
        if token == 'AND':
            parsed_tokens.append(('AND', token))
        elif token == 'OR':
            parsed_tokens.append(('OR', token))
        elif token == 'NOT':
            parsed_tokens.append(('NOT', token))
        elif token == '(':
            parsed_tokens.append(('LPAREN', token))
        elif token == ')':
            parsed_tokens.append(('RPAREN', token))
        else:
            parsed_tokens.append(('TERM', token))
    
    return parsed_tokens

def evaluate_boolean_expression(tokens: List[Tuple[str, str]], document_text: str, case_sensitive: bool = False) -> bool:
    """
    Evaluate a boolean expression against document text.
    
    Args:
        tokens: List of (token_type, value) tuples
        document_text: Full text of the document
        case_sensitive: Whether search should be case sensitive
        
    Returns:
        bool: True if the document matches the boolean expression
    """
    if not case_sensitive:
        document_text = document_text.lower()
    
    def term_exists(term: str) -> bool:
        """Check if a term exists in the document (exact or partial matching based on wildcard)"""
        search_term = term if case_sensitive else term.lower()
        
        # Check if term ends with * for partial matching
        if search_term.endswith('*'):
            # Remove the * and do partial matching
            search_term = search_term[:-1]
            return search_term in document_text
        else:
            # Exact word matching using word boundaries
            pattern = r'\b' + re.escape(search_term) + r'\b'
            return bool(re.search(pattern, document_text, re.IGNORECASE if not case_sensitive else 0))
    
    def parse_expression(pos: int) -> Tuple[bool, int]:
        """Parse and evaluate expression starting at position pos"""
        result, pos = parse_or_expression(pos)
        return result, pos
    
    def parse_or_expression(pos: int) -> Tuple[bool, int]:
        """Parse OR expressions (lowest precedence)"""
        result, pos = parse_and_expression(pos)
        
        while pos < len(tokens) and tokens[pos][0] == 'OR':
            pos += 1  # Skip OR
            right, pos = parse_and_expression(pos)
            result = result or right
            
        return result, pos
    
    def parse_and_expression(pos: int) -> Tuple[bool, int]:
        """Parse AND expressions (higher precedence than OR)"""
        result, pos = parse_not_expression(pos)
        
        while pos < len(tokens) and tokens[pos][0] == 'AND':
            pos += 1  # Skip AND
            right, pos = parse_not_expression(pos)
            result = result and right
            
        return result, pos
    
    def parse_not_expression(pos: int) -> Tuple[bool, int]:
        """Parse NOT expressions (highest precedence)"""
        if pos < len(tokens) and tokens[pos][0] == 'NOT':
            pos += 1  # Skip NOT
            result, pos = parse_primary(pos)
            return not result, pos
        else:
            return parse_primary(pos)
    
    def parse_primary(pos: int) -> Tuple[bool, int]:
        """Parse primary expressions (terms and parentheses)"""
        if pos >= len(tokens):
            return False, pos
            
        token_type, value = tokens[pos]
        
        if token_type == 'TERM':
            return term_exists(value), pos + 1
        elif token_type == 'LPAREN':
            pos += 1  # Skip (
            result, pos = parse_expression(pos)
            if pos < len(tokens) and tokens[pos][0] == 'RPAREN':
                pos += 1  # Skip )
            return result, pos
        else:
            return False, pos
    
    if not tokens:
        return False
    
    result, _ = parse_expression(0)
    return result

def search_pdf_for_boolean_text(pdf_path: str, search_query: str, case_sensitive: bool = False) -> List[Dict]:
    """
    Search a PDF file for text using boolean operators and return matches with page numbers.
    
    Args:
        pdf_path (str): Path to the PDF file
        search_query (str): Boolean search query
        case_sensitive (bool): Whether the search should be case-sensitive
        
    Returns:
        list: List of dictionaries with page_number, match_context, and file_path
    """
    matches = []
    
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            num_pages = len(reader.pages)
            
            # Extract all text from the document for boolean evaluation
            full_document_text = ""
            page_texts = []
            
            for page_num in range(num_pages):
                page = reader.pages[page_num]
                page_text = page.extract_text() or ""
                page_texts.append(page_text)
                full_document_text += page_text + "\n"
            
            # Parse the boolean query
            tokens = parse_boolean_query(search_query)
            
            # Check if the document matches the boolean expression
            if evaluate_boolean_expression(tokens, full_document_text, case_sensitive):
                # If document matches, find which pages contain any of the search terms
                search_terms = [token[1] for token in tokens if token[0] == 'TERM']
                
                for page_num, page_text in enumerate(page_texts):
                    if page_text:
                        page_text_search = page_text if case_sensitive else page_text.lower()
                        
                        # Check if any search term appears on this page
                        for term in search_terms:
                            term_search = term if case_sensitive else term.lower()
                            
                            # Handle wildcard vs exact matching
                            if term.endswith('*'):
                                # Partial matching - remove * and search for substring
                                search_pattern = term_search[:-1]
                                matches_found = list(re.finditer(re.escape(search_pattern), page_text_search))
                            else:
                                # Exact word matching using word boundaries
                                pattern = r'\b' + re.escape(term_search) + r'\b'
                                matches_found = list(re.finditer(pattern, page_text_search, re.IGNORECASE if not case_sensitive else 0))
                            
                            # Find all occurrences of this term on this page
                            for match in matches_found:
                                start_pos = max(0, match.start() - 50)
                                end_pos = min(len(page_text_search), match.end() + 50)
                                
                                # Get context (text before and after the match)
                                context = "..." + page_text[start_pos:end_pos].replace('\n', ' ') + "..."
                                
                                # Store match information
                                matches.append({
                                    'file_path': pdf_path,
                                    'page_number': page_num + 1,
                                    'matched_term': term,
                                    'match_context': context,
                                    'search_query': search_query
                                })
                                break  # Only one match per term per page needed for context
    
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
    
    return matches

def search_pdf_for_text(pdf_path, search_text, case_sensitive=False):
    """
    Legacy function for backward compatibility. Uses boolean search internally.
    """
    # If no boolean operators detected, treat as simple term search
    if not any(op in search_text.upper() for op in ['AND', 'OR', 'NOT']):
        return search_pdf_for_boolean_text(pdf_path, search_text, case_sensitive)
    else:
        return search_pdf_for_boolean_text(pdf_path, search_text, case_sensitive)

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
    parser = argparse.ArgumentParser(description='Search PDF files for specific text using Boolean operators and generate Excel report')
    parser.add_argument('search_text', help='Text to search for in PDF files. Supports Boolean operators: AND, OR, NOT. Use parentheses for grouping. Example: "term1 AND (term2 OR term3) AND NOT term4"')
    parser.add_argument('directory', help='Directory containing PDF files to search')
    parser.add_argument('--output', '-o', help='Output file path for Excel report')
    parser.add_argument('--case-sensitive', '-c', action='store_true', help='Enable case-sensitive search')

    args = parser.parse_args()

    print(f"Searching for '{args.search_text}' in {args.directory}...")
    print("Boolean operators supported: AND, OR, NOT (case-sensitive)")
    print("Use parentheses for grouping: (term1 OR term2) AND term3")
    print("Wildcard matching: 'term*' for partial match, 'term' for exact word match")
    print("Examples: 'his*' matches 'this', 'histogram'; 'this' matches only 'this'")
    print("-" * 60)

    # Perform search
    matches = search_directory_for_pdfs(args.directory, args.search_text, args.case_sensitive)

    # Prepare output filename
    safe_search = re.sub(r'[^A-Za-z0-9]+', '_', args.search_text)[:40]
    date_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_file = args.output if args.output else os.path.join(script_dir, f"results_{safe_search}_{date_str}.xlsx")

    # Create Excel report
    create_excel_report(matches, output_file, args.search_text)

    if matches:
        print(f"Excel report created: {output_file}")

if __name__ == "__main__":
    main()
