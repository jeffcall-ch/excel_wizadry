import camelot
import pandas as pd
import logging
from pathlib import Path
import fitz  # PyMuPDF
import logging
logging.getLogger('camelot').setLevel(logging.WARNING)
logging.getLogger('pdfminer').setLevel(logging.WARNING)
logging.getLogger('pdfplumber').setLevel(logging.WARNING)

# Custom Exceptions for Camelot Extraction
class CamelotExtractionError(Exception):
    """Raised when Camelot fails to extract tables"""
    pass

class InvalidTableBoundsError(Exception):
    """Raised when table boundaries are invalid or malformed"""
    pass

class PDFPageAccessError(Exception):
    """Raised when unable to access PDF page for dimension calculation"""
    pass

class EmptyTableError(Exception):
    """Raised when extracted table is empty after cleaning"""
    pass

def extract_table_with_camelot(pdf_path: str, table_bounds: tuple, page_num: int = 1):
    """
    Extract table content from a PDF file using Camelot in stream mode.

    Args:
        pdf_path (str): Path to the PDF file.
        table_bounds (tuple): Table boundaries (x0, y0, x1, y1) from table_boundary_finder.
        page_num (int): Page number to extract from (1-based).

    Returns:
        pandas.DataFrame: Extracted table data.
        
    Raises:
        InvalidTableBoundsError: If table boundaries are invalid.
        PDFPageAccessError: If unable to access the PDF page.
        CamelotExtractionError: If Camelot fails to extract tables.
        EmptyTableError: If no data is extracted from the table.
    """
    logging.debug(f"Extracting table from PDF: {pdf_path}, Page: {page_num}")
    
    # Validate table boundaries
    if not table_bounds or len(table_bounds) != 4:
        raise InvalidTableBoundsError(f"Invalid table boundaries: expected 4 coordinates, got {table_bounds}")
    
    if not all(isinstance(coord, (int, float)) for coord in table_bounds):
        raise InvalidTableBoundsError(f"Table boundaries must be numeric: {table_bounds}")
    
    # Get PDF page dimensions for coordinate conversion
    try:
        with fitz.open(pdf_path) as doc:
            if page_num - 1 >= len(doc) or page_num < 1:
                raise PDFPageAccessError(f"Invalid page number {page_num} for PDF with {len(doc)} pages")
            
            page = doc[page_num - 1]  # Convert to 0-based indexing
            page_rect = page.rect
            page_height = page_rect.height
    except Exception as e:
        raise PDFPageAccessError(f"Error reading PDF dimensions: {str(e)}")
    
    # Convert coordinates from table_boundary_finder format to Camelot format
    x0, y0, x1, y1 = table_bounds
    
    # Convert to PDF coordinates (where y=0 is bottom, y increases upward):
    left_x = x0      # Left edge unchanged
    right_x = x1     # Right edge unchanged  
    top_y_pdf = page_height - y0     # Top in PDF = page_height - top_in_pymupdf
    bottom_y_pdf = page_height - y1  # Bottom in PDF = page_height - bottom_in_pymupdf
    
    # Camelot expects: "x1,y1,x2,y2" where (x1,y1)=left-top, (x2,y2)=right-bottom  
    bbox = f"{left_x},{top_y_pdf},{right_x},{bottom_y_pdf}"
    
    logging.debug(f"PyMuPDF bounds: left={x0}, top={y0}, right={x1}, bottom={y1}")
    logging.debug(f"Page height: {page_height}")
    logging.debug(f"PDF coords: left-top=({left_x}, {top_y_pdf}), right-bottom=({right_x}, {bottom_y_pdf})")
    logging.debug(f"Camelot bbox: {bbox}")

    # Try STREAM parser first (better for tables without clear grid lines)
    df = None
    extraction_method = None
    
    try:
        logging.debug("Attempting stream parser...")
        tables = camelot.read_pdf(
            pdf_path, 
            pages=str(page_num),
            flavor='stream', 
            table_areas=[bbox]
        )
        
        if tables.n > 0:
            df = tables[0].df
            if not df.empty:
                logging.info(f"Stream parser succeeded with {len(df)} rows and {len(df.columns)} columns")
                extraction_method = "stream"
        
    except Exception as e:
        logging.warning(f"Stream parser failed: {e}")
    
    # Fallback to LATTICE parser if stream fails
    if df is None or df.empty:
        try:
            logging.debug("Attempting lattice parser as fallback...")
            tables = camelot.read_pdf(
                pdf_path, 
                pages=str(page_num),
                flavor='lattice', 
                table_areas=[bbox]
            )
            
            if tables.n > 0:
                df = tables[0].df
                if not df.empty:
                    logging.info(f"Lattice parser succeeded with {len(df)} rows and {len(df.columns)} columns")
                    extraction_method = "lattice"
            
        except Exception as e:
            logging.warning(f"Lattice parser also failed: {e}")
    
    # Check if we got any data
    if df is None or df.empty:
        raise EmptyTableError(f"No tables found in the specified area on page {page_num}")
    
    logging.debug(f"Extraction method used: {extraction_method}")
    logging.debug(f"Raw table shape: {df.shape}")
    
    # Clean up the DataFrame
    df_cleaned = clean_extracted_dataframe(df)
    
    logging.info(f"Successfully extracted table with {len(df_cleaned)} rows and {len(df_cleaned.columns)} columns")
    logging.debug(f"Final table columns: {list(df_cleaned.columns)}")
    
    return df_cleaned

def clean_extracted_dataframe(df):
    """
    Clean and process the extracted dataframe.
    
    Args:
        df (pandas.DataFrame): Raw extracted dataframe
        
    Returns:
        pandas.DataFrame: Cleaned dataframe
    """
    # Remove completely empty rows and columns first
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    if df.empty:
        raise EmptyTableError("Table became empty after removing blank rows/columns")
    
    # Reset index to ensure clean output
    df = df.reset_index(drop=True)
    
    # Fix column names - handle empty or problematic column names
    new_columns = []
    for i, col in enumerate(df.columns):
        col_str = str(col).strip()
        if not col_str or col_str in ['', 'nan', 'None', 'NaN']:
            new_columns.append(f'Column_{i}')
        else:
            # Clean up column names - remove newlines and extra spaces
            clean_col = col_str.replace('\n', ' ').replace('\r', ' ')
            clean_col = ' '.join(clean_col.split())  # Normalize whitespace
            new_columns.append(clean_col[:50])  # Limit column name length
    
    df.columns = new_columns
    
    # Clean cell values safely
    for col in df.columns:
        try:
            # Convert column to string, handling different data types
            df[col] = df[col].apply(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != 'nan' else '')
            # Replace empty strings with NaN
            df[col] = df[col].replace(['', 'nan', 'None', 'NaN'], pd.NA)
        except Exception as e:
            logging.warning(f"Could not clean column '{col}': {e}")
            # Fallback: just convert to string
            df[col] = df[col].astype(str)
    
    # Final cleanup - remove any rows that became completely empty
    df = df.dropna(how='all').reset_index(drop=True)
    
    if df.empty:
        raise EmptyTableError("Table became empty after data cleaning")
    
    return df

def extract_table_stream_only(pdf_path: str, table_bounds: tuple, page_num: int = 1):
    """
    Extract table using only the stream parser (no lattice fallback).
    
    Args:
        pdf_path (str): Path to the PDF file.
        table_bounds (tuple): Table boundaries (x0, y0, x1, y1) from table_boundary_finder.
        page_num (int): Page number to extract from (1-based).
        
    Returns:
        pandas.DataFrame: Extracted table data or raises exception.
    """
    logging.debug(f"Stream-only extraction from: {pdf_path}, Page: {page_num}")
    
    # Validate inputs
    if not table_bounds or len(table_bounds) != 4:
        raise InvalidTableBoundsError(f"Invalid table boundaries: expected 4 coordinates, got {table_bounds}")
    
    # Get page dimensions
    try:
        with fitz.open(pdf_path) as doc:
            page = doc[page_num - 1]
            page_height = page.rect.height
    except Exception as e:
        raise PDFPageAccessError(f"Error reading PDF dimensions: {str(e)}")
    
    # Convert coordinates
    x0, y0, x1, y1 = table_bounds
    bbox = f"{x0},{page_height - y0},{x1},{page_height - y1}"
    
    # Extract with stream parser only
    try:
        tables = camelot.read_pdf(
            pdf_path, 
            pages=str(page_num),
            flavor='stream', 
            table_areas=[bbox]
        )
        
        if tables.n == 0:
            raise EmptyTableError(f"Stream parser found no tables in the specified area on page {page_num}")
        
        df = tables[0].df
        if df.empty:
            raise EmptyTableError(f"Stream parser extracted empty table on page {page_num}")
        
        logging.info(f"Stream parser extracted {len(df)} rows and {len(df.columns)} columns")
        return clean_extracted_dataframe(df)
        
    except Exception as e:
        if isinstance(e, (EmptyTableError, CamelotExtractionError)):
            raise
        else:
            raise CamelotExtractionError(f"Stream parser extraction failed: {str(e)}")

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Extract tables from a PDF using Camelot stream parser.")
    parser.add_argument("pdf_path", type=str, help="Path to the PDF file.")
    parser.add_argument("x0", type=float, help="Left boundary of the table.")
    parser.add_argument("y0", type=float, help="Top boundary of the table.")
    parser.add_argument("x1", type=float, help="Right boundary of the table.")
    parser.add_argument("y1", type=float, help="Bottom boundary of the table.")
    parser.add_argument("--page", type=int, default=1, help="Page number to extract from (default: 1)")
    parser.add_argument("--stream-only", action="store_true", help="Use only stream parser (no lattice fallback)")

    args = parser.parse_args()

    pdf_path = args.pdf_path
    table_bounds = (args.x0, args.y0, args.x1, args.y1)

    # Configure logging for standalone usage
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

    if not Path(pdf_path).is_file():
        logging.error(f"PDF file not found: {pdf_path}")
    else:
        try:
            if args.stream_only:
                result_df = extract_table_stream_only(pdf_path, table_bounds, args.page)
            else:
                result_df = extract_table_with_camelot(pdf_path, table_bounds, args.page)
            
            if result_df is not None and not result_df.empty:
                print("Extracted table:")
                print(result_df.to_string())
                
                # Save to CSV for inspection
                output_file = f"extracted_table_stream_page_{args.page}.csv"
                result_df.to_csv(output_file, index=False)
                print(f"\nTable saved to: {output_file}")
            else:
                print("No table data extracted")
                
        except Exception as e:
            logging.error(f"Extraction failed: {e}")
            print(f"Error: {e}")