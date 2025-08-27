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
    Extract table content from a PDF file using Camelot in lattice mode.

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
    
    # Your boundary finder returns: (min_x0, min_y0, max_x1, max_y1) in PyMuPDF coords
    # where: min_y0 = top of table, max_y1 = bottom of table (in PyMuPDF where y=0 is top)
    
    # Convert to PDF coordinates (where y=0 is bottom, y increases upward):
    left_x = x0      # Left edge unchanged
    right_x = x1     # Right edge unchanged  
    top_y_pdf = page_height - y0     # Top in PDF = page_height - top_in_pymupdf
    bottom_y_pdf = page_height - y1  # Bottom in PDF = page_height - bottom_in_pymupdf
    
    # Camelot expects: "x1,y1,x2,y2" where (x1,y1)=left-top, (x2,y2)=right-bottom  
    # In PDF space, top_y > bottom_y because y-axis points upward
    bbox = f"{left_x},{top_y_pdf},{right_x},{bottom_y_pdf}"
    
    logging.debug(f"PyMuPDF bounds: left={x0}, top={y0}, right={x1}, bottom={y1}")
    logging.debug(f"Page height: {page_height}")
    logging.debug(f"PDF coords: left-top=({left_x}, {top_y_pdf}), right-bottom=({right_x}, {bottom_y_pdf})")
    logging.debug(f"Camelot bbox: {bbox}")

    # Extract tables using Camelot in lattice mode
    try:
        tables = camelot.read_pdf(
            pdf_path, 
            pages=str(page_num),  # Specify the page number
            flavor='lattice', 
            table_areas=[bbox]
        )
        
        if tables.n == 0:
            raise EmptyTableError(f"No tables found in the specified area on page {page_num}")
        
        logging.debug(f"Number of tables found: {tables.n}")
        
        # Get the first (and hopefully only) table
        table = tables[0]
        df = table.df
        
        # Clean up the DataFrame
        if df.empty:
            raise EmptyTableError(f"Extracted table is empty on page {page_num}")
            
        # Remove completely empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        if df.empty:
            raise EmptyTableError(f"Table became empty after removing blank rows/columns on page {page_num}")
        
        # Reset index to ensure clean CSV output
        df = df.reset_index(drop=True)
        
        # Clean up column names (use first row as headers if it looks like headers)
        if len(df) > 0:
            # Check if first row contains header-like text (non-numeric)
            first_row = df.iloc[0]
            if any(isinstance(val, str) and val.strip() and not val.replace('.', '').replace('-', '').isdigit() for val in first_row):
                # Use first row as column names
                df.columns = [str(val).strip() if pd.notna(val) else f'Column_{i}' for i, val in enumerate(first_row)]
                df = df.iloc[1:].reset_index(drop=True)  # Remove the header row from data
            else:
                # Generate generic column names
                df.columns = [f'Column_{i}' for i in range(len(df.columns))]
        
        # Clean up cell values
        for col in df.columns:
            if col in df.columns:  # Check if column still exists after cleaning
                # Handle different data types safely
                try:
                    # Convert to string and clean
                    df[col] = df[col].astype(str).str.strip()
                    # Replace various empty representations with NaN
                    df[col] = df[col].replace(['', 'nan', 'None', 'NaN'], pd.NA)
                except Exception as e:
                    logging.warning(f"Could not clean column '{col}': {e}")
                    # If cleaning fails, just convert to string
                    df[col] = df[col].astype(str)
        
        # Final cleanup - remove any rows that became empty after string cleaning
        df = df.dropna(how='all').reset_index(drop=True)
        
        if df.empty:
            raise EmptyTableError(f"Table became empty after data cleaning on page {page_num}")
        
        logging.info(f"Successfully extracted table with {len(df)} rows and {len(df.columns)} columns")
        logging.debug(f"Table columns: {list(df.columns)}")
        
        return df
            
    except Exception as e:
        if isinstance(e, (EmptyTableError, CamelotExtractionError)):
            raise  # Re-raise our custom exceptions
        else:
            raise CamelotExtractionError(f"Camelot extraction failed: {str(e)}")

def extract_table_stream_fallback(pdf_path: str, table_bounds: tuple, page_num: int = 1):
    """
    Fallback method using Camelot's stream parser if lattice fails.
    
    Args:
        pdf_path (str): Path to the PDF file.
        table_bounds (tuple): Table boundaries (x0, y0, x1, y1) from table_boundary_finder.
        page_num (int): Page number to extract from (1-based).
        
    Returns:
        pandas.DataFrame: Extracted table data or None if extraction fails.
    """
    logging.debug(f"Trying stream parser fallback for: {pdf_path}, Page: {page_num}")
    
    try:
        # Use the same coordinate conversion as lattice method
        with fitz.open(pdf_path) as doc:
            page = doc[page_num - 1]
            page_height = page.rect.height
        
        x0, y0, x1, y1 = table_bounds
        pdf_x0 = x0
        pdf_y0 = page_height - y1
        pdf_x1 = x1
        pdf_y1 = page_height - y0
        bbox = f"{pdf_x0},{pdf_y1},{pdf_x1},{pdf_y0}"
        
        # Try stream parser
        tables = camelot.read_pdf(
            pdf_path, 
            pages=str(page_num),
            flavor='stream', 
            table_areas=[bbox]
        )
        
        if tables.n > 0:
            df = tables[0].df
            if not df.empty:
                logging.info(f"Stream parser succeeded with {len(df)} rows")
                return df
        
        return pd.DataFrame()
        
    except Exception as e:
        logging.error(f"Stream parser also failed: {e}")
        return pd.DataFrame()

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Extract tables from a PDF using Camelot.")
    parser.add_argument("pdf_path", type=str, help="Path to the PDF file.")
    parser.add_argument("x0", type=float, help="Left boundary of the table.")
    parser.add_argument("y0", type=float, help="Top boundary of the table.")
    parser.add_argument("x1", type=float, help="Right boundary of the table.")
    parser.add_argument("y1", type=float, help="Bottom boundary of the table.")
    parser.add_argument("--page", type=int, default=1, help="Page number to extract from (default: 1)")

    args = parser.parse_args()

    pdf_path = args.pdf_path
    table_bounds = (args.x0, args.y0, args.x1, args.y1)

    # Configure logging for standalone usage
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

    if not Path(pdf_path).is_file():
        logging.error(f"PDF file not found: {pdf_path}")
    else:
        result_df = extract_table_with_camelot(pdf_path, table_bounds, args.page)
        if result_df is not None and not result_df.empty:
            print("Extracted table:")
            print(result_df.to_string())
            
            # Save to CSV for inspection
            output_file = f"extracted_table_page_{args.page}.csv"
            result_df.to_csv(output_file, index=False)
            print(f"\nTable saved to: {output_file}")
        else:
            print("No table data extracted")