import camelot
import pandas as pd
import logging
from pathlib import Path
import fitz  # PyMuPDF
import logging
logging.getLogger('camelot').setLevel(logging.WARNING)
logging.getLogger('pdfminer').setLevel(logging.WARNING)
logging.getLogger('pdfplumber').setLevel(logging.WARNING)

# -----------------------
# Custom Exceptions for Camelot Extraction - Improved Design
# -----------------------

class CamelotExtractionError(Exception):
    """Base exception for Camelot table extraction errors"""
    pass

class PDFDimensionError(CamelotExtractionError):
    """Base class for PDF dimension-related errors"""
    pass

class PDFPageNotAccessibleError(PDFDimensionError):
    """Raised when unable to access PDF page for dimension calculation"""
    pass

class PDFPageCountMismatchError(PDFDimensionError):
    """Raised when requested page number exceeds available pages"""
    pass

class TableBoundsError(CamelotExtractionError):
    """Base class for table boundary validation errors"""
    pass

class InvalidTableBoundsFormatError(TableBoundsError):
    """Raised when table boundaries are invalid or malformed"""
    pass

class NonNumericTableBoundsError(TableBoundsError):
    """Raised when table boundaries contain non-numeric values"""
    pass

class ParserError(CamelotExtractionError):
    """Base class for parser-related errors"""
    pass

class StreamParserFailedError(ParserError):
    """Raised when Camelot stream parser fails to extract tables"""
    pass

class LatticeParserFailedError(ParserError):
    """Raised when Camelot lattice parser fails to extract tables"""
    pass

class AllParsersFailedError(ParserError):
    """Raised when both stream and lattice parsers fail"""
    pass

class TableDataError(CamelotExtractionError):
    """Base class for table data processing errors"""
    pass

class NoTablesDetectedError(TableDataError):
    """Raised when no tables are detected in the specified area"""
    pass

class EmptyTableExtractedError(TableDataError):
    """Raised when extracted table is empty after processing"""
    pass

class TableCleaningError(TableDataError):
    """Raised when table cleaning/processing fails"""
    pass

class CSVExportError(CamelotExtractionError):
    """Raised when CSV export fails"""
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
        InvalidTableBoundsFormatError: If table boundaries format is invalid.
        NonNumericTableBoundsError: If table boundaries contain non-numeric values.
        PDFPageNotAccessibleError: If unable to access the PDF page.
        PDFPageCountMismatchError: If page number is invalid.
        AllParsersFailedError: If both parsers fail to extract tables.
        NoTablesDetectedError: If no data is extracted from the table.
    """
    logging.debug(f"Extracting table from PDF: {pdf_path}, Page: {page_num}")
    
    # Validate table boundaries format
    if not table_bounds or len(table_bounds) != 4:
        raise InvalidTableBoundsFormatError(f"Invalid table boundaries format: expected 4 coordinates, got {table_bounds}")
    
    if not all(isinstance(coord, (int, float)) for coord in table_bounds):
        raise NonNumericTableBoundsError(f"Table boundaries must be numeric values: {table_bounds}")
    
    # Get PDF page dimensions for coordinate conversion
    try:
        with fitz.open(pdf_path) as doc:
            if page_num - 1 >= len(doc) or page_num < 1:
                raise PDFPageCountMismatchError(f"Invalid page number {page_num} for PDF with {len(doc)} pages")
            
            page = doc[page_num - 1]  # Convert to 0-based indexing
            page_rect = page.rect
            page_height = page_rect.height
    except fitz.FileDataError as e:
        raise PDFPageNotAccessibleError(f"Cannot access PDF file {pdf_path}: {str(e)}")
    except fitz.FileNotFoundError as e:
        raise PDFPageNotAccessibleError(f"PDF file not found: {pdf_path}")
    except Exception as e:
        raise PDFPageNotAccessibleError(f"Error reading PDF dimensions: {str(e)}")
    
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
    stream_error = None
    lattice_error = None
    
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
        stream_error = f"Stream parser failed: {str(e)}"
        logging.warning(stream_error)
    
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
            lattice_error = f"Lattice parser failed: {str(e)}"
            logging.warning(lattice_error)
    
    # Check if we got any data
    if df is None or df.empty:
        error_details = []
        if stream_error:
            error_details.append(stream_error)
        if lattice_error:
            error_details.append(lattice_error)
        
        if error_details:
            combined_error = "; ".join(error_details)
            raise AllParsersFailedError(f"Both parsers failed on page {page_num}: {combined_error}")
        else:
            raise NoTablesDetectedError(f"No tables found in the specified area on page {page_num}")
    
    logging.debug(f"Extraction method used: {extraction_method}")
    logging.debug(f"Raw table shape: {df.shape}")
    
    # Clean up the DataFrame
    try:
        df_cleaned = clean_extracted_dataframe(df)
    except Exception as e:
        raise TableCleaningError(f"Failed to clean extracted table data: {str(e)}")
    
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
        
    Raises:
        EmptyTableExtractedError: If table becomes empty after cleaning.
        TableCleaningError: If cleaning process encounters errors.
    """
    try:
        # Remove completely empty rows and columns first
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        if df.empty:
            raise EmptyTableExtractedError("Table became empty after removing blank rows/columns")
        
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
            raise EmptyTableExtractedError("Table became empty after data cleaning")
        
        return df
        
    except EmptyTableExtractedError:
        raise  # Re-raise our custom exception
    except Exception as e:
        raise TableCleaningError(f"Unexpected error during table cleaning: {str(e)}")

def extract_table_stream_only(pdf_path: str, table_bounds: tuple, page_num: int = 1):
    """
    Extract table using only the stream parser (no lattice fallback).
    
    Args:
        pdf_path (str): Path to the PDF file.
        table_bounds (tuple): Table boundaries (x0, y0, x1, y1) from table_boundary_finder.
        page_num (int): Page number to extract from (1-based).
        
    Returns:
        pandas.DataFrame: Extracted table data or raises exception.
        
    Raises:
        InvalidTableBoundsFormatError: If table boundaries format is invalid.
        PDFPageNotAccessibleError: If unable to access PDF page.
        StreamParserFailedError: If stream parser fails.
        EmptyTableExtractedError: If extracted table is empty.
    """
    logging.debug(f"Stream-only extraction from: {pdf_path}, Page: {page_num}")
    
    # Validate inputs
    if not table_bounds or len(table_bounds) != 4:
        raise InvalidTableBoundsFormatError(f"Invalid table boundaries format: expected 4 coordinates, got {table_bounds}")
    
    # Get page dimensions
    try:
        with fitz.open(pdf_path) as doc:
            if page_num - 1 >= len(doc) or page_num < 1:
                raise PDFPageCountMismatchError(f"Invalid page number {page_num} for PDF with {len(doc)} pages")
            page = doc[page_num - 1]
            page_height = page.rect.height
    except Exception as e:
        raise PDFPageNotAccessibleError(f"Error reading PDF dimensions: {str(e)}")
    
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
            raise NoTablesDetectedError(f"Stream parser found no tables in the specified area on page {page_num}")
        
        df = tables[0].df
        if df.empty:
            raise EmptyTableExtractedError(f"Stream parser extracted empty table on page {page_num}")
        
        logging.info(f"Stream parser extracted {len(df)} rows and {len(df.columns)} columns")
        return clean_extracted_dataframe(df)
        
    except (NoTablesDetectedError, EmptyTableExtractedError):
        raise  # Re-raise our custom exceptions
    except Exception as e:
        raise StreamParserFailedError(f"Stream parser extraction failed: {str(e)}")


def export_to_csv_safe(df, output_path):
    """
    Export DataFrame to CSV with safe formatting for Excel compatibility.
    
    Args:
        df (pandas.DataFrame): DataFrame to export
        output_path (str): Path for output CSV file
        
    Returns:
        pandas.DataFrame: The exported dataframe
        
    Raises:
        CSVExportError: If CSV export fails
    """
    try:
        # Clone the dataframe to avoid modifying original
        df_export = df.copy()
        
        # Ensure all data is single-line
        for col in df_export.columns:
            if df_export[col].dtype == 'object':  # String columns
                df_export[col] = df_export[col].astype(str)
                df_export[col] = df_export[col].str.replace('\n', ' ').str.replace('\r', ' ')
                df_export[col] = df_export[col].str.replace(r'\s+', ' ', regex=True)
        
        # Export with proper CSV parameters
        df_export.to_csv(
            output_path, 
            index=False, 
            encoding='utf-8-sig',  # BOM for Excel compatibility
            lineterminator='\n',   # Consistent line endings
            quoting=1              # Quote all text fields
        )
        
        return df_export
        
    except Exception as e:
        raise CSVExportError(f"Failed to export DataFrame to CSV '{output_path}': {str(e)}")