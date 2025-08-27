#!/usr/bin/env python3
"""
MVP PDF BOM Extractor - Single Process, Simple CSV Output
Updated with custom exceptions while preserving original logic
"""

import os
import csv
import fitz  # PyMuPDF
from pathlib import Path
from typing import List, Tuple, NamedTuple
import logging
logging.getLogger('camelot').setLevel(logging.WARNING)
logging.getLogger('pdfminer').setLevel(logging.WARNING)
logging.getLogger('pdfplumber').setLevel(logging.WARNING)

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# -----------------------
# Custom Exceptions
# -----------------------

class TableBoundaryError(Exception):
    """Base exception for table boundary detection errors"""
    pass

class AnchorTextNotFoundError(TableBoundaryError):
    """Raised when anchor text is not found on the page"""
    pass

class TableStructureError(TableBoundaryError):
    """Raised when table structure cannot be detected"""
    pass

class PDFProcessingError(TableBoundaryError):
    """Raised when PDF cannot be processed"""
    pass

class PageNotFoundError(TableBoundaryError):
    """Raised when specified page doesn't exist in PDF"""
    pass

class KKSCodeNotFoundError(TableBoundaryError):
    """Raised when KKS code is not found on the page"""
    pass

class KKSSUCodeNotFoundError(TableBoundaryError):
    """Raised when KKS/SU code is not found on the page"""
    pass

# -----------------------
# Data Structures
# -----------------------

class TableCell:
    def __init__(self, text: str, x0: float, y0: float, x1: float, y1: float):
        self.text = text
        self.x0, self.y0, self.x1, self.y1 = x0, y0, x1, y1

    @property
    def center_x(self): return (self.x0 + self.x1) / 2
    @property
    def center_y(self): return (self.y0 + self.y1) / 2

class ColumnDefinition(NamedTuple):
    name: str
    x0: float
    x1: float

# -----------------------
# Core Extraction Logic
# -----------------------

def find_anchor_position(page_text_dict: dict, anchor_text: str):
    """Find first occurrence of anchor text."""
    normalized_anchor = anchor_text.upper()
    for block in page_text_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                if normalized_anchor in span["text"].upper():
                    return span  # return first matching span
    return None


def detect_table_structure(page_dict: dict, anchor_text: str):
    """
    Find the anchor text and detect table header text based on bounding box criteria.
    """
    # Find anchor text
    anchor = find_anchor_position(page_dict, anchor_text)
    if not anchor:
        logging.debug(f"Anchor '{anchor_text}' not found on the page.")
        raise AnchorTextNotFoundError(f"Anchor text '{anchor_text}' not found on the page")

    anchor_bbox = anchor["bbox"]
    anchor_x0 = anchor_bbox[0]  # x0 of the anchor
    anchor_x1 = anchor_bbox[2]  # x1 of the anchor
    anchor_y0 = anchor_bbox[1]  # y0 of the anchor
    anchor_y1 = anchor_bbox[3]  # y1 of the anchor

    # Calculate anchor height and vertical midpoint
    anchor_height = anchor_y1 - anchor_y0
    anchor_vertical_mid = anchor_y0 + anchor_height / 2
    tolerance = 0.2 * anchor_height  # 20% of anchor height

    logging.debug("Anchor Detection")
    logging.debug(f"  - Anchor Text: '{anchor_text}'")
    logging.debug(f"  - Anchor Position: {anchor_bbox}")
    logging.debug(f"  - Anchor Height: {anchor_height}")
    logging.debug(f"  - Anchor Vertical Midpoint: {anchor_vertical_mid}")
    logging.debug(f"  - Tolerance: {tolerance}")

    # Define the data structure for text elements
    text_elements = []
    header_top_y = anchor_y0  # Initialize with the bottom of the anchor
    header_left_x = anchor_x0  # Initialize with the left of the anchor
    header_right_x = anchor_x1  # Initialize with the right of the anchor

    # Find qualifying text elements
    logging.debug("Header Text Detection")
    for block in page_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"]
                bbox = span["bbox"]
                text_x0, text_y0, text_x1, text_y1 = bbox

                # Calculate text vertical midpoint
                text_mid = (text_y0 + text_y1) / 2

                # Check if the text is to the right of the anchor and vertically aligned
                if (
                    text_x0 > anchor_x1 and (
                        abs(text_mid - anchor_vertical_mid) <= tolerance or
                        (anchor_y0 < text_y0 < anchor_y1) or
                        (anchor_y0 < text_y1 < anchor_y1)
                    )
                ):
                    text_elements.append({"text": text, "bbox": bbox})
                    header_top_y = min(header_top_y, text_y0)  # Update header_top_y
                    header_right_x = max(header_right_x, text_x1)  # Update header_right_x
                    logging.debug(f"  - Text: '{text}'")
                    logging.debug(f"    BBox: {bbox}")

    # Debug: Calculate average character length for specific headers within text_elements
    headers_to_check = ["POS", "NUMBER", "TOTAL"]
    avg_char_lengths = []
    header_bboxes = {}
    for header in headers_to_check:
        for element in text_elements:
            stripped_text = element["text"].strip()  # Strip front and back whitespaces
            if header in stripped_text:
                bbox = element["bbox"]
                text_width = bbox[2] - bbox[0]  # x1 - x0
                avg_char_length = text_width / len(stripped_text)
                avg_char_lengths.append(avg_char_length)
                header_bboxes[header] = bbox
                logging.debug(f"Header: '{header}'")
                logging.debug(f"  - Stripped Text: '{stripped_text}'")
                logging.debug(f"  - BBox: {bbox}")
    
    if not avg_char_lengths:
        raise TableStructureError("No recognizable table headers found (POS, NUMBER, TOTAL)")
        
    avg_char_length = max(avg_char_lengths)
    logging.debug(f"  - Average Character Length (X): {avg_char_length}")

    # Call find_table_bottom to locate the 'TOTAL' below 'WEIGHT'
    table_bottom_y = find_table_bottom(page_dict, "TOTAL")
    if not table_bottom_y:
        logging.debug("Table Bottom Not Found")
        raise TableStructureError("Table bottom marker 'TOTAL' not found below 'WEIGHT' header")
    else:
        logging.debug(f"Table Bottom Found: {table_bottom_y}")
    
    # Add some padding to  bounding boxes to make sure we have all text in the box and possible the table frame too
    header_left_x -= avg_char_length * 2
    header_right_x += avg_char_length * 2
    header_top_y -= anchor_height * 2
    table_bottom_y = table_bottom_y["bbox"][3] + avg_char_length * 2
    

    logging.debug(f"Total Header Text Elements Detected: {len(text_elements)}")
    logging.debug(f"Header TOP Y: {header_top_y}")
    logging.debug(f"Header LEFT X: {header_left_x}")
    logging.debug(f"Header RIGHT X: {header_right_x}")

    return anchor_bbox, text_elements, avg_char_length, (header_top_y, header_left_x, header_right_x, table_bottom_y)

def find_table_bottom(page_dict: dict, target_text: str):
    """
    Find the target text (e.g., 'TOTAL') below the hardcoded 'WEIGHT' header.
    Return the target text and its bounding box in the usual data format.
    """
    logging.debug("Finding Table Bottom")
    logging.debug("  - Hardcoded Header Text: 'WEIGHT'")
    logging.debug(f"  - Target Text: '{target_text}'")

    # Step 1: Locate the hardcoded 'WEIGHT' header
    header_bbox = None
    for block in page_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"]
                bbox = span["bbox"]

                if text.strip() == "WEIGHT":  # Full match for 'WEIGHT'
                    header_bbox = bbox
                    logging.debug(f"  - Hardcoded Header Found: '{text}'")
                    logging.debug(f"    BBox: {bbox}")
                    break
            if header_bbox:
                break
        if header_bbox:
            break

    if not header_bbox:
        logging.debug("Hardcoded Header 'WEIGHT' not found.")
        raise TableStructureError("Required 'WEIGHT' header not found for table bottom detection")

    # Step 2: Define the search area below the hardcoded header
    search_y_min = header_bbox[3]  # Bottom of the header
    header_x0 = header_bbox[0]     # x0 of the header

    # Step 3: Find the target text
    for block in page_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"]
                bbox = span["bbox"]
                text_x0, text_y0 = bbox[0], bbox[1]

                # Check if the text is below the header and matches the target text
                if text_y0 > search_y_min and text_x0 >= header_x0 and text.strip() == target_text:  # Full match only
                    logging.debug(f"  - Target Found: '{text}'")
                    logging.debug(f"    BBox: {bbox}")
                    return {"text": text, "bbox": bbox}

    # Debug statement if TOTAL is not found
    logging.debug(f"Target '{target_text}' not found below hardcoded header 'WEIGHT'.")
    return None

def find_table_content(page_dict: dict, avg_char_length: float, table_bounds: Tuple[float, float, float, float]):
    """
    Extract all text within the table boundaries and print their coordinates.

    Args:
        page_dict (dict): The page dictionary containing text blocks.
        avg_char_length (float): Average character length for padding.
        table_bounds (Tuple[float, float, float, float]): Tuple containing (header_top, header_left, header_right, table_bottom).
        
    Returns:
        Tuple[float, float, float, float]: Final table boundaries (min_x0, min_y0, max_x1, max_y1).
        
    Raises:
        TableStructureError: If no text is found within table boundaries.
    """
    header_top, header_left, header_right, table_bottom = table_bounds

    # Debug print header bounds
    logging.debug("Header Bounds:")
    logging.debug(f"  - Header Top: {header_top}")
    logging.debug(f"  - Header Left: {header_left}")
    logging.debug(f"  - Header Right: {header_right}")
    logging.debug(f"  - Table Bottom: {table_bottom}")

    # List to store text elements within the table boundaries
    table_text_elements = []

    for block in page_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"]
                bbox = span["bbox"]
                text_x0, text_y0, text_x1, text_y1 = bbox

                # Check if the text is within the table boundaries
                if (
                    header_top <= text_y0 <= table_bottom and  # Below header_top and above table_bottom
                    header_left <= text_x0 <= header_right    # Between header_left and header_right
                ):
                    table_text_elements.append({"text": text, "bbox": bbox})

    # Find the minimum and maximum coordinates
    if table_text_elements:
        min_x0_element = min(table_text_elements, key=lambda e: e["bbox"][0])
        max_x1_element = max(table_text_elements, key=lambda e: e["bbox"][2])
        min_y0_element = min(table_text_elements, key=lambda e: e["bbox"][1])
        max_y1_element = max(table_text_elements, key=lambda e: e["bbox"][3])

        min_x0 = min_x0_element["bbox"][0] - 2*avg_char_length
        max_x1 = max_x1_element["bbox"][2] + 2*avg_char_length
        min_y0 = min_y0_element["bbox"][1] - 2*avg_char_length
        max_y1 = max_y1_element["bbox"][3] + 2*avg_char_length

        logging.debug("Table Content:")
        for element in table_text_elements:
            logging.debug(f"  - Text: '{element['text']}'")
            logging.debug(f"    BBox: {element['bbox']}")

        logging.debug("Table Boundary Coordinates:")
        logging.debug(f"  - Minimum X0: {min_x0} (Text: '{min_x0_element['text']}', BBox: {min_x0_element['bbox']})")
        logging.debug(f"  - Maximum X1: {max_x1} (Text: '{max_x1_element['text']}', BBox: {max_x1_element['bbox']})")
        logging.debug(f"  - Minimum Y0: {min_y0} (Text: '{min_y0_element['text']}', BBox: {min_y0_element['bbox']})")
        logging.debug(f"  - Maximum Y1: {max_y1} (Text: '{max_y1_element['text']}', BBox: {max_y1_element['bbox']})")

        return (min_x0, min_y0, max_x1, max_y1)

    else:
        logging.debug("No text found within the table boundaries.")
        raise TableStructureError("No text content found within calculated table boundaries")


def extract_kks_codes_from_page_dict(page_dict: dict):
    """
    Extract KKS codes and KKS codes with "/SU" postfix from a page dictionary.

    Args:
        page_dict (dict): The page dictionary containing text blocks.

    Returns:
        tuple: A tuple containing two lists:
            - List of KKS codes.
            - List of KKS codes with "/SU" postfix.

    Raises:
        ValueError: If no KKS codes are found.
        ValueError: If no KKS codes with "/SU" postfix are found.
    """
    import re

    kks_pattern = r"\b\d[A-Z]{3}\d{2}BQ\d{3}\b"
    kks_su_pattern = r"\b\d[A-Z]{3}\d{2}BQ\d{3}/SU\b"

    kks_codes = []
    kks_su_codes = []

    # Iterate through the text spans in the page dictionary
    for block in page_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "")

                # Find all KKS codes in the text
                kks_codes.extend(re.findall(kks_pattern, text))

                # Find all KKS codes with "/SU" postfix in the text
                kks_su_codes.extend(re.findall(kks_su_pattern, text))

    if not kks_codes:
        raise ValueError("No KKS codes were found.")

    if not kks_su_codes:
        raise ValueError("No KKS codes with '/SU' postfix were found.")

    return kks_codes, kks_su_codes

# Example usage in process_pdf
def process_pdf(pdf_path: str, anchor_text="POS"):  # Removed search_string parameter
    """Process a single PDF to detect table structure."""
    try:
        with fitz.open(pdf_path) as doc:
            if len(doc) == 0:
                logging.debug(f"PDF: {os.path.basename(pdf_path)}")
                logging.debug("Status: Empty PDF")
                raise PDFProcessingError(f"PDF file is empty: {pdf_path}")

            for page_num in range(len(doc)):
                page = doc[page_num]
                page_dict = page.get_text("dict")

                logging.debug(f"Processing Page {page_num + 1} of PDF: {os.path.basename(pdf_path)}")

                # Detect table structure
                anchor_bbox, text_elements, avg_char_length, table_bounds = detect_table_structure(page_dict, anchor_text)
                if anchor_bbox:
                    logging.debug(f"Page {page_num + 1}: Anchor at {anchor_bbox}")
                    
                    # Extract and analyze text within table boundaries
                    table_boundaries = find_table_content(page_dict, avg_char_length, table_bounds)  # Pass header_bounds and table bottom y-coordinate
    except fitz.fitz.FileDataError as e:
        raise PDFProcessingError(f"Cannot open PDF file {pdf_path}: {str(e)}")
    except fitz.fitz.FileNotFoundError as e:
        raise PDFProcessingError(f"PDF file not found: {pdf_path}")
    except Exception as e:
        if isinstance(e, (AnchorTextNotFoundError, TableStructureError, PDFProcessingError)):
            raise
        else:
            raise PDFProcessingError(f"Unexpected error processing PDF {pdf_path}: {str(e)}")


# -----------------------
# Main Runner
# -----------------------

def run(input_dir: str):
    pdf_files = [str(p) for p in Path(input_dir).rglob("*.pdf")]

    for pdf in pdf_files:
        try:
            process_pdf(pdf)  # Removed search_string argument
        except Exception as e:
            logging.debug(f"PDF: {os.path.basename(pdf)}")
            logging.debug(f"Error: {e}")

def get_table_boundaries(pdf_path: str, anchor_text="POS"):
    """
    Detect table boundaries in a single PDF file.

    Args:
        pdf_path (str): Path to the PDF file.
        anchor_text (str): Anchor text to locate the table header.

    Returns:
        tuple: Table boundaries (x0, y0, x1, y1).
        
    Raises:
        PDFProcessingError: If PDF cannot be processed.
        AnchorTextNotFoundError: If anchor text is not found.
        TableStructureError: If table structure cannot be detected.
    """
    try:
        with fitz.open(pdf_path) as doc:
            if len(doc) == 0:
                logging.debug(f"PDF: {os.path.basename(pdf_path)} is empty.")
                raise PDFProcessingError(f"PDF file is empty: {pdf_path}")

            for page_num in range(len(doc)):
                page = doc[page_num]
                page_dict = page.get_text("dict")

                logging.debug(f"Processing Page {page_num + 1} of PDF: {os.path.basename(pdf_path)}")

                # Detect table structure
                anchor_bbox, text_elements, avg_char_length, table_bounds = detect_table_structure(page_dict, anchor_text)
                if anchor_bbox:
                    logging.debug(f"Page {page_num + 1}: Anchor at {anchor_bbox}")
                    table_boundaries = find_table_content(page_dict, avg_char_length, table_bounds)
                    return table_boundaries

        raise AnchorTextNotFoundError(f"No table boundaries detected in PDF: {os.path.basename(pdf_path)}")
        
    except fitz.fitz.FileDataError as e:
        raise PDFProcessingError(f"Cannot open PDF file {pdf_path}: {str(e)}")
    except fitz.fitz.FileNotFoundError as e:
        raise PDFProcessingError(f"PDF file not found: {pdf_path}")
    except Exception as e:
        if isinstance(e, (AnchorTextNotFoundError, TableStructureError, PDFProcessingError)):
            raise
        else:
            raise PDFProcessingError(f"Unexpected error processing PDF {pdf_path}: {str(e)}")




def get_table_boundaries_for_page(pdf_path: str, page_num: int) -> Tuple[float, float, float, float]:
    """
    Get table boundaries for a specific page using hardcoded "POS" anchor.
    
    Args:
        pdf_path (str): Path to the PDF file.
        page_num (int): Page number (1-based).
        
    Returns:
        Tuple[float, float, float, float]: Table boundaries (x0, y0, x1, y1).
        
    Raises:
        PDFProcessingError: If PDF cannot be opened or processed.
        PageNotFoundError: If page number is invalid.
        AnchorTextNotFoundError: If "POS" anchor text is not found.
        TableStructureError: If table structure cannot be determined.
    """
    logging.debug(f"Getting table boundaries for page {page_num} in {pdf_path}")
    
    try:
        with fitz.open(pdf_path) as doc:
            if len(doc) == 0:
                raise PDFProcessingError(f"PDF file is empty: {pdf_path}")
            
            if page_num < 1 or page_num > len(doc):
                raise PageNotFoundError(f"Invalid page number {page_num}. PDF has {len(doc)} pages")
                
            page = doc[page_num - 1]  # Convert to 0-based indexing
            page_dict = page.get_text("dict")

            # Get KKS codes and KKS/SU codes as two lists
            kks_codes_and_kks_su_codes = extract_kks_codes_from_page_dict(page_dict)

            # Use original detect_table_structure function with hardcoded "POS" anchor
            anchor_bbox, text_elements, avg_char_length, table_bounds = detect_table_structure(page_dict, "POS")
            
            # Get the actual table boundaries using original function
            table_boundaries = find_table_content(page_dict, avg_char_length, table_bounds)
            
            # Validate boundary values
            x0, y0, x1, y1 = table_boundaries
            if not all(isinstance(coord, (int, float)) for coord in table_boundaries):
                raise TableStructureError(f"Non-numeric table boundaries on page {page_num}: {table_boundaries}")
            
            if x0 >= x1 or y0 >= y1:
                raise TableStructureError(f"Invalid table boundary coordinates on page {page_num}: left={x0}, top={y0}, right={x1}, bottom={y1}")
            
            logging.debug(f"Successfully detected table boundaries on page {page_num}: {table_boundaries}")
            return table_boundaries, kks_codes_and_kks_su_codes
            
    except fitz.fitz.FileDataError as e:
        raise PDFProcessingError(f"Cannot open PDF file {pdf_path}: {str(e)}")
    except fitz.fitz.FileNotFoundError as e:
        raise PDFProcessingError(f"PDF file not found: {pdf_path}")
    except Exception as e:
        # Re-raise our custom exceptions
        if isinstance(e, (AnchorTextNotFoundError, TableStructureError, PageNotFoundError, PDFProcessingError)):
            raise
        else:
            raise PDFProcessingError(f"Unexpected error processing PDF {pdf_path}, page {page_num}: {str(e)}")

def get_total_pages(pdf_path: str) -> int:
    """
    Get the total number of pages in a PDF.
    
    Args:
        pdf_path (str): Path to the PDF file.
        
    Returns:
        int: Total number of pages.
        
    Raises:
        PDFProcessingError: If PDF cannot be opened.
    """
    try:
        with fitz.open(pdf_path) as doc:
            return len(doc)
    except Exception as e:
        raise PDFProcessingError(f"Cannot determine page count for {pdf_path}: {str(e)}")

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Detect table boundaries in a PDF file.")
    parser.add_argument("pdf_path", type=str, help="Path to the PDF file.")
    args = parser.parse_args()

    try:
        boundaries = get_table_boundaries(args.pdf_path)
        print(f"Table boundaries: {boundaries}")
    except TableBoundaryError as e:
        print(f"Error: {e}")
        exit(1)