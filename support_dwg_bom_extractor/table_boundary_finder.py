#!/usr/bin/env python3
"""
MVP PDF BOM Extractor - Single Process, Simple CSV Output
"""

import os
import csv
import fitz  # PyMuPDF
from pathlib import Path
from typing import List, Tuple, NamedTuple

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
        print(f"\n[DEBUG] Anchor '{anchor_text}' not found on the page.\n")
        return None, []

    anchor_bbox = anchor["bbox"]
    anchor_x0 = anchor_bbox[0]  # x0 of the anchor
    anchor_x1 = anchor_bbox[2]  # x1 of the anchor
    anchor_y0 = anchor_bbox[1]  # y0 of the anchor
    anchor_y1 = anchor_bbox[3]  # y1 of the anchor

    # Calculate anchor height and vertical midpoint
    anchor_height = anchor_y1 - anchor_y0
    anchor_vertical_mid = anchor_y0 + anchor_height / 2
    tolerance = 0.2 * anchor_height  # 20% of anchor height

    print("\n[DEBUG] Anchor Detection")
    print(f"  - Anchor Text: '{anchor_text}'")
    print(f"  - Anchor Position: {anchor_bbox}")
    print(f"  - Anchor Height: {anchor_height}")
    print(f"  - Anchor Vertical Midpoint: {anchor_vertical_mid}")
    print(f"  - Tolerance: {tolerance}\n")

    # Define the data structure for text elements
    text_elements = []
    header_top_y = anchor_y0  # Initialize with the bottom of the anchor
    header_left_x = anchor_x0  # Initialize with the left of the anchor
    header_right_x = anchor_x1  # Initialize with the right of the anchor

    # Find qualifying text elements
    print("[DEBUG] Header Text Detection")
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
                    print(f"  - Text: '{text}'")
                    print(f"    BBox: {bbox}")

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
                print(f"[DEBUG] Header: '{header}'")
                print(f"  - Stripped Text: '{stripped_text}'")
                print(f"  - BBox: {bbox}")
                print(f"  - Average Character Length (X): {avg_char_length}\n")
    
    # Call find_table_bottom to locate the 'TOTAL' below 'WEIGHT'
    table_bottom_y = find_table_bottom(page_dict, "TOTAL")
    if table_bottom_y:
        print(f"[DEBUG] Table Bottom Found: {table_bottom_y}\n")
    else:
        print("[DEBUG] Table Bottom Not Found\n")
    
    # Add some padding to  bounding boxes to make sure we have all text in the box and possible the table frame too
    header_left_x -= avg_char_length * 2
    header_right_x += avg_char_length * 2
    header_top_y -= anchor_height * 2
    table_bottom_y = table_bottom_y["bbox"][3] + avg_char_length * 2
    

    print(f"\n[DEBUG] Total Header Text Elements Detected: {len(text_elements)}")
    print(f"[DEBUG] Header TOP Y: {header_top_y}")
    print(f"[DEBUG] Header LEFT X: {header_left_x}")
    print(f"[DEBUG] Header RIGHT X: {header_right_x}\n")

    return anchor_bbox, text_elements, avg_char_length, (header_top_y, header_left_x, header_right_x, table_bottom_y)

def find_table_bottom(page_dict: dict, target_text: str):
    """
    Find the target text (e.g., 'TOTAL') below the hardcoded 'WEIGHT' header.
    Return the target text and its bounding box in the usual data format.
    """
    print("\n[DEBUG] Finding Table Bottom")
    print("  - Hardcoded Header Text: 'WEIGHT'")
    print(f"  - Target Text: '{target_text}'\n")

    # Step 1: Locate the hardcoded 'WEIGHT' header
    header_bbox = None
    for block in page_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"]
                bbox = span["bbox"]

                if text.strip() == "WEIGHT":  # Full match for 'WEIGHT'
                    header_bbox = bbox
                    print(f"  - Hardcoded Header Found: '{text}'")
                    print(f"    BBox: {bbox}\n")
                    break
            if header_bbox:
                break
        if header_bbox:
            break

    if not header_bbox:
        print("[DEBUG] Hardcoded Header 'WEIGHT' not found.\n")
        return None

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
                    print(f"  - Target Found: '{text}'")
                    print(f"    BBox: {bbox}\n")
                    return {"text": text, "bbox": bbox}

    # Debug statement if TOTAL is not found
    print(f"[DEBUG] Target '{target_text}' not found below hardcoded header 'WEIGHT'.\n")
    return None

def find_table_content(page_dict: dict, avg_char_length: float, table_bounds: Tuple[float, float, float, float]):
    """
    Extract all text within the table boundaries and print their coordinates.

    Args:
        page_dict (dict): The page dictionary containing text blocks.
        header_bounds (Tuple[float, float, float]): Tuple containing (header_top, header_left, header_right).
        table_bottom (float): The Y-coordinate of the table bottom.
    """
    header_top, header_left, header_right, table_bottom = table_bounds

    # Debug print header bounds
    print("\n[DEBUG] Header Bounds:")
    print(f"  - Header Top: {header_top}")
    print(f"  - Header Left: {header_left}")
    print(f"  - Header Right: {header_right}")
    print(f"  - Table Bottom: {table_bottom}")

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

        print("\n[DEBUG] Table Content:")
        for element in table_text_elements:
            print(f"  - Text: '{element['text']}'")
            print(f"    BBox: {element['bbox']}")

        print("\n[DEBUG] Table Boundary Coordinates:")
        print(f"  - Minimum X0: {min_x0} (Text: '{min_x0_element['text']}', BBox: {min_x0_element['bbox']})")
        print(f"  - Maximum X1: {max_x1} (Text: '{max_x1_element['text']}', BBox: {max_x1_element['bbox']})")
        print(f"  - Minimum Y0: {min_y0} (Text: '{min_y0_element['text']}', BBox: {min_y0_element['bbox']})")
        print(f"  - Maximum Y1: {max_y1} (Text: '{max_y1_element['text']}', BBox: {max_y1_element['bbox']})\n")
    else:
        print("\n[DEBUG] No text found within the table boundaries.\n")


# Example usage in process_pdf
def process_pdf(pdf_path: str, anchor_text="POS"):  # Removed search_string parameter
    """Process a single PDF to detect table structure."""
    with fitz.open(pdf_path) as doc:
        if len(doc) == 0:
            print(f"\n[DEBUG] PDF: {os.path.basename(pdf_path)}")
            print("[DEBUG] Status: Empty PDF\n")
            return []

        for page_num in range(len(doc)):
            page = doc[page_num]
            page_dict = page.get_text("dict")

            print(f"\n[DEBUG] Processing Page {page_num + 1} of PDF: {os.path.basename(pdf_path)}")

            # Detect table structure
            anchor_bbox, text_elements, avg_char_length, table_bounds = detect_table_structure(page_dict, anchor_text)
            if anchor_bbox:
                print(f"[DEBUG] Page {page_num + 1}: Anchor at {anchor_bbox}\n")
                
                # Extract and analyze text within table boundaries
                find_table_content(page_dict, avg_char_length, table_bounds)  # Pass header_bounds and table bottom y-coordinate
            


# -----------------------
# Main Runner
# -----------------------

def run(input_dir: str):
    pdf_files = [str(p) for p in Path(input_dir).rglob("*.pdf")]

    for pdf in pdf_files:
        try:
            process_pdf(pdf)  # Removed search_string argument
        except Exception as e:
            print(f"\n[DEBUG] PDF: {os.path.basename(pdf)}")
            print(f"[DEBUG] Error: {e}\n")

if __name__ == "__main__":
    import sys
    # if len(sys.argv) < 2:
    #     print("Usage: python bom_extractor_mvp.py <input_directory>")
    #     sys.exit(1)
    # run(sys.argv[1])
    run(r"C:\Users\szil\Repos\excel_wizadry\support_dwg_bom_extractor")
