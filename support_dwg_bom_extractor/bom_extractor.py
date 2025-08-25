#!/usr/bin/env python3
"""
MVP PDF BOM Extractor - Single Process, Simple CSV Output
"""

import os
import csv
import fitz  # PyMuPDF
from pathlib import Path
from typing import List, Tuple

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
    header_bottom_y = anchor_y1  # Initialize with the bottom of the anchor

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
                    header_bottom_y = max(header_bottom_y, text_y1)  # Update header_bottom_y
                    print(f"  - Text: '{text}'")
                    print(f"    BBox: {bbox}")

    print(f"\n[DEBUG] Total Header Text Elements Detected: {len(text_elements)}")
    print(f"[DEBUG] Header Bottom Y: {header_bottom_y}\n")
    return anchor_bbox, text_elements

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

def print_text_below_anchor(page_dict: dict, anchor_bbox: Tuple[float, float, float, float]):
    """
    Print all text where text_x0 >= anchor_x0 and text_y0 > anchor_y1.
    """
    anchor_x0 = anchor_bbox[0]
    anchor_y1 = anchor_bbox[3]

    print("\n[DEBUG] Printing Text Below Anchor")
    print(f"  - Anchor X0: {anchor_x0}")
    print(f"  - Anchor Y1: {anchor_y1}\n")

    for block in page_dict.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span["text"]
                bbox = span["bbox"]
                text_x0, text_y0 = bbox[0], bbox[1]

                if text_x0 >= anchor_x0 and text_y0 > anchor_y1:
                    print(f"  - Text: '{text}'")
                    print(f"    BBox: {bbox}\n")

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
            anchor_bbox, text_elements = detect_table_structure(page_dict, anchor_text)
            if anchor_bbox:
                print(f"[DEBUG] Page {page_num + 1}: Anchor at {anchor_bbox}\n")
                print_text_below_anchor(page_dict, anchor_bbox)  # Print text below anchor

                # Call find_table_bottom to locate the 'TOTAL' below 'WEIGHT'
                table_bottom = find_table_bottom(page_dict, "TOTAL")
                if table_bottom:
                    print(f"[DEBUG] Table Bottom Found: {table_bottom}\n")
                else:
                    print("[DEBUG] Table Bottom Not Found\n")

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
