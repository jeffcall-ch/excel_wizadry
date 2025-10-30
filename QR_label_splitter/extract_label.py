import fitz  # PyMuPDF
import pandas as pd
import os
import re
import hashlib
import configparser

# --- Configuration ---
config = configparser.ConfigParser()
config.read(r'C:\Users\szil\Repos\excel_wizadry\QR_label_splitter\config.ini')

input_pdf_path = config['Paths']['input_pdf']
excel_file_path = config['Paths']['excel_coordinates']
output_dir = config['Paths']['output_dir']

# Read search strings from the .ini file and parse them
search_strings_raw = config.get('Filters', 'search_strings', fallback='')
filter_strings = [s.strip() for s in search_strings_raw.split(',') if s.strip()]


# --- Helper Functions ---

def get_bom_item_no(text):
    """Extracts the value after 'BoM Item No:' using regex."""
    match = re.search(r"BoM Item No:\s*([^\n]+)", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return ""

def get_sort_key(text):
    """Extracts the number that appears after 'LxWxH'."""
    match = re.search(r"LxWxH\s+(\d+)", text, re.IGNORECASE)
    if match:
        return match.group(1)
    return ""

def natural_sort_key(s):
    """A key for natural sorting (e.g., 'item-10' comes after 'item-2')."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

def get_filter_strings(file_path):
    """Reads search strings from a CSV file."""
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        return []
    try:
        df = pd.read_csv(file_path, header=None)
        strings = df.values.flatten().tolist()
        return [str(s) for s in strings if pd.notna(s)]
    except pd.errors.EmptyDataError:
        return []

def get_rect_coordinates(file_path, sheet_name="Coordinates_measured"):
    """Reads rectangle coordinates from an Excel file."""
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df['group'] = df['points'] // 100
    rect_coords = {}
    for name, group in df.groupby('group'):
        if len(group) == 2:
            p1 = group.iloc[0]
            p2 = group.iloc[1]
            rect_coords[name] = fitz.Rect(p1['X'], p1['Y'], p2['X'], p2['Y'])
    return rect_coords

def add_content_to_page(output_docs, category, content_rect, source_doc, source_page_num):
    """Adds a rectangle of content to the correct output PDF, managing pages and layout.
    
    Uses clip_to_rect() to permanently remove content outside the label boundaries.
    This method physically removes text, images, and vector graphics that extend beyond
    the specified rectangle, while preserving searchable text within the boundaries.
    """
    a4_width, a4_height = fitz.paper_size("a4")
    margin = 20
    col_width = (a4_width - 3 * margin) / 2
    
    state = output_docs.setdefault(category, {
        'doc': fitz.open(),
        'page': None,
        'x': margin,
        'y': margin,
        'col': 0
    })
    
    if state['page'] is None:
        state['page'] = state['doc'].new_page(width=a4_width, height=a4_height)

    if state['y'] + content_rect.height > a4_height - margin:
        if state['col'] == 0:  # Was in left column, move to right
            state['col'] = 1
            state['x'] = margin + col_width + margin
            state['y'] = margin
        else:  # Was in right column, move to new page
            state['page'] = state['doc'].new_page(width=a4_width, height=a4_height)
            state['col'] = 0
            state['x'] = margin
            state['y'] = margin
    
    # --- CLEAN AND STAMP METHOD USING clip_to_rect() ---
    # Step 1: Create temporary PDF with the label content
    temp_doc = fitz.open()
    temp_page = temp_doc.new_page(width=content_rect.width, height=content_rect.height)
    temp_rect = fitz.Rect(0, 0, content_rect.width, content_rect.height)
    
    # Copy content from source, clipped to the label rectangle
    temp_page.show_pdf_page(temp_rect, source_doc, source_page_num, clip=content_rect)
    
    # Step 2: Permanently remove all content outside the label boundaries
    # This physically removes text, graphics, and images that extend beyond temp_rect
    temp_page.clip_to_rect(temp_rect)
    
    # Step 3: Stamp the cleaned label onto the output page
    target_rect = fitz.Rect(state['x'], state['y'], state['x'] + content_rect.width, state['y'] + content_rect.height)
    state['page'].show_pdf_page(target_rect, temp_doc, 0)
    
    # Clean up temporary document
    temp_doc.close()
    
    state['y'] += content_rect.height + 5  # Add gap

# --- Main Execution ---

rect_coords = get_rect_coordinates(excel_file_path)
source_doc = fitz.open(input_pdf_path)

# 1. Extract and hash all source labels for verification
source_label_hashes = set()
print("--- Hashing Source Labels ---")
for page in source_doc:
    for rect_id in sorted(rect_coords.keys()):
        label_rect = rect_coords[rect_id]
        # Get a pixel representation of the label for hashing
        pix = page.get_pixmap(clip=label_rect, alpha=False)
        source_label_hashes.add(hashlib.sha256(pix.samples).hexdigest())
print(f"Found and hashed {len(source_label_hashes)} unique labels from the source document.")


# 2. Extract, categorize, and sort all label data
all_labels = []
for page in source_doc:
    for rect_id in sorted(rect_coords.keys()):
        label_rect = rect_coords[rect_id]
        text = page.get_text("text", clip=label_rect)
        
        sort_key = get_sort_key(text)
        
        found_category = None
        if filter_strings:
            for s in filter_strings:
                if s.lower() in text.lower():
                    found_category = s
                    break
        
        category = found_category if found_category else "rest"
        
        all_labels.append({
            "sort_key": sort_key,
            "category": category,
            "rect": label_rect,
            "page_num": page.number
        })

# 3. Group labels by category
categorized_labels = {}
for label in all_labels:
    cat = label['category']
    if cat not in categorized_labels:
        categorized_labels[cat] = []
    categorized_labels[cat].append(label)

# 4. Sort labels within each category using natural sorting
for cat in categorized_labels:
    categorized_labels[cat].sort(key=lambda x: natural_sort_key(x['sort_key']))

# 5. Build the output PDFs and track processed hashes
output_docs = {}
processed_label_hashes = set()
if not filter_strings:
    # If no filters, all 'rest' go to 'output.pdf'
    if 'rest' in categorized_labels:
        for label in categorized_labels['rest']:
            add_content_to_page(output_docs, "output", label["rect"], source_doc, label["page_num"])
            pix = source_doc[label["page_num"]].get_pixmap(clip=label["rect"], alpha=False)
            processed_label_hashes.add(hashlib.sha256(pix.samples).hexdigest())
else:
    for cat, labels in categorized_labels.items():
        for label in labels:
            add_content_to_page(output_docs, cat, label["rect"], source_doc, label["page_num"])
            pix = source_doc[label["page_num"]].get_pixmap(clip=label["rect"], alpha=False)
            processed_label_hashes.add(hashlib.sha256(pix.samples).hexdigest())

# 6. Save all generated PDFs with compression and garbage collection
for category, state in output_docs.items():
    output_path = os.path.join(output_dir, f"{category}.pdf")
    # Use garbage=4 to remove duplicate objects, deflate=True for compression
    # clean=True to optimize content streams
    state['doc'].save(output_path, garbage=4, deflate=True, clean=True)
    state['doc'].close()
    print(f"Successfully created {output_path}")

source_doc.close()

# 7. Final Verification
print("\n--- Verification ---")
if source_label_hashes == processed_label_hashes:
    print(f"Success: All {len(source_label_hashes)} labels from the source document have been successfully extracted and processed.")
else:
    print(f"ERROR: Mismatch in processed labels.")
    print(f"  - Source document had {len(source_label_hashes)} unique labels.")
    print(f"  - Processed and saved {len(processed_label_hashes)} unique labels.")
    missing_count = len(source_label_hashes - processed_label_hashes)
    if missing_count > 0:
        print(f"  - {missing_count} labels were missed during processing.")

print("Verification complete.")
