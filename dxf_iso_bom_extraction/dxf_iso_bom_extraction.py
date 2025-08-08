import os
import csv
import ezdxf
import re
from collections import defaultdict

def is_piece_number(text):
    """Check if text is a piece number like <1>, <2>, etc."""
    return bool(re.match(r'^<\d+>$', str(text).strip()))

def is_number(text):
    """Check if text is a number (can include decimal)"""
    try:
        float(str(text).strip())
        return True
    except ValueError:
        return False

def is_text_remark(text):
    """Check if text is a remark (not a piece number or pure number)"""
    text_str = str(text).strip()
    if not text_str:
        return False
    return not is_piece_number(text_str) and not is_number(text_str)

def validate_and_correct_cut_length_row(row):
    """
    Validate and correct a CUT PIPE LENGTH row to ensure proper column alignment.
    Expected structure: [PIECE NO, CUT LENGTH, N.S.(MM), REMARKS, PIECE NO, CUT LENGTH, N.S.(MM), REMARKS]
    Each piece should have exactly 3 values: PIECE NO, CUT LENGTH, N.S.(MM)
    """
    if not row:
        return row
    
    # Create an 8-column template
    corrected = [''] * 8
    
    # Find piece numbers first to establish groupings
    pieces = []
    for i, cell in enumerate(row):
        cell_str = str(cell).strip()
        if is_piece_number(cell_str):
            pieces.append((i, cell_str))
    
    print(f"[DEBUG] Row validation - Found pieces: {pieces}")
    
    if len(pieces) == 0:
        # No pieces found, return original row
        return row[:8] + [''] * (8 - len(row)) if len(row) < 8 else row[:8]
    
    # Group data by piece based on original positions
    piece_groups = []
    
    for piece_idx, (piece_pos, piece_name) in enumerate(pieces):
        group = {'piece': piece_name, 'numbers': [], 'remarks': []}
        
        # Determine the range for this piece
        if piece_idx + 1 < len(pieces):
            next_piece_pos = pieces[piece_idx + 1][0]
            end_pos = next_piece_pos
        else:
            end_pos = len(row)
        
        # Collect numbers and remarks between this piece and the next
        for i in range(piece_pos + 1, end_pos):
            if i < len(row):
                cell_str = str(row[i]).strip()
                if not cell_str:
                    continue
                    
                if is_number(cell_str):
                    group['numbers'].append(cell_str)
                elif is_text_remark(cell_str):
                    group['remarks'].append(cell_str)
        
        piece_groups.append(group)
    
    print(f"[DEBUG] Piece groups: {piece_groups}")
    
    # Fill the corrected row
    for group_idx, group in enumerate(piece_groups[:2]):  # Max 2 pieces per row
        base_col = group_idx * 4  # 0 for first piece, 4 for second piece
        
        # Place piece number
        corrected[base_col] = group['piece']
        
        # Place numbers (should be CUT LENGTH and N.S.)
        # Each piece should have exactly 2 numbers: CUT LENGTH and N.S.(MM)
        for num_idx, number in enumerate(group['numbers'][:2]):
            corrected[base_col + 1 + num_idx] = number
        
        # Place remarks
        if group['remarks']:
            corrected[base_col + 3] = group['remarks'][0]  # First remark only
    
    print(f"[DEBUG] Corrected row: {corrected}")
    return corrected

def extract_text_entities(doc):
    # Returns list of (text, x, y) for all text entities
    entities = []
    for e in doc.modelspace():
        if e.dxftype() == 'MTEXT':
            text = e.plain_text()
            x, y = e.dxf.insert.x, e.dxf.insert.y
            entities.append((text.strip(), x, y))
        elif e.dxftype() == 'TEXT':
            text = e.dxf.text
            x, y = e.dxf.insert.x, e.dxf.insert.y
            entities.append((text.strip(), x, y))
    return entities

def find_drawing_no(text_entities):
    # Find KKS code with pattern 1AAA11BR111 (1=digit, A=capital letter, BR=fixed)
    # Located in bottom right corner, below and to the right of ERECTION MATERIALS
    import re
    
    # First find ERECTION MATERIALS position to establish search area
    erection_x, erection_y = None, None
    for text, x, y in text_entities:
        if 'ERECTION MATERIALS' in text.upper():
            erection_x, erection_y = x, y
            print(f"[DEBUG] Found ERECTION MATERIALS at X={x}, Y={y}")
            break
    
    # KKS pattern: digit + 3 capital letters + 2 digits + "BR" + 3 digits
    kks_pattern = r'\b\d[A-Z]{3}\d{2}BR\d{3}\b'
    
    candidates = []
    for text, x, y in text_entities:
        match = re.search(kks_pattern, text)
        if match:
            kks_code = match.group()
            # If we found ERECTION MATERIALS, filter by position (below and to the right)
            if erection_x is not None and erection_y is not None:
                if x >= erection_x and y <= erection_y:  # Right and below
                    candidates.append((kks_code, x, y, abs(y)))  # Use abs(y) for sorting (lower Y = further down)
                    print(f"[DEBUG] Found KKS candidate '{kks_code}' at X={x}, Y={y}")
            else:
                # If no ERECTION MATERIALS found, consider all KKS codes
                candidates.append((kks_code, x, y, abs(y)))
                print(f"[DEBUG] Found KKS candidate '{kks_code}' at X={x}, Y={y} (no ERECTION MATERIALS reference)")
    
    if candidates:
        # Sort by Y coordinate (lowest Y = bottom of drawing) and pick the first one
        candidates.sort(key=lambda item: item[3])  # Sort by abs(y)
        selected_kks = candidates[0][0]
        print(f"[DEBUG] Selected KKS code: '{selected_kks}'")
        return selected_kks
    
    # Fallback: try to find Drawing-No. field if no KKS found
    for i, (text, x, y) in enumerate(text_entities):
        if 'Drawing-No.' in text:
            # Look for next text entity to the right or below
            for t, tx, ty in text_entities[i+1:i+5]:
                if tx > x or ty < y:
                    print(f"[DEBUG] Fallback: Found Drawing-No. '{t}'")
                    return t
    
    print(f"[DEBUG] No KKS code or Drawing-No. found")
    return ''

def process_dxf_file(filepath):
    try:
        print(f"[DEBUG] Opening DXF file: {filepath}")
        doc = ezdxf.readfile(filepath)
        text_entities = extract_text_entities(doc)
        drawing_no = find_drawing_no(text_entities)
        mat_header, mat_rows = extract_table(text_entities, 'ERECTION MATERIALS')
        cut_header, cut_rows = extract_table(text_entities, 'CUT PIPE LENGTH')
        # Add Drawing-No. to each row
        if mat_rows:
            mat_header_out = mat_header + ['Drawing-No.']
            mat_rows_out = [r + [drawing_no] for r in mat_rows]
        else:
            mat_header_out, mat_rows_out = [], []
        if cut_rows:
            cut_header_out = cut_header + ['Drawing-No.']
            cut_rows_out = [r + [drawing_no] for r in cut_rows]
        else:
            cut_header_out, cut_rows_out = [], []
        print(f"[DEBUG] Extracted {len(mat_rows_out)} material rows and {len(cut_rows_out)} cut length rows from {filepath}")
        return {
            'drawing_no': drawing_no,
            'mat_header': mat_header_out,
            'mat_rows': mat_rows_out,
            'cut_header': cut_header_out,
            'cut_rows': cut_rows_out,
            'error': None
        }
    except Exception as e:
        print(f"[ERROR] Failed to process {filepath}: {e}")
        return {
            'drawing_no': '',
            'mat_header': [],
            'mat_rows': [],
            'cut_header': [],
            'cut_rows': [],
            'error': str(e)
        }

def extract_table(text_entities, table_title, max_cols=20, max_rows=100):
    # Find table by title, then extract rows/columns by y/x positions
    # Returns: [header], [rows]
    # Find title
    title_entity = None
    start_x = None
    title_y = None
    for text, x, y in text_entities:
        if table_title.lower() in text.lower():
            start_x = x  # Use the start of the DXF entity
            title_y = y
            print(f"[DEBUG] Table title '{table_title}' found at X={x}, Y={y}, text='{text}'")
            title_entity = (text, x, y)
            break
    
    if not title_entity or start_x is None:
        print(f"[DEBUG] Table title '{table_title}' not found.")
        return [], []
    
    title_text, title_x, title_y = title_entity
    # For CUT PIPE LENGTH, use a more relaxed X filter since data might be to the left of title
    if table_title.lower() == 'cut pipe length':
        # Allow data to the left of the title for cut pipe length table
        filtered_entities = [(text, x, y) for text, x, y in text_entities if y < title_y and x >= (title_x - 50)]
    else:
        # Only keep entities with x >= title_x and y < title_y
        filtered_entities = [(text, x, y) for text, x, y in text_entities if y < title_y and x >= title_x]
    
    # Now process table from filtered_entities only
    rows_dict = defaultdict(list)
    for text, x, y in filtered_entities:
        rows_dict[round(y, 1)].append((x, text))
    sorted_rows = sorted(rows_dict.items(), key=lambda item: -item[0])
    
    # For each row, sort by x ascending (left to right)
    table_rows = []
    for idx, (y, cells) in enumerate(sorted_rows):
        row = [t for x, t in sorted(cells, key=lambda c: c[0])]
        xs = [x for x, t in sorted(cells, key=lambda c: c[0])]
        if table_title.lower() == 'cut pipe length' and idx == 2:
            print(f"[DEBUG] Extracted row {idx+1} at y={y}, x={xs}: {row} <-- 3RD ROW BELOW 'CUT PIPE LENGTH'")
        else:
            print(f"[DEBUG] Extracted row {idx+1} at y={y}, x={xs}: {row}")
        table_rows.append(row)
    
    if table_title.lower() == 'cut pipe length':
        print(f"[DEBUG] Total rows extracted for 'CUT PIPE LENGTH': {len(table_rows)}")
    # (Removed duplicate/old code using table_entities)
    # Heuristic: first two rows are headers, merge them
    if len(table_rows) >= 2:
        header = []
        for i in range(max(len(table_rows[0]), len(table_rows[1]))):
            h1 = table_rows[0][i] if i < len(table_rows[0]) else ''
            h2 = table_rows[1][i] if i < len(table_rows[1]) else ''
            # Special handling for CUT PIPE LENGTH table headers
            if table_title.lower() == 'cut pipe length':
                # For CUT PIPE LENGTH, handle the specific pattern more carefully
                if h1 and h2:
                    if h1 == 'N.S.' and h2 == '(MM)':
                        merged = 'N.S. (MM)'
                    elif h1 == 'PIECE' and h2 == 'NO':
                        merged = 'PIECE NO'
                    elif h1 == 'CUT' and h2 == 'LENGTH':
                        merged = 'CUT LENGTH'
                    elif h1 == 'REMARKS' and h2 == 'NO':
                        merged = 'REMARKS'  # Don't add 'NO' to REMARKS
                    elif h1 == 'REMARKS' and not h2:
                        merged = 'REMARKS'
                    # Handle the second set of columns (right side)
                    elif h1 == 'PIECE' and h2 == 'LENGTH':  # Should be 'PIECE NO'
                        merged = 'PIECE NO'
                    elif h1 == 'CUT' and h2 == '(MM)':  # Should be 'CUT LENGTH'
                        merged = 'CUT LENGTH'
                    else:
                        merged = f"{h1} {h2}".strip()
                elif h1 and not h2:
                    # Handle single header values for the right side columns
                    if h1 == 'PIECE':
                        merged = 'PIECE NO'
                    elif h1 == 'CUT':
                        merged = 'CUT LENGTH'
                    elif h1 == 'N.S.':
                        merged = 'N.S. (MM)'
                    else:
                        merged = h1
                elif h2 and not h1:
                    merged = h2
                else:
                    merged = ''
            else:
                merged = f"{h1} {h2}".strip() if h1 or h2 else ''
            header.append(merged)
        data_rows = table_rows[2:]
    else:
        header = table_rows[0] if table_rows else []
        data_rows = table_rows[1:] if len(table_rows) > 1 else []
    # For ERECTION MATERIALS, stop at 'TOTAL WEIGHT' row
    if 'ERECTION MATERIALS' in table_title.upper():
        end_idx = None
        for i, row in enumerate(data_rows):
            if any('TOTAL WEIGHT' in str(cell).upper() for cell in row):
                end_idx = i + 1
                break
        if end_idx:
            data_rows = data_rows[:end_idx]
        
        # Process ERECTION MATERIALS to move categories from column A to column F
        processed_rows = []
        current_category = ""
        
        for row in data_rows:
            # Check if this row is a category header (has content in first column but empty in other key columns)
            if row and row[0] and (len(row) < 2 or not row[1] or len(row) < 4 or not row[3]):
                # This is likely a category header like "PIPE", "FITTINGS", etc.
                if row[0] not in ["TOTAL ERECTION WEIGHT", "TOTAL WEIGHT"]:
                    current_category = row[0]
                    print(f"[DEBUG] Found category: '{current_category}'")
                    continue  # Skip category header rows, don't add to processed_rows
                else:
                    # For total rows, move the total type to column F and weight value to column E
                    new_row = row[:]
                    total_type = new_row[0]  # Save the total type
                    weight_value = new_row[1] if len(new_row) > 1 else ""  # Save the weight value from column B
                    new_row[0] = ""  # Clear column A
                    if len(new_row) > 1:
                        new_row[1] = ""  # Clear column B
                    while len(new_row) < 6:
                        new_row.append("")
                    new_row[4] = weight_value  # Put weight value in column E (WEIGHT column)
                    new_row.insert(5, total_type)  # Insert total type at position 5 (column F)
                    processed_rows.append(new_row)
            else:
                # This is a regular data row, add the current category to column F
                new_row = row[:]
                while len(new_row) < 6:
                    new_row.append("")
                new_row.insert(5, current_category)  # Insert category at position 5 (column F)
                processed_rows.append(new_row)
        
        data_rows = processed_rows
        
        # Update header to include the new CATEGORY column
        if header:
            header.insert(5, "CATEGORY")  # Insert at position 5 (column F)
    # For CUT PIPE LENGTH, filter rows with '<' 
    if table_title.lower() == 'cut pipe length':
        kept_rows = []
        for row in data_rows:
            if '<' in ''.join(row):
                kept_rows.append(row)
        print(f"[DEBUG] Kept rows for 'CUT PIPE LENGTH':")
        for r in kept_rows:
            print(f"[DEBUG] {r}")
        data_rows = kept_rows
        
        # Apply column validation and correction for CUT PIPE LENGTH
        corrected_rows = []
        for row in data_rows:
            corrected_row = validate_and_correct_cut_length_row(row)
            corrected_rows.append(corrected_row)
        data_rows = corrected_rows
    # For CUT PIPE LENGTH, stop at first empty row
    if 'CUT PIPE LENGTH' in table_title.upper():
        new_data_rows = []
        for row in data_rows:
            if all(cell.strip() == '' for cell in row):
                break
            new_data_rows.append(row)
        data_rows = new_data_rows
    # Pad all rows to header length, keep empty cells
    padded_rows = []
    for row in data_rows:
        padded = row + [''] * (len(header) - len(row)) if len(row) < len(header) else row[:len(header)]
        padded_rows.append(padded)
    return header, padded_rows

def main(directory):
    material_rows = []
    cut_rows = []
    mat_header = None
    cut_header = None
    summary = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.dxf'):
                path = os.path.join(root, file)
                result = process_dxf_file(path)
                summary_row = {
                    'file_path': path,
                    'filename': file,
                    'drawing_no': result['drawing_no'],
                    'mat_rows': len(result['mat_rows']),
                    'cut_rows': len(result['cut_rows']),
                    'mat_missing': not bool(result['mat_rows']),
                    'cut_missing': not bool(result['cut_rows']),
                    'error': result['error']
                }
                summary.append(summary_row)
                if result['mat_rows']:
                    if not mat_header:
                        mat_header = result['mat_header']
                    material_rows.extend(result['mat_rows'])
                if result['cut_rows']:
                    if not cut_header:
                        cut_header = result['cut_header']
                    cut_rows.extend(result['cut_rows'])
    # Write CSVs
    # Write all_materials.csv
    out_path = os.path.join(directory, 'all_materials.csv')
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if material_rows:
            writer.writerow(mat_header)
            writer.writerows(material_rows)
        else:
            writer.writerow(['No Data'])
    # Write all_cut_lengths.csv
    out_path = os.path.join(directory, 'all_cut_lengths.csv')
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if cut_rows:
            writer.writerow(cut_header)
            writer.writerows(cut_rows)
        else:
            writer.writerow(['No Data'])
    # Write summary.csv
    out_path = os.path.join(directory, 'summary.csv')
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['file_path', 'filename', 'drawing_no', 'mat_rows', 'cut_rows', 'mat_missing', 'cut_missing', 'error'])
        writer.writeheader()
        writer.writerows(summary)

if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print('Usage: python dxf_iso_bom_extraction.py <directory>')
    else:
        main(sys.argv[1])
