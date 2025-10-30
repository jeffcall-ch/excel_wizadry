import fitz
import pandas as pd
import os

# --- Configuration ---
input_pdf_path = r"C:\Users\szil\Repos\excel_wizadry\QR_label_splitter\2025-10-27_KVI Parts List Label - BoM Items.pdf"
excel_file_path = r"C:\Users\szil\Repos\excel_wizadry\QR_label_splitter\coordinates.xlsx"
sheet_name = "Coordinates_measured"
NUMBER_OF_LABELS_TO_DEBUG = 5

# --- Functions ---

def get_label_rects(file_path, sheet_name, limit):
    """Reads coordinates from Excel and returns a list of rectangles."""
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df['group'] = df['points'] // 100
    
    rects = []
    group_ids = sorted(df['group'].unique())[:limit]
    
    for group_id in group_ids:
        group = df[df['group'] == group_id]
        if len(group) == 2:
            p1 = group.iloc[0]
            p2 = group.iloc[1]
            rects.append(fitz.Rect(p1['X'], p1['Y'], p2['X'], p2['Y']))
    return rects

# --- Main Execution ---

# 1. Get the rectangles for the first few labels
label_rects = get_label_rects(excel_file_path, sheet_name, NUMBER_OF_LABELS_TO_DEBUG)

if not label_rects:
    print("Could not define any label rectangles from the coordinates file.")
else:
    # 2. Open the PDF and get the first page
    doc = fitz.open(input_pdf_path)
    page = doc[0]

    # 3. Iterate through each label and print its text
    for i, rect in enumerate(label_rects):
        print(f"--- Text from Label {i+1} ---")
        
        words = page.get_text("words", clip=rect)
        
        if not words:
            print("No text found in the specified area.")
        else:
            lines = {}
            for w in words:
                line_num = w[6]
                if line_num not in lines:
                    lines[line_num] = []
                lines[line_num].append(w)
            
            for line_num in lines:
                lines[line_num].sort(key=lambda w: w[0])
                
            for line_num in sorted(lines.keys()):
                line_text = " ".join([w[4] for w in lines[line_num]])
                print(line_text)
        
        print("\n") # Add a newline for readability

    print(f"--- End of text for {len(label_rects)} labels ---")
    doc.close()
