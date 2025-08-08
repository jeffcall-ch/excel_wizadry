#!/usr/bin/env python3

import os
import csv
from dxf_iso_bom_extraction import process_dxf_file

def test_single_file():
    filename = "TB020-INOV-2QFB94BR140_1.0_Pipe-Isometric-Drawing-ServiceAir-Lot_General_Piping_Engineering.dxf"
    filepath = os.path.join(".", filename)
    
    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        return
    
    print(f"Testing single file: {filename}")
    result = process_dxf_file(filepath)
    
    print(f"\n=== RESULTS ===")
    print(f"Drawing No: {result['drawing_no']}")
    print(f"Material rows: {len(result['mat_rows'])}")
    print(f"Cut length rows: {len(result['cut_rows'])}")
    
    if result['cut_header']:
        print(f"\n=== CUT PIPE LENGTH HEADER ===")
        print(result['cut_header'])
        
    if result['cut_rows']:
        print(f"\n=== CUT PIPE LENGTH ROWS ===")
        for i, row in enumerate(result['cut_rows']):
            print(f"Row {i+1}: {row}")
    
    # Generate CSV files for this single file
    print(f"\n=== GENERATING CSV FILES ===")
    
    # Write materials CSV
    mat_csv_path = f"single_file_materials_{result['drawing_no']}.csv"
    with open(mat_csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if result['mat_rows']:
            writer.writerow(result['mat_header'])
            writer.writerows(result['mat_rows'])
        else:
            writer.writerow(['No Data'])
    print(f"Materials CSV saved: {mat_csv_path}")
    
    # Write cut lengths CSV
    cut_csv_path = f"single_file_cut_lengths_{result['drawing_no']}.csv"
    with open(cut_csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if result['cut_rows']:
            writer.writerow(result['cut_header'])
            writer.writerows(result['cut_rows'])
        else:
            writer.writerow(['No Data'])
    print(f"Cut lengths CSV saved: {cut_csv_path}")
    
    # Write summary CSV
    summary_csv_path = f"single_file_summary_{result['drawing_no']}.csv"
    with open(summary_csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['filename', 'drawing_no', 'mat_rows', 'cut_rows', 'mat_missing', 'cut_missing', 'error'])
        writer.writeheader()
        writer.writerow({
            'filename': filename,
            'drawing_no': result['drawing_no'],
            'mat_rows': len(result['mat_rows']),
            'cut_rows': len(result['cut_rows']),
            'mat_missing': not bool(result['mat_rows']),
            'cut_missing': not bool(result['cut_rows']),
            'error': result['error']
        })
    print(f"Summary CSV saved: {summary_csv_path}")

if __name__ == '__main__':
    test_single_file()
