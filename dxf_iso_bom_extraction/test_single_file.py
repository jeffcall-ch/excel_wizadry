#!/usr/bin/env python3

import os
import csv
import time
from dxf_iso_bom_extraction import extract_text_entities, find_drawing_no, find_pipe_class, extract_table, convert_cut_length_to_single_row_format
import ezdxf

def test_single_file_with_timing():
    filename = "TB020-INOV-2QFB94BR140_1.0_Pipe-Isometric-Drawing-ServiceAir-Lot_General_Piping_Engineering.dxf"
    filepath = os.path.join(".", filename)
    
    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        return
    
    print(f"=== TIMING BREAKDOWN FOR SINGLE FILE PROCESSING ===")
    print(f"File: {filename}")
    print("=" * 70)
    
    # Start total timing
    total_start_time = time.time()
    
    try:
        # 1. File Opening
        step_start = time.time()
        doc = ezdxf.readfile(filepath)
        file_open_time = (time.time() - step_start) * 1000  # Convert to milliseconds
        print(f"1. File Opening:                {file_open_time:8.2f} ms")
        
        # 2. Text Entity Extraction
        step_start = time.time()
        text_entities = extract_text_entities(doc)
        text_extraction_time = (time.time() - step_start) * 1000
        print(f"2. Text Entity Extraction:      {text_extraction_time:8.2f} ms   ({len(text_entities)} entities)")
        
        # 3. Drawing Number Extraction
        step_start = time.time()
        drawing_no = find_drawing_no(text_entities)
        drawing_no_time = (time.time() - step_start) * 1000
        print(f"3. Drawing Number Extraction:   {drawing_no_time:8.2f} ms   ('{drawing_no}')")
        
        # 4. Pipe Class Extraction
        step_start = time.time()
        pipe_class = find_pipe_class(text_entities)
        pipe_class_time = (time.time() - step_start) * 1000
        print(f"4. Pipe Class Extraction:       {pipe_class_time:8.2f} ms   ('{pipe_class}')")
        
        # 5. ERECTION MATERIALS Table Extraction
        step_start = time.time()
        mat_header, mat_rows = extract_table(text_entities, 'ERECTION MATERIALS')
        mat_extraction_time = (time.time() - step_start) * 1000
        print(f"5. Materials Table Extraction:  {mat_extraction_time:8.2f} ms   ({len(mat_rows)} rows)")
        
        # 6. CUT PIPE LENGTH Table Extraction
        step_start = time.time()
        cut_header, cut_rows = extract_table(text_entities, 'CUT PIPE LENGTH')
        cut_extraction_time = (time.time() - step_start) * 1000
        print(f"6. Cut Lengths Table Extraction:{cut_extraction_time:8.2f} ms   ({len(cut_rows)} rows)")
        
        # 7. Data Processing (adding columns, format conversion)
        step_start = time.time()
        
        # Process materials table
        if mat_rows:
            mat_header_out = mat_header + ['Drawing-No.', 'Pipe Class']
            mat_rows_out = [r + [drawing_no, pipe_class] for r in mat_rows]
        else:
            mat_header_out, mat_rows_out = [], []
        
        # Process cut lengths table
        if cut_rows:
            cut_header_converted, cut_rows_converted = convert_cut_length_to_single_row_format(cut_header, cut_rows, drawing_no, pipe_class)
            cut_header_out = cut_header_converted
            cut_rows_out = cut_rows_converted
        else:
            cut_header_out, cut_rows_out = [], []
            
        data_processing_time = (time.time() - step_start) * 1000
        print(f"7. Data Processing:             {data_processing_time:8.2f} ms   (formatting & column addition)")
        
        print("-" * 70)
        
        # 8. CSV File Writing
        csv_start = time.time()
        
        # Write materials CSV
        mat_csv_start = time.time()
        mat_csv_path = f"single_file_materials_{drawing_no}.csv"
        with open(mat_csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if mat_rows_out:
                writer.writerow(mat_header_out)
                writer.writerows(mat_rows_out)
            else:
                writer.writerow(['No Data'])
        mat_csv_time = (time.time() - mat_csv_start) * 1000
        
        # Write cut lengths CSV
        cut_csv_start = time.time()
        cut_csv_path = f"single_file_cut_lengths_{drawing_no}.csv"
        with open(cut_csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if cut_rows_out:
                writer.writerow(cut_header_out)
                writer.writerows(cut_rows_out)
            else:
                writer.writerow(['No Data'])
        cut_csv_time = (time.time() - cut_csv_start) * 1000
        
        # Write summary CSV
        summary_csv_start = time.time()
        summary_csv_path = f"single_file_summary_{drawing_no}.csv"
        with open(summary_csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['filename', 'drawing_no', 'pipe_class', 'mat_rows', 'cut_rows', 'mat_missing', 'cut_missing', 'error'])
            writer.writeheader()
            writer.writerow({
                'filename': filename,
                'drawing_no': drawing_no,
                'pipe_class': pipe_class,
                'mat_rows': len(mat_rows_out),
                'cut_rows': len(cut_rows_out),
                'mat_missing': not bool(mat_rows_out),
                'cut_missing': not bool(cut_rows_out),
                'error': ''
            })
        summary_csv_time = (time.time() - summary_csv_start) * 1000
        
        total_csv_time = (time.time() - csv_start) * 1000
        
        print(f"8a. Materials CSV Write:        {mat_csv_time:8.2f} ms   ({len(mat_rows_out)} rows)")
        print(f"8b. Cut Lengths CSV Write:      {cut_csv_time:8.2f} ms   ({len(cut_rows_out)} rows)")
        print(f"8c. Summary CSV Write:          {summary_csv_time:8.2f} ms")
        print(f"8.  Total CSV Writing:          {total_csv_time:8.2f} ms")
        
        # Calculate totals
        total_time = (time.time() - total_start_time) * 1000
        
        # Processing time (excluding file I/O)
        processing_time = (text_extraction_time + drawing_no_time + pipe_class_time + 
                         mat_extraction_time + cut_extraction_time + data_processing_time)
        
        print("=" * 70)
        print(f"TOTAL PROCESSING TIME:          {total_time:8.2f} ms   ({total_time/1000:.3f} seconds)")
        print(f"  - File I/O (open + write):    {(file_open_time + total_csv_time):8.2f} ms")
        print(f"  - Core Processing:            {processing_time:8.2f} ms")
        print("=" * 70)
        
        print(f"\n=== PERFORMANCE SUMMARY ===")
        print(f"Drawing No: {drawing_no}")
        print(f"Pipe Class: {pipe_class}")
        print(f"Material rows extracted: {len(mat_rows_out)}")
        print(f"Cut length pieces extracted: {len(cut_rows_out)}")
        print(f"Total entities processed: {len(text_entities)}")
        print(f"Overall processing rate: {len(text_entities)/(total_time/1000):.0f} entities/second")
        
        print(f"\n=== OUTPUT FILES ===")
        print(f"✓ Materials CSV: {mat_csv_path}")
        print(f"✓ Cut lengths CSV: {cut_csv_path}")
        print(f"✓ Summary CSV: {summary_csv_path}")
        
    except Exception as e:
        total_time = (time.time() - total_start_time) * 1000
        print(f"[ERROR] Processing failed after {total_time:.2f} ms: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_single_file_with_timing()
