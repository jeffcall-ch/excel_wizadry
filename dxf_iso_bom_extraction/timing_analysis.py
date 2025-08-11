#!/usr/bin/env python3
"""
Comprehensive timing analysis for DXF extraction processing.
Creates a detailed breakdown of all processing steps with context.
"""

import os
import csv
import time
from dxf_iso_bom_extraction import extract_text_entities, find_drawing_no, find_pipe_class, extract_table, convert_cut_length_to_single_row_format
import ezdxf

def comprehensive_timing_analysis():
    filename = "TB020-INOV-2QFB94BR140_1.0_Pipe-Isometric-Drawing-ServiceAir-Lot_General_Piping_Engineering.dxf"
    filepath = os.path.join(".", filename)
    
    if not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        return
    
    # Get file size for context
    file_size = os.path.getsize(filepath)
    file_size_mb = file_size / (1024 * 1024)
    
    print("=" * 80)
    print("                    DXF EXTRACTION TIMING ANALYSIS")
    print("=" * 80)
    print(f"File: {os.path.basename(filename)}")
    print(f"Size: {file_size:,} bytes ({file_size_mb:.2f} MB)")
    print("=" * 80)
    
    # Start total timing
    total_start_time = time.time()
    
    try:
        # === PHASE 1: FILE LOADING ===
        print("PHASE 1: FILE LOADING")
        print("-" * 40)
        
        step_start = time.time()
        doc = ezdxf.readfile(filepath)
        file_open_time = time.time() - step_start
        print(f"  DXF File Parsing:         {file_open_time*1000:10.2f} ms  ({file_open_time:.3f} sec)")
        print(f"  Parsing Rate:             {file_size_mb/file_open_time:10.2f} MB/sec")
        
        # === PHASE 2: CONTENT EXTRACTION ===
        print("\nPHASE 2: CONTENT EXTRACTION")
        print("-" * 40)
        
        step_start = time.time()
        text_entities = extract_text_entities(doc)
        text_extraction_time = time.time() - step_start
        print(f"  Text Entity Extraction:   {text_extraction_time*1000:10.2f} ms  ({len(text_entities)} entities)")
        print(f"  Entity Processing Rate:   {len(text_entities)/text_extraction_time:10.0f} entities/sec")
        
        # === PHASE 3: METADATA EXTRACTION ===
        print("\nPHASE 3: METADATA EXTRACTION")
        print("-" * 40)
        
        step_start = time.time()
        drawing_no = find_drawing_no(text_entities)
        drawing_no_time = time.time() - step_start
        print(f"  Drawing Number Search:    {drawing_no_time*1000:10.2f} ms  ('{drawing_no}')")
        
        step_start = time.time()
        pipe_class = find_pipe_class(text_entities)
        pipe_class_time = time.time() - step_start
        print(f"  Pipe Class Search:        {pipe_class_time*1000:10.2f} ms  ('{pipe_class}')")
        
        # === PHASE 4: TABLE EXTRACTION ===
        print("\nPHASE 4: TABLE EXTRACTION")
        print("-" * 40)
        
        step_start = time.time()
        mat_header, mat_rows = extract_table(text_entities, 'ERECTION MATERIALS')
        mat_extraction_time = time.time() - step_start
        print(f"  Materials Table:          {mat_extraction_time*1000:10.2f} ms  ({len(mat_rows)} rows)")
        
        step_start = time.time()
        cut_header, cut_rows = extract_table(text_entities, 'CUT PIPE LENGTH')
        cut_extraction_time = time.time() - step_start
        print(f"  Cut Lengths Table:        {cut_extraction_time*1000:10.2f} ms  ({len(cut_rows)} original rows)")
        
        # === PHASE 5: DATA PROCESSING ===
        print("\nPHASE 5: DATA PROCESSING")
        print("-" * 40)
        
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
            
        data_processing_time = time.time() - step_start
        print(f"  Format Conversion:        {data_processing_time*1000:10.2f} ms  ({len(cut_rows_out)} final pieces)")
        print(f"  Column Addition:          {data_processing_time*1000:10.2f} ms  (Drawing-No. + Pipe Class)")
        
        # === PHASE 6: FILE OUTPUT ===
        print("\nPHASE 6: FILE OUTPUT")
        print("-" * 40)
        
        csv_start = time.time()
        
        # Write materials CSV
        mat_csv_start = time.time()
        mat_csv_path = f"timing_materials_{drawing_no}.csv"
        with open(mat_csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if mat_rows_out:
                writer.writerow(mat_header_out)
                writer.writerows(mat_rows_out)
            else:
                writer.writerow(['No Data'])
        mat_csv_time = time.time() - mat_csv_start
        mat_size = os.path.getsize(mat_csv_path)
        
        # Write cut lengths CSV
        cut_csv_start = time.time()
        cut_csv_path = f"timing_cut_lengths_{drawing_no}.csv"
        with open(cut_csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if cut_rows_out:
                writer.writerow(cut_header_out)
                writer.writerows(cut_rows_out)
            else:
                writer.writerow(['No Data'])
        cut_csv_time = time.time() - cut_csv_start
        cut_size = os.path.getsize(cut_csv_path)
        
        # Write summary CSV
        summary_csv_start = time.time()
        summary_csv_path = f"timing_summary_{drawing_no}.csv"
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
        summary_csv_time = time.time() - summary_csv_start
        summary_size = os.path.getsize(summary_csv_path)
        
        total_csv_time = time.time() - csv_start
        total_output_size = mat_size + cut_size + summary_size
        
        print(f"  Materials CSV:            {mat_csv_time*1000:10.2f} ms  ({mat_size} bytes, {len(mat_rows_out)} rows)")
        print(f"  Cut Lengths CSV:          {cut_csv_time*1000:10.2f} ms  ({cut_size} bytes, {len(cut_rows_out)} rows)")
        print(f"  Summary CSV:              {summary_csv_time*1000:10.2f} ms  ({summary_size} bytes)")
        print(f"  Total Output:             {total_csv_time*1000:10.2f} ms  ({total_output_size} bytes)")
        
        # === PERFORMANCE SUMMARY ===
        total_time = time.time() - total_start_time
        
        core_processing_time = (text_extraction_time + drawing_no_time + pipe_class_time + 
                              mat_extraction_time + cut_extraction_time + data_processing_time)
        
        print("\n" + "=" * 80)
        print("                          PERFORMANCE SUMMARY")
        print("=" * 80)
        print(f"Total Processing Time:    {total_time*1000:10.2f} ms  ({total_time:.3f} seconds)")
        print(f"  File I/O (read + write): {(file_open_time + total_csv_time)*1000:10.2f} ms  ({((file_open_time + total_csv_time)/total_time)*100:.1f}%)")
        print(f"  Core Processing:         {core_processing_time*1000:10.2f} ms  ({(core_processing_time/total_time)*100:.1f}%)")
        print()
        print(f"Data Throughput:          {file_size_mb/total_time:10.2f} MB/sec")
        print(f"Entity Processing Rate:   {len(text_entities)/total_time:10.0f} entities/sec")
        print(f"Data Compression Ratio:   {total_output_size/file_size:10.1%} (output/input)")
        print()
        print(f"Extraction Results:")
        print(f"  Drawing Number:           {drawing_no}")
        print(f"  Pipe Class:               {pipe_class}")
        print(f"  Material Components:      {len(mat_rows_out)} items")
        print(f"  Cut Pipe Pieces:          {len(cut_rows_out)} pieces")
        print(f"  Total Text Entities:      {len(text_entities)} processed")
        
        print("\n" + "=" * 80)
        print("                            OUTPUT FILES")
        print("=" * 80)
        print(f"✓ {mat_csv_path}")
        print(f"✓ {cut_csv_path}")
        print(f"✓ {summary_csv_path}")
        
    except Exception as e:
        total_time = time.time() - total_start_time
        print(f"\n[ERROR] Processing failed after {total_time*1000:.2f} ms: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    comprehensive_timing_analysis()
