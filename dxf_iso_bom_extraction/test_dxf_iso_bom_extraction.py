import os
import shutil
import tempfile
import time
from dxf_iso_bom_extraction import main

def test_dxf_extraction():
    # Run extraction in the script folder
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Running extraction in: {script_dir}")
    
    # Count DXF files before processing
    dxf_count = 0
    for root, _, files in os.walk(script_dir):
        for file in files:
            if file.lower().endswith('.dxf'):
                dxf_count += 1
    
    print(f"Found {dxf_count} DXF files to process")
    
    # Start timing
    start_time = time.time()
    
    main(script_dir)
    
    # End timing
    end_time = time.time()
    processing_time = end_time - start_time
    
    print(f"\n=== PERFORMANCE SUMMARY ===")
    print(f"Files processed: {dxf_count}")
    print(f"Total time: {processing_time:.2f} seconds")
    if dxf_count > 0:
        print(f"Average time per file: {processing_time/dxf_count:.2f} seconds")
    
    # Check output files
    for fname in ['all_materials.csv', 'all_cut_lengths.csv', 'summary.csv']:
        out_path = os.path.join(script_dir, fname)
        print(f"Checking output file: {out_path}")
        if not os.path.exists(out_path):
            print(f"[ERROR] Missing output file: {fname}")
        else:
            with open(out_path, encoding='utf-8') as f:
                lines = f.readlines()
                print(f"{fname}: {len(lines)} lines")
                if len(lines) == 0:
                    print(f"[ERROR] Output file {fname} is empty")
    print('DXF extraction test completed.')

if __name__ == "__main__":
    test_dxf_extraction()
