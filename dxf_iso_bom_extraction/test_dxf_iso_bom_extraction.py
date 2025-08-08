import os
import shutil
import tempfile
from dxf_iso_bom_extraction import main

def test_dxf_extraction():
    # Run extraction in the script folder
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Running extraction in: {script_dir}")
    main(script_dir)
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
