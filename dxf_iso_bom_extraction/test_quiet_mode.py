#!/usr/bin/env python3

import os
import time
from dxf_iso_bom_extraction import main

def test_quiet_mode():
    # Run extraction in quiet mode (no debug output)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Testing quiet mode in: {script_dir}")
    
    # Start timing
    start_time = time.time()
    
    # Run without debug mode (quiet)
    main(script_dir, debug=False)
    
    # End timing
    end_time = time.time()
    processing_time = end_time - start_time
    
    print(f"\nQuiet mode test completed in {processing_time:.2f} seconds")

if __name__ == "__main__":
    test_quiet_mode()
