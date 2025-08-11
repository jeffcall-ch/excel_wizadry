#!/usr/bin/env python3
"""
Quick parallel processing demo using just 2 files to show the concept.
"""

import os
import time
from concurrent.futures import ProcessPoolExecutor
from dxf_iso_bom_extraction import process_dxf_file

def demo_parallel_processing():
    """Demonstrate parallel vs sequential processing."""
    
    # Find available DXF files
    dxf_files = [f for f in os.listdir('.') if f.lower().endswith('.dxf')][:2]  # Just take 2 files
    
    if len(dxf_files) < 2:
        print("Need at least 2 DXF files for demo.")
        return
    
    print("=" * 60)
    print("PARALLEL PROCESSING DEMONSTRATION")
    print("=" * 60)
    print(f"Files to process: {dxf_files}")
    
    # Sequential processing
    print("\n1. SEQUENTIAL PROCESSING:")
    print("-" * 30)
    sequential_start = time.time()
    
    for i, file_path in enumerate(dxf_files, 1):
        file_start = time.time()
        result = process_dxf_file(file_path)
        file_time = time.time() - file_start
        print(f"  File {i}: {os.path.basename(file_path)} - {file_time:.2f}s")
    
    sequential_total = time.time() - sequential_start
    print(f"  Total Sequential Time: {sequential_total:.2f}s")
    
    # Parallel processing
    print("\n2. PARALLEL PROCESSING (2 workers):")
    print("-" * 30)
    parallel_start = time.time()
    
    with ProcessPoolExecutor(max_workers=2) as executor:
        futures = {executor.submit(process_dxf_file, file_path): file_path for file_path in dxf_files}
        
        for future in futures:
            file_path = futures[future]
            result = future.result()
            print(f"  Completed: {os.path.basename(file_path)}")
    
    parallel_total = time.time() - parallel_start
    print(f"  Total Parallel Time: {parallel_total:.2f}s")
    
    # Performance comparison
    print("\n3. PERFORMANCE COMPARISON:")
    print("-" * 30)
    speedup = sequential_total / parallel_total
    time_saved = sequential_total - parallel_total
    efficiency = speedup / 2 * 100  # 2 workers
    
    print(f"  Sequential Time:    {sequential_total:.2f}s")
    print(f"  Parallel Time:      {parallel_total:.2f}s")
    print(f"  Speedup Factor:     {speedup:.2f}x")
    print(f"  Time Saved:         {time_saved:.2f}s ({time_saved/sequential_total*100:.1f}%)")
    print(f"  Parallel Efficiency: {efficiency:.1f}%")
    
    print("\n" + "=" * 60)

if __name__ == '__main__':
    demo_parallel_processing()
