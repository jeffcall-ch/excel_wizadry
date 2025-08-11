#!/usr/bin/env python3
"""
Asynchronous processing using asyncio for concurrent DXF processing.
This approach is good for I/O-bound operations with modern Python async features.
"""

import os
import csv
import time
import asyncio
from concurrent.futures import ProcessPoolExecutor
from dxf_iso_bom_extraction import process_dxf_file

async def process_file_async(file_path, semaphore):
    """Process a single DXF file asynchronously."""
    async with semaphore:  # Limit concurrent operations
        loop = asyncio.get_event_loop()
        
        # Run the CPU-bound operation in a process pool
        with ProcessPoolExecutor() as executor:
            start_time = time.time()
            filename = os.path.basename(file_path)
            
            try:
                # Run the blocking operation in a separate process
                result = await loop.run_in_executor(executor, process_dxf_file, file_path)
                processing_time = time.time() - start_time
                
                result['processing_time'] = processing_time
                result['filename'] = filename
                result['file_path'] = file_path
                
                return result
                
            except Exception as e:
                processing_time = time.time() - start_time
                return {
                    'filename': filename,
                    'file_path': file_path,
                    'drawing_no': '',
                    'pipe_class': '',
                    'mat_header': [],
                    'mat_rows': [],
                    'cut_header': [],
                    'cut_rows': [],
                    'error': str(e),
                    'processing_time': processing_time
                }

async def async_process_directory(directory, max_concurrent=None, debug=False):
    """Process DXF files using async/await pattern."""
    # Find all DXF files
    dxf_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.dxf'):
                dxf_files.append(os.path.join(root, file))
    
    if not dxf_files:
        print("No DXF files found.")
        return
    
    # Determine concurrent limit
    if max_concurrent is None:
        max_concurrent = min(len(dxf_files), os.cpu_count())
    
    total_files = len(dxf_files)
    print(f"Processing {total_files} DXF files with {max_concurrent} concurrent operations...")
    
    # Create semaphore to limit concurrent operations
    semaphore = asyncio.Semaphore(max_concurrent)
    overall_start = time.time()
    
    # Create tasks for all files
    tasks = [process_file_async(file_path, semaphore) for file_path in dxf_files]
    
    # Process with progress monitoring
    results = []
    completed_count = 0
    
    for coro in asyncio.as_completed(tasks):
        result = await coro
        results.append(result)
        completed_count += 1
        
        filename = result['filename']
        processing_time = result['processing_time']
        
        if result['error']:
            print(f"[{completed_count}/{total_files}] Failed: {filename} ({processing_time:.2f}s) - {result['error']}")
        else:
            print(f"[{completed_count}/{total_files}] Completed: {filename} ({processing_time:.2f}s)")
    
    overall_time = time.time() - overall_start
    
    # Aggregate and save results
    await save_async_results(directory, results, overall_time, max_concurrent)

async def save_async_results(directory, results, overall_time, max_concurrent):
    """Save results and print performance report."""
    # Aggregate data
    material_rows = []
    cut_rows = []
    mat_header = None
    cut_header = None
    summary = []
    
    for result in results:
        if result['mat_rows']:
            if not mat_header:
                mat_header = result['mat_header']
            material_rows.extend(result['mat_rows'])
        
        if result['cut_rows']:
            if not cut_header:
                cut_header = result['cut_header']
            cut_rows.extend(result['cut_rows'])
        
        summary.append({
            'filename': result['filename'],
            'file_path': result['file_path'],
            'drawing_no': result['drawing_no'],
            'pipe_class': result['pipe_class'],
            'mat_rows': len(result['mat_rows']),
            'cut_rows': len(result['cut_rows']),
            'mat_missing': not bool(result['mat_rows']),
            'cut_missing': not bool(result['cut_rows']),
            'error': result['error'],
            'processing_time': result['processing_time']
        })
    
    # Write output files
    await write_async_files(directory, material_rows, cut_rows, summary, mat_header, cut_header)
    
    # Performance report
    successful_files = sum(1 for r in results if not r['error'])
    total_processing_time = sum(r['processing_time'] for r in results)
    
    print_async_performance_report(len(results), successful_files, overall_time, 
                                 total_processing_time, max_concurrent, 
                                 len(material_rows), len(cut_rows))

async def write_async_files(directory, material_rows, cut_rows, summary, mat_header, cut_header):
    """Write results to files asynchronously."""
    loop = asyncio.get_event_loop()
    
    # Define file writing functions
    def write_materials():
        out_path = os.path.join(directory, 'async_all_materials.csv')
        with open(out_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if material_rows and mat_header:
                writer.writerow(mat_header)
                writer.writerows(material_rows)
            else:
                writer.writerow(['No Data'])
    
    def write_cuts():
        out_path = os.path.join(directory, 'async_all_cut_lengths.csv')
        with open(out_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if cut_rows and cut_header:
                writer.writerow(cut_header)
                writer.writerows(cut_rows)
            else:
                writer.writerow(['No Data'])
    
    def write_summary():
        out_path = os.path.join(directory, 'async_summary.csv')
        with open(out_path, 'w', newline='', encoding='utf-8') as f:
            fieldnames = ['filename', 'file_path', 'drawing_no', 'pipe_class', 'mat_rows', 'cut_rows', 
                         'mat_missing', 'cut_missing', 'error', 'processing_time']
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(summary)
    
    # Run file operations concurrently
    await asyncio.gather(
        loop.run_in_executor(None, write_materials),
        loop.run_in_executor(None, write_cuts),
        loop.run_in_executor(None, write_summary)
    )

def print_async_performance_report(total_files, successful_files, overall_time, 
                                 total_processing_time, max_concurrent, total_materials, total_cuts):
    """Print performance report for async approach."""
    print("\n" + "=" * 80)
    print("                         ASYNC PROCESSING REPORT")
    print("=" * 80)
    
    print(f"Files Processed:              {total_files}")
    print(f"Successful:                   {successful_files}")
    print(f"Failed:                       {total_files - successful_files}")
    print(f"Concurrent Operations:        {max_concurrent}")
    
    print(f"\nTiming Analysis:")
    print(f"Overall Wall Time:            {overall_time:.2f} seconds")
    print(f"Total Processing Time:        {total_processing_time:.2f} seconds")
    print(f"Average per File:             {total_processing_time/max(successful_files, 1):.2f} seconds")
    
    if successful_files > 0:
        speedup = total_processing_time / overall_time
        efficiency = speedup / max_concurrent * 100
        
        print(f"\nAsync Efficiency:")
        print(f"Speedup Factor:               {speedup:.2f}x")
        print(f"Async Efficiency:             {efficiency:.1f}%")
        print(f"Time Saved:                   {total_processing_time - overall_time:.2f} seconds")
    
    print(f"\nData Extracted:")
    print(f"Total Material Rows:          {total_materials}")
    print(f"Total Cut Length Pieces:      {total_cuts}")
    print(f"Processing Rate:              {total_files/overall_time:.2f} files/second")
    
    print("=" * 80)

def main():
    """Main entry point for async processing."""
    import sys
    
    if len(sys.argv) < 2:
        print('Usage: python async_dxf_extraction.py <directory> [--concurrent N] [--debug]')
        print('  <directory>: Root folder containing DXF files')
        print('  --concurrent N: Number of concurrent operations (default: CPU count)')
        print('  --debug: Show detailed debug output')
        return
    
    directory = sys.argv[1]
    max_concurrent = None
    debug_mode = '--debug' in sys.argv
    
    if '--concurrent' in sys.argv:
        try:
            concurrent_idx = sys.argv.index('--concurrent') + 1
            max_concurrent = int(sys.argv[concurrent_idx])
        except (IndexError, ValueError):
            print("Invalid --concurrent argument. Using default.")
    
    if not os.path.exists(directory):
        print(f"Directory not found: {directory}")
        return
    
    # Run the async processing
    asyncio.run(async_process_directory(directory, max_concurrent, debug_mode))

if __name__ == '__main__':
    main()
