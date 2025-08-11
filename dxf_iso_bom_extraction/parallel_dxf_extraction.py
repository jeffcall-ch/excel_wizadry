#!/usr/bin/env python3
"""
Parallel DXF processing using multiprocessing for CPU-bound tasks.
This approach processes multiple DXF files simultaneously using separate processes.
"""

import os
import csv
import time
import multiprocessing as mp
from concurrent.futures import ProcessPoolExecutor, as_completed
from dxf_iso_bom_extraction import process_dxf_file, DEBUG_MODE

def process_file_with_timing(file_path):
    """
    Process a single DXF file and return results with timing information.
    This function will run in a separate process.
    """
    start_time = time.time()
    filename = os.path.basename(file_path)
    
    try:
        # Process the file
        result = process_dxf_file(file_path)
        processing_time = time.time() - start_time
        
        # Add timing info to result
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

def parallel_process_directory(directory, max_workers=None, debug=False):
    """
    Process all DXF files in a directory using parallel processing.
    
    Args:
        directory: Root directory containing DXF files
        max_workers: Maximum number of parallel processes (default: CPU count)
        debug: Enable debug output
    """
    global DEBUG_MODE
    DEBUG_MODE = debug
    
    # Find all DXF files
    dxf_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.dxf'):
                dxf_files.append(os.path.join(root, file))
    
    if not dxf_files:
        print("No DXF files found.")
        return
    
    # Determine optimal number of workers
    if max_workers is None:
        max_workers = min(len(dxf_files), mp.cpu_count())
    
    total_files = len(dxf_files)
    print(f"Processing {total_files} DXF files using {max_workers} parallel processes...")
    print(f"Available CPU cores: {mp.cpu_count()}")
    
    # Track overall timing
    overall_start = time.time()
    
    # Process files in parallel
    results = []
    completed_count = 0
    
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        # Submit all jobs
        future_to_file = {executor.submit(process_file_with_timing, file_path): file_path 
                         for file_path in dxf_files}
        
        # Process completed jobs as they finish
        for future in as_completed(future_to_file):
            file_path = future_to_file[future]
            completed_count += 1
            
            try:
                result = future.result()
                results.append(result)
                
                filename = os.path.basename(file_path)
                processing_time = result.get('processing_time', 0)
                print(f"[{completed_count}/{total_files}] Completed: {filename} ({processing_time:.2f}s)")
                
            except Exception as e:
                print(f"[{completed_count}/{total_files}] Failed: {os.path.basename(file_path)} - {e}")
                # Add error result
                results.append({
                    'filename': os.path.basename(file_path),
                    'file_path': file_path,
                    'drawing_no': '',
                    'pipe_class': '',
                    'mat_header': [],
                    'mat_rows': [],
                    'cut_header': [],
                    'cut_rows': [],
                    'error': str(e),
                    'processing_time': 0
                })
    
    overall_time = time.time() - overall_start
    
    # Aggregate results
    material_rows = []
    cut_rows = []
    mat_header = None
    cut_header = None
    summary = []
    
    total_processing_time = 0
    successful_files = 0
    
    for result in results:
        # Collect data
        if result['mat_rows']:
            if not mat_header:
                mat_header = result['mat_header']
            material_rows.extend(result['mat_rows'])
        
        if result['cut_rows']:
            if not cut_header:
                cut_header = result['cut_header']
            cut_rows.extend(result['cut_rows'])
        
        # Summary statistics
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
        
        if not result['error']:
            successful_files += 1
            total_processing_time += result['processing_time']
    
    # Write output files
    write_output_files(directory, material_rows, cut_rows, summary, mat_header, cut_header)
    
    # Performance report
    print_performance_report(total_files, successful_files, overall_time, 
                           total_processing_time, max_workers, len(material_rows), len(cut_rows))

def write_output_files(directory, material_rows, cut_rows, summary, mat_header, cut_header):
    """Write aggregated results to CSV files."""
    # Write all_materials.csv
    out_path = os.path.join(directory, 'parallel_all_materials.csv')
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if material_rows and mat_header:
            writer.writerow(mat_header)
            writer.writerows(material_rows)
        else:
            writer.writerow(['No Data'])
    
    # Write all_cut_lengths.csv
    out_path = os.path.join(directory, 'parallel_all_cut_lengths.csv')
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if cut_rows and cut_header:
            writer.writerow(cut_header)
            writer.writerows(cut_rows)
        else:
            writer.writerow(['No Data'])
    
    # Write summary.csv with timing information
    out_path = os.path.join(directory, 'parallel_summary.csv')
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        fieldnames = ['filename', 'file_path', 'drawing_no', 'pipe_class', 'mat_rows', 'cut_rows', 
                     'mat_missing', 'cut_missing', 'error', 'processing_time']
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(summary)

def print_performance_report(total_files, successful_files, overall_time, total_processing_time, 
                           max_workers, total_materials, total_cuts):
    """Print detailed performance analysis."""
    print("\n" + "=" * 80)
    print("                        PARALLEL PROCESSING REPORT")
    print("=" * 80)
    
    print(f"Files Processed:              {total_files}")
    print(f"Successful:                   {successful_files}")
    print(f"Failed:                       {total_files - successful_files}")
    print(f"Parallel Workers:             {max_workers}")
    
    print(f"\nTiming Analysis:")
    print(f"Overall Wall Time:            {overall_time:.2f} seconds")
    print(f"Total Processing Time:        {total_processing_time:.2f} seconds")
    print(f"Average per File:             {total_processing_time/max(successful_files, 1):.2f} seconds")
    
    # Calculate efficiency
    if successful_files > 0:
        sequential_estimate = total_processing_time
        speedup = sequential_estimate / overall_time
        efficiency = speedup / max_workers * 100
        
        print(f"\nParallel Efficiency:")
        print(f"Estimated Sequential Time:    {sequential_estimate:.2f} seconds")
        print(f"Speedup Factor:               {speedup:.2f}x")
        print(f"Parallel Efficiency:          {efficiency:.1f}%")
        print(f"Time Saved:                   {sequential_estimate - overall_time:.2f} seconds")
    
    print(f"\nData Extracted:")
    print(f"Total Material Rows:          {total_materials}")
    print(f"Total Cut Length Pieces:      {total_cuts}")
    print(f"Processing Rate:              {total_files/overall_time:.2f} files/second")
    
    print("=" * 80)

def main():
    """Main entry point for parallel processing."""
    import sys
    
    if len(sys.argv) < 2:
        print('Usage: python parallel_dxf_extraction.py <directory> [--workers N] [--debug]')
        print('  <directory>: Root folder containing DXF files')
        print('  --workers N: Number of parallel processes (default: CPU count)')
        print('  --debug: Show detailed debug output')
        return
    
    directory = sys.argv[1]
    max_workers = None
    debug_mode = '--debug' in sys.argv
    
    # Parse workers argument
    if '--workers' in sys.argv:
        try:
            workers_idx = sys.argv.index('--workers') + 1
            max_workers = int(sys.argv[workers_idx])
        except (IndexError, ValueError):
            print("Invalid --workers argument. Using default.")
    
    if not os.path.exists(directory):
        print(f"Directory not found: {directory}")
        return
    
    parallel_process_directory(directory, max_workers, debug_mode)

if __name__ == '__main__':
    main()
