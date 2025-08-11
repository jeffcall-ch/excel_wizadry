#!/usr/bin/env python3
"""
Threading-based parallel processing for I/O-bound DXF processing.
Since DXF parsing involves significant I/O, threading can be effective.
"""

import os
import csv
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue
from dxf_iso_bom_extraction import process_dxf_file, DEBUG_MODE

class ThreadSafeCollector:
    """Thread-safe collector for aggregating results."""
    
    def __init__(self):
        self.lock = threading.Lock()
        self.material_rows = []
        self.cut_rows = []
        self.mat_header = None
        self.cut_header = None
        self.summary = []
        self.completed_count = 0
        
    def add_result(self, result):
        """Add a processing result in a thread-safe manner."""
        with self.lock:
            # Collect material data
            if result['mat_rows']:
                if not self.mat_header:
                    self.mat_header = result['mat_header']
                self.material_rows.extend(result['mat_rows'])
            
            # Collect cut length data
            if result['cut_rows']:
                if not self.cut_header:
                    self.cut_header = result['cut_header']
                self.cut_rows.extend(result['cut_rows'])
            
            # Add to summary
            self.summary.append({
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
            
            self.completed_count += 1

def process_file_threaded(file_path, collector, total_files, progress_queue):
    """Process a single file in a thread and update the collector."""
    start_time = time.time()
    filename = os.path.basename(file_path)
    
    try:
        # Process the file
        result = process_dxf_file(file_path)
        processing_time = time.time() - start_time
        
        # Add metadata
        result['processing_time'] = processing_time
        result['filename'] = filename
        result['file_path'] = file_path
        
        # Add to collector
        collector.add_result(result)
        
        # Report progress
        progress_queue.put((collector.completed_count, filename, processing_time, None))
        
        return result
        
    except Exception as e:
        processing_time = time.time() - start_time
        
        # Create error result
        error_result = {
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
        
        collector.add_result(error_result)
        progress_queue.put((collector.completed_count, filename, processing_time, str(e)))
        
        return error_result

def threaded_process_directory(directory, max_workers=None, debug=False):
    """Process DXF files using threading."""
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
    
    # Determine optimal number of threads
    if max_workers is None:
        max_workers = min(len(dxf_files), 8)  # Limit to 8 threads for I/O bound tasks
    
    total_files = len(dxf_files)
    print(f"Processing {total_files} DXF files using {max_workers} threads...")
    
    # Initialize collector and progress tracking
    collector = ThreadSafeCollector()
    progress_queue = Queue()
    overall_start = time.time()
    
    # Process files with threading
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Submit all tasks
        futures = [executor.submit(process_file_threaded, file_path, collector, total_files, progress_queue) 
                  for file_path in dxf_files]
        
        # Monitor progress
        while collector.completed_count < total_files:
            try:
                completed, filename, proc_time, error = progress_queue.get(timeout=1)
                if error:
                    print(f"[{completed}/{total_files}] Failed: {filename} ({proc_time:.2f}s) - {error}")
                else:
                    print(f"[{completed}/{total_files}] Completed: {filename} ({proc_time:.2f}s)")
            except:
                continue  # Timeout, continue monitoring
        
        # Wait for all to complete
        for future in as_completed(futures):
            pass  # Results already collected by threads
    
    overall_time = time.time() - overall_start
    
    # Write output files
    write_threaded_output_files(directory, collector)
    
    # Performance report
    successful_files = sum(1 for s in collector.summary if not s['error'])
    total_processing_time = sum(s['processing_time'] for s in collector.summary)
    
    print_threaded_performance_report(total_files, successful_files, overall_time, 
                                    total_processing_time, max_workers, 
                                    len(collector.material_rows), len(collector.cut_rows))

def write_threaded_output_files(directory, collector):
    """Write results to files."""
    # Materials
    out_path = os.path.join(directory, 'threaded_all_materials.csv')
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if collector.material_rows and collector.mat_header:
            writer.writerow(collector.mat_header)
            writer.writerows(collector.material_rows)
        else:
            writer.writerow(['No Data'])
    
    # Cut lengths
    out_path = os.path.join(directory, 'threaded_all_cut_lengths.csv')
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if collector.cut_rows and collector.cut_header:
            writer.writerow(collector.cut_header)
            writer.writerows(collector.cut_rows)
        else:
            writer.writerow(['No Data'])
    
    # Summary
    out_path = os.path.join(directory, 'threaded_summary.csv')
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        fieldnames = ['filename', 'file_path', 'drawing_no', 'pipe_class', 'mat_rows', 'cut_rows', 
                     'mat_missing', 'cut_missing', 'error', 'processing_time']
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(collector.summary)

def print_threaded_performance_report(total_files, successful_files, overall_time, 
                                    total_processing_time, max_workers, total_materials, total_cuts):
    """Print performance report for threading approach."""
    print("\n" + "=" * 80)
    print("                         THREADED PROCESSING REPORT")
    print("=" * 80)
    
    print(f"Files Processed:              {total_files}")
    print(f"Successful:                   {successful_files}")
    print(f"Failed:                       {total_files - successful_files}")
    print(f"Thread Workers:               {max_workers}")
    
    print(f"\nTiming Analysis:")
    print(f"Overall Wall Time:            {overall_time:.2f} seconds")
    print(f"Total Processing Time:        {total_processing_time:.2f} seconds")
    print(f"Average per File:             {total_processing_time/max(successful_files, 1):.2f} seconds")
    
    if successful_files > 0:
        speedup = total_processing_time / overall_time
        efficiency = speedup / max_workers * 100
        
        print(f"\nThreading Efficiency:")
        print(f"Speedup Factor:               {speedup:.2f}x")
        print(f"Threading Efficiency:         {efficiency:.1f}%")
        print(f"Time Saved:                   {total_processing_time - overall_time:.2f} seconds")
    
    print(f"\nData Extracted:")
    print(f"Total Material Rows:          {total_materials}")
    print(f"Total Cut Length Pieces:      {total_cuts}")
    print(f"Processing Rate:              {total_files/overall_time:.2f} files/second")
    
    print("=" * 80)

if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print('Usage: python threaded_dxf_extraction.py <directory> [--threads N] [--debug]')
        sys.exit(1)
    
    directory = sys.argv[1]
    max_workers = None
    debug_mode = '--debug' in sys.argv
    
    if '--threads' in sys.argv:
        try:
            threads_idx = sys.argv.index('--threads') + 1
            max_workers = int(sys.argv[threads_idx])
        except (IndexError, ValueError):
            print("Invalid --threads argument. Using default.")
    
    threaded_process_directory(directory, max_workers, debug_mode)
