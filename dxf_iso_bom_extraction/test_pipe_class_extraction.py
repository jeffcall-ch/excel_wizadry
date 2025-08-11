#!/usr/bin/env python3
"""
Test script to verify pipe class extraction functionality.
"""

import os
import sys
from dxf_iso_bom_extraction import process_dxf_file, DEBUG_MODE

def test_pipe_class_extraction():
    """Test pipe class extraction on a single DXF file"""
    
    # Set debug mode for detailed output
    global DEBUG_MODE
    DEBUG_MODE = True
    
    # Look for DXF files in current directory
    dxf_files = []
    for file in os.listdir('.'):
        if file.lower().endswith('.dxf'):
            dxf_files.append(file)
    
    if not dxf_files:
        print("No DXF files found for testing.")
        return
    
    # Use the first DXF file found
    test_file = dxf_files[0]
    print(f"Available DXF files: {len(dxf_files)}")
    for i, file in enumerate(dxf_files):
        print(f"  {i+1}. {file}")
    print(f"\nTesting with: {test_file}")
    print("=" * 60)
    
    # Test the selected file
    try:
        result = process_dxf_file(test_file)
        
        print(f"Drawing No: '{result['drawing_no']}'")
        print(f"Pipe Class: '{result['pipe_class']}'")
        print(f"Material rows: {len(result['mat_rows'])}")
        print(f"Cut length rows: {len(result['cut_rows'])}")
        
        if result['error']:
            print(f"Error: {result['error']}")
        
        # Check if headers include pipe class
        if result['mat_header'] and 'Pipe Class' in result['mat_header']:
            print("✓ ERECTION MATERIALS header includes 'Pipe Class'")
            pipe_class_col = result['mat_header'].index('Pipe Class')
            print(f"  Pipe Class column position: {pipe_class_col}")
            print(f"  Full header: {result['mat_header']}")
        else:
            print("✗ ERECTION MATERIALS header missing 'Pipe Class'")
            if result['mat_header']:
                print(f"  Current header: {result['mat_header']}")
        
        if result['cut_header'] and 'Pipe Class' in result['cut_header']:
            print("✓ CUT PIPE LENGTH header includes 'Pipe Class'")
            pipe_class_col = result['cut_header'].index('Pipe Class')
            print(f"  Pipe Class column position: {pipe_class_col}")
            print(f"  Full header: {result['cut_header']}")
        else:
            print("✗ CUT PIPE LENGTH header missing 'Pipe Class'")
            if result['cut_header']:
                print(f"  Current header: {result['cut_header']}")
        
        # Show sample rows if available
        if result['mat_rows']:
            print(f"\nSample ERECTION MATERIALS row:")
            print(f"  {result['mat_rows'][0]}")
        
        if result['cut_rows']:
            print(f"\nSample CUT PIPE LENGTH row:")
            print(f"  {result['cut_rows'][0]}")
            
    except Exception as e:
        print(f"Error processing {test_file}: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    test_pipe_class_extraction()
