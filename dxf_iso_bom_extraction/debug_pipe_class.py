#!/usr/bin/env python3
"""
Debug script to analyze pipe class extraction in detail.
"""

import os
import sys
from dxf_iso_bom_extraction import extract_text_entities, find_pipe_class
import ezdxf

def debug_pipe_class_extraction():
    """Debug pipe class extraction to see what text entities are found"""
    
    # Use the same file as before
    filename = "TB020-INOV-2QFB94BR140_1.0_Pipe-Isometric-Drawing-ServiceAir-Lot_General_Piping_Engineering.dxf"
    
    if not os.path.exists(filename):
        print(f"File not found: {filename}")
        return
    
    print(f"=== DEBUG PIPE CLASS EXTRACTION ===")
    print(f"File: {filename}\n")
    
    try:
        doc = ezdxf.readfile(filename)
        text_entities = extract_text_entities(doc)
        
        print(f"Total text entities found: {len(text_entities)}\n")
        
        # Look for text entities containing "pipe", "class", "AHDX", "IFC", etc.
        relevant_terms = ['pipe', 'class', 'AHDX', 'IFC', 'design', 'data']
        relevant_entities = []
        
        for text, x, y in text_entities:
            text_lower = text.lower().replace(' ', '').replace('\n', '')
            for term in relevant_terms:
                if term.lower() in text_lower:
                    relevant_entities.append((text, x, y, term))
                    break
        
        print(f"=== RELEVANT TEXT ENTITIES ===")
        for text, x, y, term in relevant_entities:
            print(f"Text: '{text}' | X: {x:.1f} | Y: {y:.1f} | Match: {term}")
        
        print(f"\n=== PIPE CLASS DETECTION PROCESS ===")
        
        # Enable debug mode temporarily
        import dxf_iso_bom_extraction
        dxf_iso_bom_extraction.DEBUG_MODE = True
        
        pipe_class = find_pipe_class(text_entities)
        print(f"\nFinal result: '{pipe_class}'")
        
        # Also show entities in bottom area
        print(f"\n=== BOTTOM AREA ENTITIES (Y < 0) ===")
        bottom_entities = [(t, x, y) for t, x, y in text_entities if y < 0]
        bottom_entities.sort(key=lambda e: e[2])  # Sort by Y coordinate
        
        for text, x, y in bottom_entities[:20]:  # Show first 20
            print(f"'{text}' at X:{x:.1f}, Y:{y:.1f}")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    debug_pipe_class_extraction()
