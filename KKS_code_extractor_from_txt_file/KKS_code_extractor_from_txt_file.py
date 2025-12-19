#!/usr/bin/env python3
"""
KKS Code Extractor
Extracts KKS codes from P&ID text files handling various formatting patterns.

KKS Pattern: [0-3][A-Z]{3}[0-9]{2}[A-Z]{2}[0-9]{3}
Example: 1HRK30BR210
"""

import re
import pandas as pd
from pathlib import Path


def extract_kks_codes(file_path):
    """
    Extract KKS codes from text file handling multiple formatting patterns.
    
    Patterns handled:
    1. Split across lines: "1 HTX68" + "BR920"
    2. Single line with spaces: "1 HTQ10 CL301"
    3. No spaces: "1HTF14BZ010"
    """
    
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
    
    kks_codes = set()  # Use set to automatically handle duplicates
    
    # Pattern for potential KKS code components
    # Part 1: digit [0-3] + optional space + 3 letters + 2 digits
    part1_pattern = r'\b([0-3])\s*([A-Z]{3})(\d{2})\b'
    # Part 2: 2 letters + 3 digits
    part2_pattern = r'\b([A-Z]{2})(\d{3})\b'
    # Complete KKS on single line (with possible spaces)
    complete_pattern = r'\b([0-3])\s*([A-Z]{3})(\d{2})\s*([A-Z]{2})(\d{3})\b'
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # First, try to find complete KKS codes on the current line
        complete_matches = re.finditer(complete_pattern, line)
        for match in complete_matches:
            # Combine all groups without spaces
            kks_code = ''.join(match.groups())
            kks_codes.add(kks_code)
        
        # Now look for split patterns (part1 on current line, part2 on next line)
        part1_matches = list(re.finditer(part1_pattern, line))
        
        if part1_matches and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            part2_matches = list(re.finditer(part2_pattern, next_line))
            
            # Try to match parts across lines
            for p1_match in part1_matches:
                part1_text = ''.join(p1_match.groups())
                
                # Check if this part1 was already matched as a complete code on this line
                # by seeing if it's a substring of any complete match on this line
                already_complete = False
                for c_match in re.finditer(complete_pattern, line):
                    complete_text = ''.join(c_match.groups())
                    if complete_text.startswith(part1_text):
                        already_complete = True
                        break
                
                if not already_complete:
                    # Look for corresponding part2 on the next line
                    for p2_match in part2_matches:
                        part2_text = ''.join(p2_match.groups())
                        # Combine to form complete KKS code
                        kks_code = part1_text + part2_text
                        kks_codes.add(kks_code)
        
        i += 1
    
    return sorted(list(kks_codes))


def validate_kks_code(code):
    """
    Validate that a code matches the exact KKS pattern.
    Returns True if valid.
    """
    pattern = r'^[0-3][A-Z]{3}\d{2}[A-Z]{2}\d{3}$'
    return bool(re.match(pattern, code))


def main():
    # File paths
    input_file = 'C:\\Users\\szil\\Repos\\excel_wizadry\\KKS_code_extractor_from_txt_file\\softened_water_all_pdf_text.txt'
    output_file = 'C:\\Users\\szil\\Repos\\excel_wizadry\\KKS_code_extractor_from_txt_file\\kks_codes_extracted.xlsx'
    
    # Extract codes
    kks_codes = extract_kks_codes(input_file)
    
    # Validate all codes
    valid_codes = [code for code in kks_codes if validate_kks_code(code)]
    
    # Create DataFrame without headers
    df = pd.DataFrame(valid_codes)
    
    # Save to Excel without headers
    df.to_excel(output_file, index=False, header=False, engine='openpyxl')


if __name__ == "__main__":
    main()