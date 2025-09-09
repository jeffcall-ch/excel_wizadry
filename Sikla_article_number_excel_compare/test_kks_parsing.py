import pandas as pd
import numpy as np

# Test the KKS parsing functions
from excel_tab_compare import parse_kks_value

def test_parse_kks_value():
    """Test KKS value parsing"""
    print("Testing KKS Value Parsing:")
    
    test_cases = [
        # List format from comparison data
        ("['0EAX10BQ001', '0EAX10BQ001', '0EAX10BQ001']", ['0EAX10BQ001', '0EAX10BQ001', '0EAX10BQ001']),
        ("['0EAX10BQ002', '0EAX10BQ002', '0EAX10BQ002']", ['0EAX10BQ002', '0EAX10BQ002', '0EAX10BQ002']),
        
        # Plain string format from baseline data
        ("/0CFX50BQ003/SU", ["/0CFX50BQ003/SU"]),
        ("/0CFX50BQ005/SU", ["/0CFX50BQ005/SU"]),
        
        # Edge cases
        ("", []),
        ("[]", []),
        ("['']", []),
        (np.nan, []),
        (None, []),
        
        # Mixed cases
        ("['ABC', 'DEF']", ['ABC', 'DEF']),
        ("Single_Value", ["Single_Value"]),
    ]
    
    for input_val, expected in test_cases:
        result = parse_kks_value(input_val)
        print(f"  Input: {repr(input_val)} -> Output: {result}")
        print(f"    Expected: {expected}")
        
        if result == expected:
            print("    ✓ PASS")
        else:
            print("    ✗ FAIL")
        print()

def test_kks_matching_logic():
    """Test how the new KKS matching should work"""
    print("Testing KKS Matching Logic:")
    
    # Simulate the real data scenario
    baseline_kks = "/0CFX50BQ003/SU"
    comparison_kks_samples = [
        "['0EAX10BQ001', '0EAX10BQ001', '0EAX10BQ001']",
        "['0CFX50BQ003', '0CFX50BQ003', '0CFX50BQ003']",  # This should match!
        "['0EAX10BQ002', '0EAX10BQ002', '0EAX10BQ002']",
    ]
    
    print(f"Baseline KKS: {baseline_kks}")
    baseline_parsed = parse_kks_value(baseline_kks)
    print(f"Baseline parsed: {baseline_parsed}")
    
    for i, comp_kks in enumerate(comparison_kks_samples):
        comp_parsed = parse_kks_value(comp_kks)
        print(f"\nComparison {i+1}: {comp_kks}")
        print(f"Comparison parsed: {comp_parsed}")
        
        # Test partial matching
        match_found = False
        for base_part in baseline_parsed:
            for comp_part in comp_parsed:
                if base_part in comp_part or comp_part in base_part:
                    match_found = True
                    print(f"  MATCH FOUND: '{base_part}' matches '{comp_part}'")
        
        if not match_found:
            print("  NO MATCH")

if __name__ == "__main__":
    test_parse_kks_value()
    test_kks_matching_logic()
