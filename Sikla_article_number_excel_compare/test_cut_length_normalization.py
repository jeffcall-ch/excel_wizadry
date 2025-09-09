import pandas as pd
import numpy as np

# Test the cut length normalization functions
from excel_tab_compare import normalize_cut_length, is_cut_length_equivalent

def test_normalize_cut_length():
    """Test cut length normalization"""
    print("Testing Cut Length Normalization:")
    
    test_cases = [
        # Integer cases
        (1193, 1193),
        (1193.0, 1193),
        ("1193", 1193),
        (" 1193 ", 1193),
        ("1193.0", 1193),
        (" 1193.00 ", 1193),
        
        # Float rounding cases
        (1193.4, 1193),
        (1193.5, 1194),  # Should round up
        (1193.6, 1194),
        ("1193.4", 1193),
        ("1193.5", 1194),
        
        # Empty/NaN cases
        ("", np.nan),
        (None, None),
        (np.nan, np.nan),
        (0, np.nan),
        ("  ", np.nan),
        
        # Non-numeric cases
        ("ABC", "ABC"),
    ]
    
    for input_val, expected in test_cases:
        result = normalize_cut_length(input_val)
        print(f"  Input: {repr(input_val)} -> Output: {repr(result)} (Expected: {repr(expected)})")
        
        # Check if result matches expected (handling NaN comparison)
        if pd.isna(expected) and pd.isna(result):
            print("    ✓ PASS")
        elif result == expected:
            print("    ✓ PASS")
        else:
            print("    ✗ FAIL")

def test_cut_length_equivalent_updated():
    """Test updated cut length equivalence checking"""
    print("\nTesting Updated Cut Length Equivalence:")
    
    test_cases = [
        # Both empty/NaN cases (should match)
        (np.nan, np.nan, True),
        (None, np.nan, True),
        ("", np.nan, True),
        ("", "", True),
        (0, np.nan, True),
        ("  ", np.nan, True),
        
        # One empty, one not (should not match)
        (np.nan, 100, False),
        ("", 200, False),
        (0, 300, False),
        
        # Same values in different formats (should match after normalization)
        (1193, 1193, True),
        (1193.0, 1193, True),
        ("1193", 1193, True),
        (" 1193 ", 1193.0, True),
        ("1193.0", 1193, True),
        
        # Rounding cases
        (1193.4, 1193, True),  # Both round to 1193
        (1193.5, 1194, True),  # 1193.5 rounds to 1194
        (1193.4, 1193.6, True), # Both round to 1193
        (1193.5, 1193.4, False), # 1194 vs 1193
        
        # Different values (should not match)
        (1193, 1194, False),
        (100, 200, False),
    ]
    
    for val1, val2, expected in test_cases:
        result = is_cut_length_equivalent(val1, val2)
        print(f"  {repr(val1)} vs {repr(val2)} -> {result} (Expected: {expected})")
        
        if result == expected:
            print("    ✓ PASS")
        else:
            print("    ✗ FAIL")

if __name__ == "__main__":
    test_normalize_cut_length()
    test_cut_length_equivalent_updated()
