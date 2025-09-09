import pandas as pd
import numpy as np

# Test the normalization functions
from excel_tab_compare import normalize_article_number, is_cut_length_equivalent

def test_normalize_article_number():
    """Test article number normalization"""
    print("Testing Article Number Normalization:")
    
    test_cases = [
        ("123", 123),
        (" 123 ", 123),
        ("123.0", 123),
        (" 456.00 ", 456),
        ("ABC123", "ABC123"),
        ("", np.nan),
        (None, None),
        (np.nan, np.nan),
        (123, 123),
        (123.0, 123)
    ]
    
    for input_val, expected in test_cases:
        result = normalize_article_number(input_val)
        print(f"  Input: {repr(input_val)} -> Output: {repr(result)} (Expected: {repr(expected)})")
        
        # Check if result matches expected (handling NaN comparison)
        if pd.isna(expected) and pd.isna(result):
            print("    ✓ PASS")
        elif result == expected:
            print("    ✓ PASS")
        else:
            print("    ✗ FAIL")

def test_cut_length_equivalent():
    """Test cut length equivalence checking"""
    print("\nTesting Cut Length Equivalence:")
    
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
        
        # Both have values (should match only if equal)
        (100, 100, True),
        (100, 200, False),
        (100.5, 100.5, True),
        (100.0, 100, True),
    ]
    
    for val1, val2, expected in test_cases:
        result = is_cut_length_equivalent(val1, val2)
        print(f"  {repr(val1)} vs {repr(val2)} -> {result} (Expected: {expected})")
        
        if result == expected:
            print("    ✓ PASS")
        else:
            print("    ✗ FAIL")

if __name__ == "__main__":
    test_normalize_article_number()
    test_cut_length_equivalent()
