import pandas as pd
import os
from datetime import datetime
import numpy as np
import time

# Import our functions from the main script
from excel_tab_compare import (
    load_excel_tabs, find_kks_column, create_kks_lookup_dict, 
    find_kks_matches_optimized, find_exact_match_by_hierarchy_vectorized
)

def test_small_subset():
    """Test the optimized functions on a small subset"""
    file_path = r"C:\Users\szil\Repos\excel_wizadry\Sikla_article_number_excel_compare\C4_and_C5_BOM_01.09.2025_to_compare_with_dwg_bom_extract.xlsx"
    
    print("Loading Excel file...")
    baseline_df, comparison_df, sheet_names = load_excel_tabs(file_path)
    
    if baseline_df is None or comparison_df is None:
        print("Failed to load Excel file")
        return
    
    # Test with first 100 rows only
    baseline_subset = baseline_df.head(100).copy()
    
    baseline_kks_col = find_kks_column(baseline_subset)
    comparison_kks_col = find_kks_column(comparison_df)
    
    print(f"Testing with {len(baseline_subset)} baseline rows against {len(comparison_df)} comparison rows")
    print(f"Baseline KKS column: {baseline_kks_col}")
    print(f"Comparison KKS column: {comparison_kks_col}")
    
    # Build KKS lookup
    start_time = time.time()
    kks_lookup = create_kks_lookup_dict(comparison_df, comparison_kks_col)
    lookup_time = time.time() - start_time
    print(f"KKS lookup built in {lookup_time:.2f} seconds with {len(kks_lookup)} entries")
    
    # Test matching on subset
    required_cols = ['ARTICLE_NUMBER', 'QTY', 'CUT_LENGTH']
    for col in required_cols:
        if col not in comparison_df.columns:
            comparison_df[col] = np.nan
        if col not in baseline_subset.columns:
            baseline_subset[col] = np.nan
    
    comparison_subset = comparison_df[required_cols + [comparison_kks_col]].copy()
    
    match_results = []
    match_details = []
    
    start_time = time.time()
    for idx in range(len(baseline_subset)):
        baseline_row = baseline_subset.iloc[idx]
        baseline_kks = baseline_row[baseline_kks_col]
        
        # Find KKS matches
        matched_indices = find_kks_matches_optimized(baseline_kks, kks_lookup)
        
        if not matched_indices:
            match_results.append('no match')
            match_details.append('No KKS/SU partial match found')
        else:
            best_match_idx, match_detail = find_exact_match_by_hierarchy_vectorized(
                baseline_row, comparison_subset, matched_indices
            )
            
            if best_match_idx is not None:
                match_results.append('match')
                match_details.append(match_detail)
            else:
                match_results.append('no match')
                match_details.append(match_detail)
    
    processing_time = time.time() - start_time
    print(f"Processed {len(baseline_subset)} rows in {processing_time:.2f} seconds")
    print(f"Average time per row: {processing_time/len(baseline_subset)*1000:.2f} ms")
    
    # Show match statistics
    match_counts = pd.Series(match_results).value_counts()
    print(f"\nMatch results:")
    for status, count in match_counts.items():
        print(f"  {status}: {count} ({count/len(baseline_subset)*100:.1f}%)")
    
    # Show some example details
    print(f"\nFirst 5 match details:")
    for i in range(min(5, len(match_details))):
        print(f"  Row {i}: {match_results[i]} - {match_details[i]}")

if __name__ == "__main__":
    test_small_subset()
