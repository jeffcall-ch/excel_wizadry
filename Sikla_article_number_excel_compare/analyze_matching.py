import pandas as pd
import numpy as np
from excel_tab_compare import (
    load_excel_tabs, find_kks_column, create_kks_lookup_dict, 
    find_kks_matches_optimized, normalize_article_number
)

def analyze_matching_issues():
    """Analyze why we have such low match rates"""
    file_path = r"C:\Users\szil\Repos\excel_wizadry\Sikla_article_number_excel_compare\C4_and_C5_BOM_01.09.2025_to_compare_with_dwg_bom_extract.xlsx"
    
    print("Loading Excel file...")
    baseline_df, comparison_df, sheet_names = load_excel_tabs(file_path)
    
    if baseline_df is None or comparison_df is None:
        print("Failed to load Excel file")
        return
    
    baseline_kks_col = find_kks_column(baseline_df)
    comparison_kks_col = find_kks_column(comparison_df)
    
    print(f"\n=== DATA OVERVIEW ===")
    print(f"Baseline rows: {len(baseline_df)}")
    print(f"Comparison rows: {len(comparison_df)}")
    print(f"Baseline KKS column: {baseline_kks_col}")
    print(f"Comparison KKS column: {comparison_kks_col}")
    
    print(f"\n=== SAMPLE DATA ===")
    print("First 5 baseline rows:")
    print(baseline_df[['ARTICLE_NUMBER', 'QTY', 'CUT_LENGTH', 'KKS/SU']].head())
    print("\nFirst 5 comparison rows:")
    print(comparison_df[['ARTICLE_NUMBER', 'QTY', 'CUT_LENGTH', 'KKS/SU']].head())
    
    # Analyze KKS matching
    print(f"\n=== KKS/SU ANALYSIS ===")
    baseline_kks_unique = set(str(x).strip() for x in baseline_df[baseline_kks_col].dropna())
    comparison_kks_unique = set(str(x).strip() for x in comparison_df[comparison_kks_col].dropna())
    
    print(f"Unique KKS values in baseline: {len(baseline_kks_unique)}")
    print(f"Unique KKS values in comparison: {len(comparison_kks_unique)}")
    
    print(f"\nSample baseline KKS values:")
    for kks in list(baseline_kks_unique)[:10]:
        print(f"  '{kks}'")
    
    print(f"\nSample comparison KKS values:")
    for kks in list(comparison_kks_unique)[:10]:
        print(f"  '{kks}'")
    
    # Check for direct KKS overlaps
    kks_overlap = baseline_kks_unique.intersection(comparison_kks_unique)
    print(f"\nDirect KKS overlaps: {len(kks_overlap)}")
    if kks_overlap:
        print("Sample overlapping KKS values:")
        for kks in list(kks_overlap)[:5]:
            print(f"  '{kks}'")
    
    # Build KKS lookup and test partial matching
    print(f"\n=== KKS PARTIAL MATCHING ANALYSIS ===")
    kks_lookup = create_kks_lookup_dict(comparison_df, comparison_kks_col)
    
    # Test first 20 baseline KKS values for partial matches
    baseline_sample = baseline_df.head(20)
    kks_match_stats = {'has_matches': 0, 'no_matches': 0}
    
    print("Testing KKS partial matching on first 20 rows:")
    for idx, row in baseline_sample.iterrows():
        baseline_kks = row[baseline_kks_col]
        matched_indices = find_kks_matches_optimized(baseline_kks, kks_lookup)
        
        if matched_indices:
            kks_match_stats['has_matches'] += 1
            print(f"  Row {idx}: '{baseline_kks}' -> {len(matched_indices)} matches")
        else:
            kks_match_stats['no_matches'] += 1
            print(f"  Row {idx}: '{baseline_kks}' -> NO MATCHES")
    
    print(f"\nKKS matching stats (first 20): {kks_match_stats}")
    
    # Analyze article numbers
    print(f"\n=== ARTICLE NUMBER ANALYSIS ===")
    baseline_articles = set(normalize_article_number(x) for x in baseline_df['ARTICLE_NUMBER'].dropna())
    comparison_articles = set(normalize_article_number(x) for x in comparison_df['ARTICLE_NUMBER'].dropna())
    
    # Remove NaN values
    baseline_articles = {x for x in baseline_articles if not pd.isna(x)}
    comparison_articles = {x for x in comparison_articles if not pd.isna(x)}
    
    print(f"Unique normalized article numbers in baseline: {len(baseline_articles)}")
    print(f"Unique normalized article numbers in comparison: {len(comparison_articles)}")
    
    article_overlap = baseline_articles.intersection(comparison_articles)
    print(f"Article number overlaps: {len(article_overlap)}")
    
    print(f"\nSample baseline articles: {list(baseline_articles)[:10]}")
    print(f"Sample comparison articles: {list(comparison_articles)[:10]}")
    
    if article_overlap:
        print(f"Sample overlapping articles: {list(article_overlap)[:10]}")
    
    # Deep dive into a specific case that should match
    print(f"\n=== DETAILED CASE ANALYSIS ===")
    
    # Find a case where KKS matches exist
    for idx, row in baseline_df.head(10).iterrows():
        baseline_kks = row[baseline_kks_col]
        matched_indices = find_kks_matches_optimized(baseline_kks, kks_lookup)
        
        if matched_indices:
            print(f"\nAnalyzing row {idx} with KKS '{baseline_kks}':")
            print(f"Baseline row: Article={row['ARTICLE_NUMBER']}, QTY={row['QTY']}, Cut_Length={row['CUT_LENGTH']}")
            
            print(f"Found {len(matched_indices)} KKS matches:")
            for match_idx in matched_indices[:3]:  # Show first 3 matches
                match_row = comparison_df.loc[match_idx]
                print(f"  Match {match_idx}: Article={match_row['ARTICLE_NUMBER']}, QTY={match_row['QTY']}, Cut_Length={match_row['CUT_LENGTH']}, KKS={match_row['KKS/SU']}")
            
            # Check if any have the same article number
            baseline_article_norm = normalize_article_number(row['ARTICLE_NUMBER'])
            for match_idx in matched_indices:
                match_row = comparison_df.loc[match_idx]
                match_article_norm = normalize_article_number(match_row['ARTICLE_NUMBER'])
                if baseline_article_norm == match_article_norm:
                    print(f"  *** ARTICLE MATCH FOUND: {baseline_article_norm} ***")
                    print(f"      Baseline QTY: {row['QTY']}, Match QTY: {match_row['QTY']}")
                    print(f"      Baseline Cut Length: {row['CUT_LENGTH']}, Match Cut Length: {match_row['CUT_LENGTH']}")
            break
    else:
        print("No KKS matches found in first 10 rows!")

if __name__ == "__main__":
    analyze_matching_issues()
