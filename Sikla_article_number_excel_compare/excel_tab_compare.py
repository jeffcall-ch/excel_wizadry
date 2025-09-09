import pandas as pd
import os
from datetime import datetime
import numpy as np


def load_excel_tabs(file_path):
    """
    Load both tabs from the Excel file
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        tuple: (baseline_df, comparison_df) - DataFrames for both tabs
    """
    try:
        # Read all sheets to see what's available
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        print(f"Available sheets: {sheet_names}")
        
        # Read the first two sheets
        baseline_df = pd.read_excel(file_path, sheet_name=sheet_names[0])
        comparison_df = pd.read_excel(file_path, sheet_name=sheet_names[1])
        
        print(f"Baseline tab '{sheet_names[0]}' shape: {baseline_df.shape}")
        print(f"Comparison tab '{sheet_names[1]}' shape: {comparison_df.shape}")
        
        return baseline_df, comparison_df, sheet_names
        
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None, None, None


def find_kks_column(df):
    """
    Find the KKS/SU column in the dataframe
    
    Args:
        df (pd.DataFrame): DataFrame to search
        
    Returns:
        str: Column name containing KKS/SU data
    """
    # Look for columns that might contain KKS/SU data
    possible_columns = [col for col in df.columns if 'KKS' in str(col).upper() or 'SU' in str(col).upper()]
    
    if possible_columns:
        return possible_columns[0]
    
    # If not found, print column names to help identify
    print("Available columns:")
    for i, col in enumerate(df.columns):
        print(f"{i}: {col}")
    
    return None


def parse_kks_value(kks_value):
    """
    Parse KKS/SU value which might be a string representation of a list or a plain string
    
    Args:
        kks_value: The KKS/SU value to parse
        
    Returns:
        list: List of KKS values extracted from the input
    """
    if pd.isna(kks_value):
        return []
    
    kks_str = str(kks_value).strip()
    
    # Check if it looks like a string representation of a Python list
    if kks_str.startswith('[') and kks_str.endswith(']'):
        try:
            # Try to evaluate it as a Python literal
            import ast
            kks_list = ast.literal_eval(kks_str)
            if isinstance(kks_list, list):
                return [str(item).strip() for item in kks_list if item]
        except (ValueError, SyntaxError):
            # If parsing fails, treat as single string
            pass
    
    # If it's not a list format, treat as single KKS value
    return [kks_str] if kks_str else []


def create_kks_lookup_dict(comparison_df, comparison_kks_col):
    """
    Create optimized lookup dictionary for KKS/SU partial matching
    Now handles list-format KKS values
    
    Args:
        comparison_df: The comparison DataFrame
        comparison_kks_col: The KKS column name
        
    Returns:
        dict: Dictionary mapping KKS strings to list of row indices
    """
    kks_lookup = {}
    
    for idx, kks_value in comparison_df[comparison_kks_col].items():
        # Parse the KKS value (might be a list format)
        kks_list = parse_kks_value(kks_value)
        
        for kks_str in kks_list:
            if kks_str not in kks_lookup:
                kks_lookup[kks_str] = []
            kks_lookup[kks_str].append(idx)
    
    return kks_lookup


def find_kks_matches_optimized(search_value, kks_lookup):
    """
    Optimized partial match search for KKS/SU codes using pre-built lookup
    Now handles both plain strings and list formats
    
    Args:
        search_value: The KKS/SU value to search for
        kks_lookup: Pre-built lookup dictionary
        
    Returns:
        list: Indices where partial matches are found
    """
    if pd.isna(search_value):
        return []
    
    # Parse the search value (in case it's also in list format)
    search_kks_list = parse_kks_value(search_value)
    
    matches = set()  # Use set to avoid duplicates
    
    for search_kks in search_kks_list:
        # Check each KKS string in lookup for partial matches
        for lookup_kks, indices in kks_lookup.items():
            # Check if search_kks is contained in lookup_kks or vice versa
            if search_kks in lookup_kks or lookup_kks in search_kks:
                matches.update(indices)
    
    return list(matches)


def normalize_article_number(value):
    """
    Normalize article number by stripping whitespace and converting to integer if possible
    
    Args:
        value: The article number value to normalize
        
    Returns:
        int or original value: Normalized article number
    """
    if pd.isna(value):
        return value
    
    # Convert to string and strip whitespace
    str_val = str(value).strip()
    
    if str_val == '' or str_val.lower() == 'nan':
        return np.nan
    
    # Try to convert to integer
    try:
        return int(float(str_val))  # float first to handle cases like "123.0"
    except (ValueError, TypeError):
        return str_val  # Return stripped string if conversion fails


def normalize_cut_length(value):
    """
    Normalize cut length by stripping whitespace and converting to integer when possible
    
    Args:
        value: The cut length value to normalize
        
    Returns:
        int or original value: Normalized cut length
    """
    if pd.isna(value):
        return value
    
    # Handle empty string or zero as NaN (empty)
    if value == '' or value == 0:
        return np.nan
    
    # Convert to string and strip whitespace
    str_val = str(value).strip()
    
    if str_val == '' or str_val.lower() == 'nan':
        return np.nan
    
    # Try to convert to number and round to integer
    try:
        float_val = float(str_val)
        # Round to nearest integer
        return int(round(float_val))
    except (ValueError, TypeError):
        return str_val  # Return stripped string if conversion fails


def is_cut_length_equivalent(val1, val2):
    """
    Check if two cut length values are equivalent (both empty/NaN or equal after normalization)
    
    Args:
        val1, val2: Values to compare
        
    Returns:
        bool: True if values are equivalent
    """
    # Normalize both values
    norm_val1 = normalize_cut_length(val1)
    norm_val2 = normalize_cut_length(val2)
    
    # Check if both are NaN/empty after normalization
    def is_empty_or_nan(val):
        return pd.isna(val)
    
    # If both are empty/NaN, they match
    if is_empty_or_nan(norm_val1) and is_empty_or_nan(norm_val2):
        return True
    
    # If one is empty/NaN and other is not, they don't match
    if is_empty_or_nan(norm_val1) or is_empty_or_nan(norm_val2):
        return False
    
    # Both have values, compare them after normalization
    return norm_val1 == norm_val2


def find_exact_match_by_hierarchy_vectorized(baseline_row, comparison_subset, matched_indices):
    """
    Vectorized version: Find exact match from KKS/SU candidates using hierarchical matching
    
    Args:
        baseline_row: The baseline row to match
        comparison_subset: Pre-filtered comparison DataFrame subset
        matched_indices: List of indices that matched on KKS/SU
        
    Returns:
        tuple: (match_index, match_details, mismatch_category)
    """
    if len(matched_indices) == 0:
        return None, "No KKS/SU match found", "NO_KKS_MATCH"
    
    # Get the baseline values for comparison
    baseline_article_raw = baseline_row.get('ARTICLE_NUMBER')
    baseline_qty = baseline_row.get('QTY')
    baseline_cut_length = baseline_row.get('CUT_LENGTH')
    
    # Normalize baseline article number
    baseline_article = normalize_article_number(baseline_article_raw)
    
    # Filter comparison subset to only matched indices
    candidates = comparison_subset.loc[matched_indices].copy()
    
    if candidates.empty:
        return None, "No valid candidates found", "NO_VALID_CANDIDATES"
    
    # Step 1: Filter by Article Number (with normalization)
    candidates['ARTICLE_NUMBER_NORMALIZED'] = candidates['ARTICLE_NUMBER'].apply(normalize_article_number)
    
    if pd.isna(baseline_article):
        article_matches = candidates[candidates['ARTICLE_NUMBER_NORMALIZED'].isna()]
    else:
        article_matches = candidates[candidates['ARTICLE_NUMBER_NORMALIZED'] == baseline_article]
    
    if article_matches.empty:
        # No article matches
        article_mismatches = len(candidates)
        return None, f"Article number mismatch: {baseline_article_raw} (normalized: {baseline_article}) not found in {article_mismatches} KKS candidates", "ARTICLE_MISMATCH"
    
    # Step 2: Filter by QTY
    if pd.isna(baseline_qty):
        qty_matches = article_matches[article_matches['QTY'].isna()]
    else:
        qty_matches = article_matches[article_matches['QTY'] == baseline_qty]
    
    if qty_matches.empty:
        # Article matches but no QTY matches
        qty_mismatches = len(article_matches)
        return None, f"Article {baseline_article} matches but QTY mismatch: {baseline_qty} not found in {qty_mismatches} candidates", "QTY_MISMATCH"
    
    # Step 3: Filter by Cut Length (with normalization and special empty/NaN handling)
    baseline_cut_length_norm = normalize_cut_length(baseline_cut_length)
    
    final_matches = []
    for idx, row in qty_matches.iterrows():
        comp_cut_length = row.get('CUT_LENGTH')
        if is_cut_length_equivalent(baseline_cut_length, comp_cut_length):
            final_matches.append(idx)
    
    if not final_matches:
        # Article and QTY match but no Cut Length matches
        cut_length_mismatches = len(qty_matches)
        # Show some examples of the mismatched cut lengths for debugging
        mismatch_examples = []
        for idx, row in qty_matches.head(3).iterrows():
            comp_cut_norm = normalize_cut_length(row.get('CUT_LENGTH'))
            mismatch_examples.append(f"comp:{comp_cut_norm}")
        examples_str = ", ".join(mismatch_examples)
        return None, f"Article {baseline_article}, QTY {baseline_qty} match but Cut Length mismatch: baseline:{baseline_cut_length_norm} vs [{examples_str}] in {cut_length_mismatches} candidates", "CUT_LENGTH_MISMATCH"
    
    # We have exact match(es)
    if len(final_matches) == 1:
        match_idx = final_matches[0]
        return match_idx, f"Exact match: Article {baseline_article}, QTY {baseline_qty}, Cut Length {baseline_cut_length_norm}", "EXACT_MATCH"
    else:
        # Multiple exact matches - return the first one but note the issue
        match_idx = final_matches[0]
        return match_idx, f"Multiple exact matches ({len(final_matches)}): Article {baseline_article}, QTY {baseline_qty}, Cut Length {baseline_cut_length_norm}", "MULTIPLE_MATCHES"


def compare_tabs(baseline_df, comparison_df):
    """
    Compare the two tabs according to the specified logic - OPTIMIZED VERSION
    
    Args:
        baseline_df (pd.DataFrame): The baseline tab
        comparison_df (pd.DataFrame): The comparison tab
        
    Returns:
        pd.DataFrame: Baseline DataFrame with added match column
    """
    # Find KKS/SU columns in both DataFrames
    baseline_kks_col = find_kks_column(baseline_df)
    comparison_kks_col = find_kks_column(comparison_df)
    
    if not baseline_kks_col:
        print("Could not find KKS/SU column in baseline tab")
        return None
    
    if not comparison_kks_col:
        print("Could not find KKS/SU column in comparison tab")
        return None
    
    print(f"Using baseline KKS column: '{baseline_kks_col}'")
    print(f"Using comparison KKS column: '{comparison_kks_col}'")
    
    # Create optimized lookup for KKS matching
    print("Building KKS lookup dictionary...")
    kks_lookup = create_kks_lookup_dict(comparison_df, comparison_kks_col)
    print(f"KKS lookup built with {len(kks_lookup)} unique KKS entries")
    
    # Create a copy of the baseline DataFrame
    result_df = baseline_df.copy()
    
    # Prepare columns for vectorized operations
    required_cols = ['ARTICLE_NUMBER', 'QTY', 'CUT_LENGTH']
    
    # Ensure all required columns exist, fill missing with NaN
    for col in required_cols:
        if col not in comparison_df.columns:
            comparison_df[col] = np.nan
        if col not in baseline_df.columns:
            baseline_df[col] = np.nan
    
    # Pre-compute comparison DataFrame subsets for faster lookups
    comparison_subset = comparison_df[required_cols + [comparison_kks_col]].copy()
    
    print(f"Processing {len(baseline_df)} rows with optimized matching...")
    
    # Initialize result lists
    match_results = []
    match_details = []
    mismatch_categories = []
    
    # Process in chunks for better memory management and progress reporting
    chunk_size = 1000
    total_chunks = (len(baseline_df) + chunk_size - 1) // chunk_size
    
    for chunk_idx in range(total_chunks):
        start_idx = chunk_idx * chunk_size
        end_idx = min((chunk_idx + 1) * chunk_size, len(baseline_df))
        
        print(f"Processing chunk {chunk_idx + 1}/{total_chunks} (rows {start_idx}-{end_idx-1})...")
        
        chunk_match_results = []
        chunk_match_details = []
        chunk_mismatch_categories = []
        
        for idx in range(start_idx, end_idx):
            baseline_row = baseline_df.iloc[idx]
            baseline_kks = baseline_row[baseline_kks_col]
            
            # Find KKS matches using optimized lookup
            matched_indices = find_kks_matches_optimized(baseline_kks, kks_lookup)
            
            if not matched_indices:
                chunk_match_results.append('no match')
                chunk_match_details.append('No KKS/SU partial match found')
                chunk_mismatch_categories.append('NO_KKS_MATCH')
            else:
                # Use hierarchical matching: Article Number -> QTY -> Cut Length
                best_match_idx, match_detail, mismatch_category = find_exact_match_by_hierarchy_vectorized(
                    baseline_row, comparison_subset, matched_indices
                )
                
                if best_match_idx is not None:
                    chunk_match_results.append('match')
                    chunk_match_details.append(match_detail)
                    chunk_mismatch_categories.append(mismatch_category)
                else:
                    chunk_match_results.append('no match')
                    chunk_match_details.append(match_detail)
                    chunk_mismatch_categories.append(mismatch_category)
        
        match_results.extend(chunk_match_results)
        match_details.extend(chunk_match_details)
        mismatch_categories.extend(chunk_mismatch_categories)
    
    # Add the results to the DataFrame
    result_df['Match_Status'] = match_results
    result_df['Match_Details'] = match_details
    result_df['Mismatch_Category'] = mismatch_categories
    
    return result_df


def main():
    """
    Main function to execute the comparison process
    """
    # File path
    file_path = r"C:\Users\szil\Repos\excel_wizadry\Sikla_article_number_excel_compare\C4_and_C5_BOM_01.09.2025_to_compare_with_dwg_bom_extract.xlsx"
    
    # Check if file exists
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return
    
    print(f"Loading Excel file: {file_path}")
    
    # Load the Excel tabs
    baseline_df, comparison_df, sheet_names = load_excel_tabs(file_path)
    
    if baseline_df is None or comparison_df is None:
        print("Failed to load Excel file")
        return
    
    # Show some basic info about the data
    print(f"\nBaseline tab columns: {list(baseline_df.columns)}")
    print(f"Comparison tab columns: {list(comparison_df.columns)}")
    
    # Perform the comparison
    print("\nStarting comparison process...")
    result_df = compare_tabs(baseline_df, comparison_df)
    
    if result_df is None:
        print("Comparison failed")
        return
    
    # Generate output filename with timestamp
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_filename = f"comparison_result_{timestamp}.csv"
    output_path = os.path.join(os.path.dirname(file_path), output_filename)
    
    # Save the result
    try:
        result_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"\nResults saved to: {output_path}")
        
        # Print summary statistics
        match_counts = result_df['Match_Status'].value_counts()
        print(f"\nSummary:")
        print(f"Total rows processed: {len(result_df)}")
        for status, count in match_counts.items():
            print(f"{status}: {count} ({count/len(result_df)*100:.1f}%)")
        
        # Print mismatch category breakdown
        print(f"\nMismatch Category Breakdown:")
        mismatch_counts = result_df['Mismatch_Category'].value_counts()
        for category, count in mismatch_counts.items():
            print(f"{category}: {count} ({count/len(result_df)*100:.1f}%)")
            
    except Exception as e:
        print(f"Error saving results: {e}")


if __name__ == "__main__":
    main()
