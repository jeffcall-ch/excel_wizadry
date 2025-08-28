import pandas as pd
import numpy as np
import re
from typing import List, Dict, Any

def load_and_process_bom(file_path: str, output_path: str = None) -> pd.DataFrame:
    """
    Load BOM CSV and add data quality flags
    """
    # Load the CSV
    df = pd.read_csv(file_path)
    
    # Clean and normalize numeric fields
    df = clean_numeric_fields(df)
    
    # Initialize the flag column
    df['DATA_QUALITY_FLAGS'] = ''
    
    # Apply all validation checks
    df = flag_inconsistent_data(df)
    
    # Save processed file if output path provided
    if output_path:
        df.to_csv(output_path, index=False)
        print(f"Processed file saved to: {output_path}")
    
    # Print summary
    print_quality_summary(df)
    
    return df

def clean_numeric_fields(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean and normalize numeric fields by removing units and formatting to 2 decimals
    """
    # Fields that should be cleaned of 'kg' unit
    weight_fields = ['WEIGHT', 'TOTAL_WEIGHT']
    
    # Fields that should be cleaned of 'mm' unit  
    length_fields = ['CUT_LENGTH']
    
    # Clean weight fields
    for field in weight_fields:
        if field in df.columns:
            df[field] = df[field].apply(clean_weight_field)
    
    # Clean length fields
    for field in length_fields:
        if field in df.columns:
            df[field] = df[field].apply(clean_length_field)
    
    # Clean quantity field (should be whole numbers but stored as float with 2 decimals)
    if 'QTY' in df.columns:
        df['QTY'] = df['QTY'].apply(clean_quantity_field)
    
    return df

def clean_weight_field(value) -> float:
    """
    Clean weight field: remove 'kg', convert to float with 2 decimals
    """
    if pd.isna(value) or value == '' or str(value).strip() == '':
        return np.nan
    
    try:
        # Convert to string and remove 'kg' (case insensitive)
        clean_val = str(value).strip()
        clean_val = re.sub(r'kg\b', '', clean_val, flags=re.IGNORECASE).strip()
        
        # Convert to float and round to 2 decimals
        if clean_val == '' or clean_val == 'nan':
            return np.nan
        
        return round(float(clean_val), 2)
    
    except (ValueError, TypeError):
        return np.nan

def clean_length_field(value) -> float:
    """
    Clean length field: remove 'mm', convert to float with 2 decimals
    """
    if pd.isna(value) or value == '' or str(value).strip() == '':
        return np.nan
    
    try:
        # Convert to string and remove 'mm' (case insensitive)
        clean_val = str(value).strip()
        clean_val = re.sub(r'mm\b', '', clean_val, flags=re.IGNORECASE).strip()
        
        # Convert to float and round to 2 decimals
        if clean_val == '' or clean_val == 'nan':
            return np.nan
            
        return round(float(clean_val), 2)
    
    except (ValueError, TypeError):
        return np.nan

def clean_quantity_field(value) -> float:
    """
    Clean quantity field: ensure it's a proper number with 2 decimal format
    """
    if pd.isna(value) or value == '' or str(value).strip() == '':
        return np.nan
    
    try:
        # Convert to float and round to 2 decimals
        return round(float(value), 2)
    
    except (ValueError, TypeError):
        return np.nan

def flag_inconsistent_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply multiple data quality checks and flag inconsistent rows
    """
    flags = []
    
    for idx, row in df.iterrows():
        row_flags = []
        
        # Check 1: Subtotal/Total rows mixed with component data
        if check_subtotal_total_rows(row):
            row_flags.append("SUBTOTAL_TOTAL_ROW")
        
        # Check 2: Missing position numbers for component rows
        if check_missing_position(row):
            row_flags.append("MISSING_POSITION")
        
        # Check 3: Inconsistent article number patterns
        if check_article_number_pattern(row):
            row_flags.append("INVALID_ARTICLE_NUMBER")
        
        # Check 4: Weight calculation inconsistencies
        if check_weight_calculations(row):
            row_flags.append("WEIGHT_CALC_ERROR")
        
        # Check 5: Unreasonable cut lengths
        if check_cut_length_validity(row):
            row_flags.append("INVALID_CUT_LENGTH")
        
        # Check 6: Description patterns
        if check_description_patterns(row):
            row_flags.append("SUSPICIOUS_DESCRIPTION")
        
        # Check 7: Quantity validation
        if check_quantity_validity(row):
            row_flags.append("INVALID_QUANTITY")
        
        # Check 8: KKS code consistency
        if check_kks_consistency(row):
            row_flags.append("KKS_INCONSISTENT")
        
        flags.append("; ".join(row_flags))
    
    df['DATA_QUALITY_FLAGS'] = flags
    return df

def check_subtotal_total_rows(row) -> bool:
    """Check if row is a subtotal/total that shouldn't have component data"""
    description = str(row.get('DESCRIPTION', '')).upper()
    
    if 'SUBTOTAL' in description or 'TOTAL' in description:
        # These rows should not have POS, ARTICLE_NUMBER, or CUT_LENGTH
        has_pos = pd.notna(row.get('POS')) and str(row.get('POS')).strip() != ''
        has_article = pd.notna(row.get('ARTICLE_NUMBER'))
        has_cut_length = pd.notna(row.get('CUT_LENGTH')) and str(row.get('CUT_LENGTH')).strip() != ''
        
        return has_pos or has_article or has_cut_length
    
    return False

def check_missing_position(row) -> bool:
    """Check for missing position numbers in component rows"""
    description = str(row.get('DESCRIPTION', '')).upper()
    
    # Skip subtotal/total rows
    if 'SUBTOTAL' in description or 'TOTAL' in description:
        return False
    
    # If we have an article number, we should have a position
    if pd.notna(row.get('ARTICLE_NUMBER')):
        pos = str(row.get('POS', '')).strip()
        return pos == '' or pos == 'nan'
    
    return False

def check_article_number_pattern(row) -> bool:
    """Check for invalid article number patterns"""
    article_num = str(row.get('ARTICLE_NUMBER', ''))
    
    if article_num == 'nan' or article_num == '':
        return False
    
    # Article numbers should be numeric (6 digits typically)
    if not re.match(r'^\d{6}$', article_num):
        return True
    
    return False

def check_weight_calculations(row) -> bool:
    """Check for weight calculation inconsistencies"""
    try:
        weight = row.get('WEIGHT')
        total_weight = row.get('TOTAL_WEIGHT')
        qty = row.get('QTY')
        
        # Skip if any values are missing or NaN
        if pd.isna(weight) or pd.isna(total_weight) or pd.isna(qty):
            return False
        
        # Values are already cleaned and numeric
        weight_val = float(weight)
        total_weight_val = float(total_weight)
        qty_val = float(qty)
        
        # Check if total_weight = weight * qty (with small tolerance)
        expected_total = round(weight_val * qty_val, 2)
        tolerance = 0.01
        
        return abs(total_weight_val - expected_total) > tolerance
        
    except (ValueError, TypeError):
        return True  # Flag if we can't parse the weights

def check_cut_length_validity(row) -> bool:
    """Check for unreasonable cut lengths"""
    cut_length = row.get('CUT_LENGTH')
    
    # Skip if missing or NaN
    if pd.isna(cut_length):
        return False
    
    try:
        # Value is already cleaned and numeric
        length_val = float(cut_length)
        
        # Flag extremely short or long lengths (in mm)
        if length_val < 10.00 or length_val > 5000.00:
            return True
                
    except (ValueError, TypeError):
        return True
    
    return False

def check_description_patterns(row) -> bool:
    """Check for suspicious description patterns"""
    description = str(row.get('DESCRIPTION', ''))
    
    # Flag descriptions that are too short or have unusual patterns
    if len(description.strip()) < 5 and 'SUBTOTAL' not in description and 'TOTAL' not in description:
        return True
    
    # Flag descriptions with unusual characters or patterns
    if re.search(r'[^\w\s\-/,."()]', description):
        return True
    
    return False

def check_quantity_validity(row) -> bool:
    """Check for invalid quantities"""
    qty = row.get('QTY')
    
    if pd.isna(qty):
        return False
    
    try:
        qty_val = float(qty)
        # Flag negative quantities or extremely high quantities
        return qty_val <= 0 or qty_val > 100
    except (ValueError, TypeError):
        return True

def check_kks_consistency(row) -> bool:
    """Check KKS code consistency patterns"""
    kks = str(row.get('KKS', ''))
    kks_su = str(row.get('KKS/SU', ''))
    
    if kks == 'nan' or kks_su == 'nan':
        return False
    
    # Basic pattern checking - KKS codes should follow certain formats
    # This is a simplified check - adjust based on your specific KKS standards
    try:
        # Remove brackets and quotes for analysis
        kks_clean = re.sub(r'[\[\]\'"]', '', kks)
        kks_su_clean = re.sub(r'[\[\]\'"]', '', kks_su)
        
        # Check if KKS/SU contains /SU suffix pattern
        if '/SU' not in kks_su_clean and kks_su_clean != '':
            return True
            
    except:
        return True
    
    return False

def print_quality_summary(df: pd.DataFrame):
    """Print summary of data quality issues"""
    print("\n" + "="*60)
    print("DATA QUALITY SUMMARY")
    print("="*60)
    
    total_rows = len(df)
    flagged_rows = len(df[df['DATA_QUALITY_FLAGS'] != ''])
    clean_rows = total_rows - flagged_rows
    
    print(f"Total rows: {total_rows}")
    print(f"Clean rows: {clean_rows} ({clean_rows/total_rows*100:.1f}%)")
    print(f"Flagged rows: {flagged_rows} ({flagged_rows/total_rows*100:.1f}%)")
    
    # Count each type of flag
    all_flags = df['DATA_QUALITY_FLAGS'].str.split('; ').explode()
    flag_counts = all_flags.value_counts()
    
    print("\nISSUE BREAKDOWN:")
    print("-" * 40)
    for flag, count in flag_counts.items():
        if flag != '':
            print(f"{flag}: {count} occurrences")
    
    # Show sample of flagged rows
    flagged_df = df[df['DATA_QUALITY_FLAGS'] != '']
    if len(flagged_df) > 0:
        print(f"\nSAMPLE FLAGGED ROWS (first 5):")
        print("-" * 40)
        cols_to_show = ['filename', 'POS', 'ARTICLE_NUMBER', 'DESCRIPTION', 'DATA_QUALITY_FLAGS']
        available_cols = [col for col in cols_to_show if col in flagged_df.columns]
        print(flagged_df[available_cols].head().to_string(index=False))

# Usage example
if __name__ == "__main__":
    # Process the BOM file
    input_file = "extracted_tables_20250827_145355_EXAMPE - Copy.txt"  # Your input file
    output_file = "processed_bom_with_flags.csv"  # Output file
    
    try:
        df_processed = load_and_process_bom(input_file, output_file)
        
        # Optionally, show all flagged rows
        flagged_rows = df_processed[df_processed['DATA_QUALITY_FLAGS'] != '']
        if len(flagged_rows) > 0:
            print(f"\nAll {len(flagged_rows)} flagged rows saved to output file.")
            
    except FileNotFoundError:
        print(f"File '{input_file}' not found. Please check the file path.")
    except Exception as e:
        print(f"Error processing file: {str(e)}")