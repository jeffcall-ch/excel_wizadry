"""
Convert E3D pipe specification CSV file to XLSX format.

This script processes a semicolon-separated CSV file where:
- First column is "SPRE" (pipe specification reference)
- Remaining columns are in alternating name;value pairs
- Output is an Excel (.xlsx) file with proper column structure
"""

import pandas as pd
import sys
from pathlib import Path


def parse_e3d_csv(file_path):
    """
    Parse E3D CSV file with alternating column name/value pairs.
    
    Args:
        file_path: Path to the input CSV file
        
    Returns:
        pandas.DataFrame with parsed data
    """
    data_rows = []
    
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
                
            # Split by semicolon
            parts = [p.strip() for p in line.split(';')]
            
            # First part is SPRE
            spre = parts[0]
            
            # Create row dictionary starting with SPRE
            row = {'SPRE': spre}
            
            # Process remaining parts as alternating column_name, column_value
            # Start from index 1 (after SPRE)
            for i in range(1, len(parts) - 1, 2):
                if i + 1 < len(parts):
                    col_name = parts[i].strip()
                    col_value = parts[i + 1].strip()
                    
                    # Only add if column name is not empty
                    if col_name:
                        row[col_name] = col_value
            
            data_rows.append(row)
    
    # Create DataFrame
    df = pd.DataFrame(data_rows)
    
    # Clean P1 CONN and P2 CONN columns by removing commas
    if 'P1 CONN' in df.columns:
        df['P1 CONN'] = df['P1 CONN'].str.replace(',', '', regex=False)
    if 'P2 CONN' in df.columns:
        df['P2 CONN'] = df['P2 CONN'].str.replace(',', '', regex=False)
    
    return df


def main():
    """Main function to convert CSV to XLSX."""
    
    # Hard-coded paths for easy execution
    if len(sys.argv) < 2:
        # Use hard-coded default paths
        input_file = Path(__file__).parent / "TBY_all_pspecs_wure_macro_08.12.2025.csv"
        output_file = input_file.with_suffix('.xlsx')
    else:
        # Get input file path from command line argument
        input_file = Path(sys.argv[1])
        
        # Determine output file path
        if len(sys.argv) >= 3:
            output_file = Path(sys.argv[2])
        else:
            # Use same name as input but with .xlsx extension
            output_file = input_file.with_suffix('.xlsx')
    
    # Check if file exists
    if not input_file.exists():
        print(f"Error: Input file '{input_file}' not found!")
        sys.exit(1)
    
    print(f"Reading input file: {input_file}")
    
    # Parse the CSV file
    df = parse_e3d_csv(input_file)
    
    print(f"Parsed {len(df)} rows with {len(df.columns)} columns")
    print(f"Columns: {', '.join(df.columns.tolist())}")
    
    # Write to Excel
    print(f"\nWriting to Excel file: {output_file}")
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    print(f"\n✓ Successfully converted to {output_file}")
    print(f"  Total rows: {len(df)}")
    print(f"  Total columns: {len(df.columns)}")


if __name__ == "__main__":
    main()
