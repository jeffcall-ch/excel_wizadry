import pandas as pd
import openpyxl

# Path to the Excel file
file_path = r"C:\Users\szil\Repos\excel_wizadry\Mech_szam\mech_szam.xlsm"

# Load the Excel file to see all sheet names
xl_file = pd.ExcelFile(file_path, engine='openpyxl')
print("Sheet names:", xl_file.sheet_names)
print("\n" + "="*80 + "\n")

# Read each sheet
for sheet_name in xl_file.sheet_names:
    print(f"SHEET: {sheet_name}")
    print("="*80)
    
    # Read the sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    
    # Display basic info
    print(f"Shape: {df.shape} (rows, columns)")
    print(f"\nColumn names:\n{df.columns.tolist()}")
    print(f"\nFirst few rows:")
    print(df.head(15))
    
    # For the second sheet, also show rows 30-38 (Mapress data)
    if sheet_name == 'Spans_with_Mapress_from TV046':
        print(f"\nRows 30-38 (Mapress data):")
        print(df.iloc[29:38])
    
    print(f"\nData types:")
    print(df.dtypes)
    
    print(f"\nBasic statistics:")
    print(df.describe())
    
    print("\n" + "="*80 + "\n")
