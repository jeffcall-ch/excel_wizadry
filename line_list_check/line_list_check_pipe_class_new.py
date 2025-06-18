import pandas as pd
import os

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Define input and output directories
input_dir = os.path.join(current_dir, "input_file")
pipe_class_dir = os.path.join(current_dir, "pipe_class_summary_file")
output_dir = os.path.join(current_dir, "output")

# Excel file paths
line_list_file = os.path.join(input_dir, "Export 17.06.2025_LS.xlsx")
pipe_class_file = os.path.join(pipe_class_dir, "PIPE CLASS SUMMARY_LS_06.06.2025.xlsx")

# 1. Read the Line List file - 'Query' tab
try:
    # Read the Excel file with default headers
    query_df = pd.read_excel(line_list_file, sheet_name='Query')
    
    # Check if the first row contains the actual column names
    if query_df.iloc[0]["F1"] == "Line No." and query_df.iloc[0]["F2"] == "KKS":
        # Use the first row values to rename the columns
        column_mapping = {f"F{i+1}": str(query_df.iloc[0][f"F{i+1}"]) for i in range(31)}
        # Handle the special column name
        column_mapping["ComosSystemInfo"] = "ComosSystemInfo"
        query_df = query_df.rename(columns=column_mapping)
        
        # Drop the first row since it's now used as headers
        query_df = query_df.iloc[1:].reset_index(drop=True)
    
    print(f"Successfully read 'Query' tab from {line_list_file}")
    print(f"Shape after setting headers: {query_df.shape}")
    print("\nFirst few rows of Line List:")
    print(query_df.head())
    print("\nLine List column names:")
    print(query_df.columns.tolist())
except Exception as e:
    print(f"Error reading Line List file: {e}")

# 2. Read the Pipe Class Summary file - 'Pipe Class Summary' sheet
try:
    # Read the second Excel file
    pipe_class_df = pd.read_excel(pipe_class_file, sheet_name='Pipe Class Summary')
    
    print(f"\nSuccessfully read 'Pipe Class Summary' tab from {pipe_class_file}")
    print(f"Shape: {pipe_class_df.shape}")
    print("\nFirst few rows of Pipe Class Summary:")
    print(pipe_class_df.head())
    print("\nPipe Class Summary column names:")
    print(pipe_class_df.columns.tolist())
except Exception as e:
    print(f"Error reading Pipe Class Summary file: {e}")

# 3. Create a dictionary of row data with specific columns
try:
    # List of columns to include in the dictionary (excluding KKS which will be the key)
    columns_to_include = [
        'Line No.',
        'Medium',
        'PS [bar(g)]',
        'TS [°C]',
        'DN',
        'PN',
        'OD [mm]',
        'EN No. Material',
        'Pipe Class',
        'Linked Process Section',
        'Description Process Section',
        'Supply Lot'
    ]
    
    # Create the row dictionary structure
    row_data_dict = {}
    
    # Iterate through the dataframe rows
    for index, row in query_df.iterrows():
        kks_value = row['KKS']
        
        # Create list of dictionaries for each column
        column_data_list = []
        for column in columns_to_include:
            if column in query_df.columns:
                column_data_list.append({column: row[column]})
        
        # Add to the main dictionary with KKS as key
        row_data_dict[kks_value] = column_data_list
    
    # Print the first 2 rows of the dictionary for verification
    print("\nCreated dictionary with row data. First 2 KKS entries:")
    kks_keys = list(row_data_dict.keys())[:2]
    
    # Print first entry
    print(f"KKS: {kks_keys[0]}")
    for item in row_data_dict[kks_keys[0]]:
        print(f"    {item}")
    
    print()  # Separator
    
    # Print second entry
    print(f"KKS: {kks_keys[1]}")
    for item in row_data_dict[kks_keys[1]]:
        print(f"    {item}")
    
    print(f"\nTotal entries in dictionary: {len(row_data_dict)}")
    
except Exception as e:
    print(f"Error creating row data dictionary: {e}")

# Optional: Save the updated DataFrame to a new Excel file with formatting
try:
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Define output file path
    output_file = os.path.join(output_dir, "Line_List_with_Matches.xlsx")
    
    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Convert the DataFrame to an XlsxWriter Excel object
        query_df.to_excel(writer, sheet_name='Results', index=False)
        
        # Get the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Results']
        
        # Create a red fill format for cells that might need highlighting in the future
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        
        # Auto-adjust column widths based on content
        for col_num, value in enumerate(query_df.columns.values):
            # Find the maximum length in the column
            max_len = max(
                query_df[value].astype(str).map(len).max(),
                len(str(value))
            ) + 2  # Add a little extra space
            
            # Set the column width
            worksheet.set_column(col_num, col_num, max_len)
            
        # Add filter to the header row (row 0)
        num_columns = len(query_df.columns) - 1
        worksheet.autofilter(0, 0, len(query_df), num_columns)
            
    print(f"\nSaved line list to {output_file} with formatting applied.")
    print("- Auto-sized columns for better readability")
    print("- Filter added to header row for easy data filtering")
except Exception as e:
    print(f"Error saving output file: {e}")
