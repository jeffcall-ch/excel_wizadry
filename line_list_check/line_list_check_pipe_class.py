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
pipe_class_file = os.path.join(pipe_class_dir, "PIPE CLASS SUMMARY_LS_06.06.2025_updated_column_names.xlsx")

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
        
        # Add a sequential number column at the beginning, starting from 1
        query_df.insert(0, "Row Number", range(1, len(query_df) + 1))
    
    # Add new check columns after each specified column
    columns_to_add_check = [
        'Medium',
        'PS [bar(g)]',
        'TS [°C]',
        'DN',
        'PN',
        'EN No. Material',
        'Pipe Class'
    ]
    
    # Get the list of existing columns
    existing_columns = query_df.columns.tolist()
    
    # For each specified column, add a new check column after it
    for column in columns_to_add_check:
        if column in existing_columns:
            # Find the position of the column
            col_index = existing_columns.index(column)
            # Insert the new check column right after it
            query_df.insert(col_index + 1, f"{column}_check", "")
            # Update the existing columns list to reflect the new column
            existing_columns.insert(col_index + 1, f"{column}_check")
    
    print(f"Successfully read 'Query' tab from {line_list_file}")
    print(f"Shape after setting headers and adding row numbers: {query_df.shape}")
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
    
    # Create a dictionary with Pipe Class as key and row data as values
    pipe_class_dict = {}
    
    # Process each row in the pipe class dataframe
    for index, row in pipe_class_df.iterrows():
        # Get the pipe class value which will be the key
        pipe_class = row['Pipe Class']
        
        # Skip if pipe class is missing or NaN
        if pd.isna(pipe_class) or pipe_class == '':
            continue
        
        # Create a list to hold dictionaries for each column
        column_data_list = []
        
        # Create a dictionary for each column in the row
        for column in pipe_class_df.columns:
            column_data_list.append({column: row[column]})
        
        # Add to the main dictionary with Pipe Class as key
        pipe_class_dict[pipe_class] = column_data_list
      # Print the first 2 rows of the dictionary for verification
    print("\nCreated dictionary with pipe class data. First 2 rows:")
    pipe_class_keys = list(pipe_class_dict.keys())[:2]
    
    if len(pipe_class_keys) >= 1:
        # Print first entry
        print(f"Pipe Class {pipe_class_keys[0]}:")
        for item in pipe_class_dict[pipe_class_keys[0]]:
            for key, value in item.items():
                print(f"    {key}: {value}")
        # Also print the full dictionary structure for clarity
        print("\nFirst row dict structure:")
        print(pipe_class_dict[pipe_class_keys[0]])
    
    if len(pipe_class_keys) >= 2:
        print("\n----------------------------------------")  # Clear separator
        
        # Print second entry
        print(f"Pipe Class {pipe_class_keys[1]}:")
        for item in pipe_class_dict[pipe_class_keys[1]]:
            for key, value in item.items():
                print(f"    {key}: {value}")
        # Also print the full dictionary structure for clarity
        print("\nSecond row dict structure:")
        print(pipe_class_dict[pipe_class_keys[1]])
    
    print(f"\nTotal entries in pipe class dictionary: {len(pipe_class_dict)}")
    
except Exception as e:
    print(f"Error reading or processing Pipe Class Summary file: {e}")

# 3. Create a dictionary of row data with specific columns
try:
    # List of columns to include in the dictionary (excluding KKS which will be the key)
    columns_to_include = [
        'Line No.',
        'Medium',
        'Medium_check',
        'PS [bar(g)]',
        'PS [bar(g)]_check',
        'TS [°C]',
        'TS [°C]_check',
        'DN',
        'DN_check',
        'PN',
        'PN_check',
        'OD [mm]',
        'EN No. Material',
        'EN No. Material_check',
        'Pipe Class',
        'Pipe Class_check',
        'Linked Process Section',
        'Description Process Section',
        'Supply Lot'
    ]
    
    # Create the row dictionary structure
    row_data_dict = {}    # Iterate through the dataframe rows
    for index, row in query_df.iterrows():
        row_number = row['Row Number']
        
        # Create list of dictionaries for each column, including KKS
        column_data_list = []
        # Add KKS to the column list
        column_data_list.append({'KKS': row['KKS']})
        
        for column in columns_to_include:
            if column in query_df.columns:
                column_data_list.append({column: row[column]})
        
        # Add to the main dictionary with Row Number as key
        row_data_dict[row_number] = column_data_list
      # Print the first 2 rows of the dictionary for verification
    print("\nCreated dictionary with row data. First 2 rows:")
    row_keys = list(row_data_dict.keys())[:2]
    
    # Print first entry
    print(f"Row {row_keys[0]}:")
    for item in row_data_dict[row_keys[0]]:
        print(f"    {item}")
    
    print()  # Separator
    
    # Print second entry
    print(f"Row {row_keys[1]}:")
    for item in row_data_dict[row_keys[1]]:
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
