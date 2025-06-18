import pandas as pd
import os
import re

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
    
    # Create a cleaner dictionary structure for pipe class data
    # Each key is a pipe class, and the value is a dictionary of attributes
    pipe_class_dict = {}
    
    for index, row in pipe_class_df.iterrows():
        pipe_class = row['Pipe Class']
        
        # Skip if pipe class is missing or NaN
        if pd.isna(pipe_class) or pipe_class == '':
            continue
            
        # Create a dictionary of column:value pairs for this pipe class
        pipe_data = {}
        for column in pipe_class_df.columns:
            pipe_data[column] = row[column]
            
        # Store the dictionary with the pipe class as the key
        pipe_class_dict[pipe_class] = pipe_data
    
    # Print sample of the restructured dictionary
    print("\nRestructured pipe class dictionary (sample):")
    sample_keys = list(pipe_class_dict.keys())[:2]
    
    for key in sample_keys:
        print(f"Pipe Class: {key}")
        print(pipe_class_dict[key])
        print()
        
    print(f"Total entries in pipe class dictionary: {len(pipe_class_dict)}")
    
except Exception as e:
    print(f"Error reading or processing Pipe Class Summary file: {e}")

# 3. Create a dictionary of row data with specific columns
try:
    # List of columns to include in the dictionary
    columns_to_include = [
        'KKS',
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
        row_number = row['Row Number']
        
        # Create list of dictionaries for each column
        column_data_list = []
        for column in columns_to_include:
            if column in query_df.columns:
                column_data_list.append({column: row[column]})
        
        # Add to the main dictionary with Row Number as key
        row_data_dict[row_number] = column_data_list
    
    # Print the first 2 rows of the dictionary for verification
    print("\nCreated dictionary with row data. First 2 entries:")
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

# 4. Process each row in the Line List and perform validation checks
try:
    # Function to extract numeric part from strings like "PN 16" or "PN16"
    def extract_numeric_part(value):
        if pd.isna(value):
            return "nan"
        
        # Convert to string if not already
        value_str = str(value)
        
        # Use regex to extract digits
        match = re.search(r'\d+', value_str)
        if match:
            return int(match.group())
        return "nan"
    
    # Process each row in the query dataframe
    for index, row in query_df.iterrows():        # Get the pipe class value
        pipe_class = row.get('Pipe Class')
        
        # Check if pipe class is empty or missing
        if pd.isna(pipe_class) or pipe_class == '':
            query_df.at[index, 'Pipe Class_check'] = 'No pipe class assigned'
            continue
            
        # Check if pipe class exists in the pipe class dictionary
        if pipe_class not in pipe_class_dict:
            query_df.at[index, 'Pipe Class_check'] = 'Pipe class not found in summary'
            continue
        
        # If pipe class exists, set check to 'OK' and get the pipe class data
        query_df.at[index, 'Pipe Class_check'] = 'OK'
        pipe_data = pipe_class_dict[pipe_class]
        
        # Check Medium
        medium_value = row.get('Medium')
        if not pd.isna(medium_value) and not pd.isna(pipe_data.get('Medium')):
            # Split the pipe class medium by comma and trim whitespace
            pipe_mediums = [m.strip() for m in str(pipe_data.get('Medium')).split(',')]
            
            if str(medium_value).strip() in pipe_mediums:
                query_df.at[index, 'Medium_check'] = 'OK'
            else:
                query_df.at[index, 'Medium_check'] = 'Medium is missing from the pipe class'
        else:
            query_df.at[index, 'Medium_check'] = 'nan'
        
        # Check PS [bar(g)]
        ps_value = row.get('PS [bar(g)]')
        if not pd.isna(ps_value) and not pd.isna(pipe_data.get('PN')):
            try:
                ps_numeric = float(ps_value)
                pn_numeric = float(pipe_data.get('PN'))
                
                if ps_numeric <= pn_numeric:
                    query_df.at[index, 'PS [bar(g)]_check'] = 'OK'
                else:
                    query_df.at[index, 'PS [bar(g)]_check'] = 'NOK'
            except (ValueError, TypeError):
                query_df.at[index, 'PS [bar(g)]_check'] = 'nan'
        else:
            query_df.at[index, 'PS [bar(g)]_check'] = 'nan'
        
        # Check TS [°C]
        ts_value = row.get('TS [°C]')
        if not pd.isna(ts_value) and not pd.isna(pipe_data.get('Min temperature (°C)')) and not pd.isna(pipe_data.get('Max temperature (°C)')):
            try:
                ts_numeric = float(ts_value)
                min_temp = float(pipe_data.get('Min temperature (°C)'))
                max_temp = float(pipe_data.get('Max temperature (°C)'))
                
                if min_temp <= ts_numeric <= max_temp:
                    query_df.at[index, 'TS [°C]_check'] = 'OK'
                else:
                    query_df.at[index, 'TS [°C]_check'] = 'NOK'
            except (ValueError, TypeError):
                query_df.at[index, 'TS [°C]_check'] = 'nan'
        else:
            query_df.at[index, 'TS [°C]_check'] = 'nan'
          # Check DN
        dn_value = row.get('DN')
        if not pd.isna(dn_value) and not pd.isna(pipe_data.get('Diameter from [DN, NPS]')) and not pd.isna(pipe_data.get('Diameter to [DN, NPS]')):
            try:
                # Extract numeric part from DN string like "DN 25"
                dn_numeric = extract_numeric_part(dn_value)
                if dn_numeric != "nan":
                    dn_numeric = float(dn_numeric)
                    min_dn = float(pipe_data.get('Diameter from [DN, NPS]'))
                    max_dn = float(pipe_data.get('Diameter to [DN, NPS]'))
                    
                    if min_dn <= dn_numeric <= max_dn:
                        query_df.at[index, 'DN_check'] = 'OK'
                    else:
                        query_df.at[index, 'DN_check'] = 'NOK'
                else:
                    query_df.at[index, 'DN_check'] = 'nan'
            except (ValueError, TypeError):
                query_df.at[index, 'DN_check'] = 'nan'
        else:
            query_df.at[index, 'DN_check'] = 'nan'
        
        # Check PN
        pn_value = row.get('PN')
        if not pd.isna(pn_value) and not pd.isna(pipe_data.get('PN')):
            try:
                # Extract numeric part from PN string like "PN 16"
                pn_numeric = extract_numeric_part(pn_value)
                summary_pn_numeric = float(pipe_data.get('PN'))
                
                if pn_numeric == summary_pn_numeric:
                    query_df.at[index, 'PN_check'] = 'OK'
                else:
                    query_df.at[index, 'PN_check'] = 'NOK'
            except (ValueError, TypeError):
                query_df.at[index, 'PN_check'] = 'nan'
        else:
            query_df.at[index, 'PN_check'] = 'nan'
          # Check EN No. Material
        material_value = row.get('EN No. Material')
        summary_material = pipe_data.get('EN No. Material')
        
        if not pd.isna(material_value) and not pd.isna(summary_material):
            if str(material_value).strip() == str(summary_material).strip():
                query_df.at[index, 'EN No. Material_check'] = 'OK'
            else:
                query_df.at[index, 'EN No. Material_check'] = 'NOK'
        else:
            query_df.at[index, 'EN No. Material_check'] = 'nan'
    
    print("\nValidation checks completed for all rows")
      # Add a new 'Pipe Class status check' column after 'Pipe Class_check'
    try:
        # Find the position of 'Pipe Class_check' column
        pipe_class_check_idx = query_df.columns.get_loc('Pipe Class_check')
        
        # Initialize the new column with empty strings
        query_df.insert(pipe_class_check_idx + 1, 'Pipe Class status check', "")
        
        # Get all check columns
        check_columns = [col for col in query_df.columns if col.endswith('_check')]
          # For each row, populate the status check column
        for idx, row in query_df.iterrows():
            # Initialize lists for different types of non-OK checks
            nan_checks = []
            non_ok_checks = []
            
            # Check all check columns for this row
            for check_col in check_columns:
                val = str(row[check_col])
                # Only include checks that have actual content (not empty strings)
                if val and val.strip():
                    if val == 'nan':
                        nan_checks.append(f"{check_col}: '{val}'")
                    elif val != 'OK':
                        non_ok_checks.append(f"{check_col}: '{val}'")
            
            # Set the value in the new column
            if not nan_checks and not non_ok_checks:
                query_df.at[idx, 'Pipe Class status check'] = 'OK'
            else:
                # Combine all issues, putting non-OK ones first
                all_issues = non_ok_checks + nan_checks
                query_df.at[idx, 'Pipe Class status check'] = ', '.join(all_issues)
        
        print("Added 'Pipe Class status check' column with summarized status")
    except Exception as e:
        print(f"Error adding status check column: {e}")
    
except Exception as e:
    print(f"Error during validation checks: {e}")

# Optional: Save the updated DataFrame to a new Excel file with formatting
try:
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Define output file path
    output_file = os.path.join(output_dir, "Line_List_with_Matches.xlsx")
    
    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Convert the DataFrame to an XlsxWriter Excel object
        query_df.to_excel(writer, sheet_name='Results', index=False)        # Get the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Results']
      # Create formats for cell highlighting
        center_align = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        ok_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'align': 'center', 'valign': 'vcenter'})  # Green for 'OK'
        nok_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'center', 'valign': 'vcenter'})  # Red for non-OK, non-nan
        nan_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'align': 'center', 'valign': 'vcenter'})  # Yellow for 'nan'
        header_red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'center', 'valign': 'vcenter', 'bold': True})
        header_yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'align': 'center', 'valign': 'vcenter', 'bold': True})
        row_number_red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'align': 'center', 'valign': 'vcenter'})
        row_number_yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500', 'align': 'center', 'valign': 'vcenter'})
        
        # Get check column indices and names
        check_columns = [col_name for col_name in query_df.columns if col_name.endswith('_check')]
        check_column_indices = [query_df.columns.get_loc(col) for col in check_columns]
        
        # Apply center alignment to all cells
        for col in range(len(query_df.columns)):
            worksheet.set_column(col, col, None, center_align)        # Define function to check the value and determine appropriate format
        def get_cell_format(value):
            str_val = str(value).strip().lower()
            if str_val == 'ok':
                return 'ok'
            elif str_val == 'nan':
                return 'nan'
            elif str_val:  # Any other non-empty value
                return 'nok'
            return None  # Empty value
        
        # Get all check columns including the new status check column
        check_columns = [col for col in query_df.columns if col.endswith('_check')]
        status_check_column = 'Pipe Class status check'
        
        # Apply color coding to column A (Row Number) based on the row's check values
        for row_idx in range(len(query_df)):
            row_has_red = False
            row_has_yellow = False
            
            # Check all check columns for this row
            for check_col in check_columns:
                val = str(query_df.loc[row_idx, check_col])
                format_type = get_cell_format(val)
                
                if format_type == 'nok':
                    row_has_red = True
                    break  # Red takes precedence, no need to check further
                elif format_type == 'nan':
                    row_has_yellow = True
            
            # Apply appropriate formatting to Row Number cell
            if row_has_red:
                worksheet.write(row_idx + 1, 0, query_df.iloc[row_idx, 0], row_number_red_format)
            elif row_has_yellow:
                worksheet.write(row_idx + 1, 0, query_df.iloc[row_idx, 0], row_number_yellow_format)
        
        # Apply formatting to check columns
        for col_name in check_columns:
            col_idx = query_df.columns.get_loc(col_name)
            col_letter = chr(65 + col_idx) if col_idx < 26 else chr(64 + col_idx // 26) + chr(65 + col_idx % 26)
            
            # Track column header color status
            col_has_red = False
            col_has_yellow = False
            
            # Apply cell formatting based on values
            for row_idx in range(len(query_df)):
                cell_value = str(query_df.loc[row_idx, col_name])
                format_type = get_cell_format(cell_value)
                
                # Apply cell formatting
                if format_type == 'ok':
                    worksheet.write(row_idx + 1, col_idx, cell_value, ok_format)
                elif format_type == 'nan':
                    worksheet.write(row_idx + 1, col_idx, cell_value, nan_format)
                    col_has_yellow = True
                elif format_type == 'nok':
                    worksheet.write(row_idx + 1, col_idx, cell_value, nok_format)
                    col_has_red = True
            
            # Apply header formatting based on worst-case scenario in column
            if col_has_red:
                worksheet.write(0, col_idx, col_name, header_red_format)
            elif col_has_yellow:
                worksheet.write(0, col_idx, col_name, header_yellow_format)
                  # Create formats with text wrapping for the status column
        ok_format_wrap = workbook.add_format({
            'bg_color': '#C6EFCE', 
            'font_color': '#006100', 
            'align': 'center', 
            'valign': 'vcenter',
            'text_wrap': True
        })
        nok_format_wrap = workbook.add_format({
            'bg_color': '#FFC7CE', 
            'font_color': '#9C0006', 
            'align': 'center', 
            'valign': 'vcenter',
            'text_wrap': True
        })
        nan_format_wrap = workbook.add_format({
            'bg_color': '#FFEB9C', 
            'font_color': '#9C6500', 
            'align': 'center', 
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Apply formatting to the new status summary column
        if status_check_column in query_df.columns:
            status_col_idx = query_df.columns.get_loc(status_check_column)
            status_col_has_red = False
            status_col_has_yellow = False
              # Set fixed width for the status column 
            # Excel column width is measured in characters, where the column width value represents 
            # the number of characters that can be displayed in a cell using the standard font
            # A width of 8.43 characters is approximately 50 pixels with the default font
            worksheet.set_column(status_col_idx, status_col_idx, 8.43)
            
            for row_idx in range(len(query_df)):
                status_value = str(query_df.loc[row_idx, status_check_column])
                
                if status_value == 'OK':
                    worksheet.write(row_idx + 1, status_col_idx, status_value, ok_format_wrap)
                else:
                    # Check if the status contains only 'nan' values or has other non-'OK' values
                    only_nans = True
                    
                    # Parse the status value to check for non-'nan' issues
                    if status_value != '':
                        check_parts = status_value.split(', ')
                        for part in check_parts:
                            if "'nan'" not in part and part != 'nan':
                                only_nans = False
                                break
                    
                    if only_nans:
                        worksheet.write(row_idx + 1, status_col_idx, status_value, nan_format_wrap)
                        status_col_has_yellow = True
                    else:
                        worksheet.write(row_idx + 1, status_col_idx, status_value, nok_format_wrap)
                        status_col_has_red = True
              # Create header formats with text wrapping
            header_red_format_wrap = workbook.add_format({
                'bg_color': '#FFC7CE', 
                'font_color': '#9C0006', 
                'align': 'center', 
                'valign': 'vcenter', 
                'bold': True,
                'text_wrap': True
            })
            header_yellow_format_wrap = workbook.add_format({
                'bg_color': '#FFEB9C', 
                'font_color': '#9C6500', 
                'align': 'center', 
                'valign': 'vcenter', 
                'bold': True,
                'text_wrap': True
            })
            header_default_format_wrap = workbook.add_format({
                'align': 'center', 
                'valign': 'vcenter', 
                'bold': True,
                'text_wrap': True
            })
            
            # Apply header formatting with text wrapping
            if status_col_has_red:
                worksheet.write(0, status_col_idx, status_check_column, header_red_format_wrap)
            elif status_col_has_yellow:
                worksheet.write(0, status_col_idx, status_check_column, header_yellow_format_wrap)
            else:
                worksheet.write(0, status_col_idx, status_check_column, header_default_format_wrap)
        
        # Auto-adjust column widths based on content
        for col_num, value in enumerate(query_df.columns.values):
            # Find the maximum length in the column
            max_len = max(
                query_df[value].astype(str).map(len).max(),
                len(str(value))
            ) + 2  # Add a little extra space
            
            # Set the column width
            worksheet.set_column(col_num, col_num, max_len)
            
        # Add filter to the header row
        num_columns = len(query_df.columns) - 1
        worksheet.autofilter(0, 0, len(query_df), num_columns)
            
    print(f"\nSaved line list to {output_file} with formatting applied.")
    print("- Auto-sized columns for better readability")
    print("- Filter added to header row for easy data filtering")
    print("- Conditional formatting added to check columns (green for OK, red for NOK, yellow for nan)")
except Exception as e:
    print(f"Error saving output file: {e}")
