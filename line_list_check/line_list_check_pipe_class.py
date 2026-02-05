import os
import re
import pandas as pd


def setup_paths():
    """Setup file and directory paths."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define directories
    dirs = {
        'input_dir': os.path.join(current_dir, "input_file"),
        'pipe_class_dir': os.path.join(current_dir, "pipe_class_summary_file"),
        'output_dir': os.path.join(current_dir, "output")
    }    # Define file paths
    files = {
        'line_list_file': os.path.join(dirs['input_dir'], "MED_pipe_list_02.02.2026.xlsx"),
        'pipe_class_file': os.path.join(dirs['pipe_class_dir'], "PIPE CLASS SUMMARY_LS_06.06.2025_updated_column_names.xlsx"),
        'output_file': os.path.join(dirs['output_dir'], "Line_List_with_Matches.xlsx"),
        'summary_file': os.path.join(dirs['output_dir'], "Pipe_Class_Summary.xlsx"),
        'process_section_file': os.path.join(dirs['output_dir'], "Process_Section_Summary.xlsx")
    }
    return dirs, files


def read_line_list(file_path):
    """Read the line list Excel file and prepare the dataframe."""
    print(f"Reading Line List from '{file_path}'...")
    
    # Read the Excel file with default headers
    query_df = pd.read_excel(file_path, sheet_name='Query')
    
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
    
    # Add check columns for the specified columns
    add_check_columns(query_df)
    
    print(f"Successfully read 'Query' tab - Shape: {query_df.shape}")
    return query_df


def add_check_columns(df):
    """Add check columns after specified columns."""
    columns_to_add_check = [
        'Medium', 'PS [bar(g)]', 'TS [°C]', 'DN', 'PN', 'EN No. Material', 'Pipe Class'
    ]
    
    # Get the list of existing columns
    existing_columns = df.columns.tolist()
    
    # Add check columns after each specified column
    for column in columns_to_add_check:
        if column in existing_columns:
            col_index = existing_columns.index(column)
            df.insert(col_index + 1, f"{column}_check", "")
            existing_columns.insert(col_index + 1, f"{column}_check")
    
    return df


def read_pipe_class_summary(file_path):
    """Read the pipe class summary Excel file and create a dictionary of pipe class data."""
    print(f"Reading Pipe Class Summary from '{file_path}'...")
    
    # Read the Excel file
    pipe_class_df = pd.read_excel(file_path, sheet_name='Pipe Class Summary')
    
    # Create dictionary structure for pipe class data
    pipe_class_dict = {}
    for index, row in pipe_class_df.iterrows():
        pipe_class = row['Pipe Class']
        
        # Skip if pipe class is missing or NaN
        if pd.isna(pipe_class) or pipe_class == '':
            continue
        
        # Store all column values as a dictionary for this pipe class
        pipe_class_dict[pipe_class] = {
            column: row[column] for column in pipe_class_df.columns
        }
    
    print(f"Successfully read Pipe Class Summary - {len(pipe_class_dict)} pipe classes found")
    return pipe_class_dict


def extract_numeric_part(value):
    """Extract numeric part from strings like 'PN 16' or 'DN 25'."""
    if pd.isna(value):
        return "nan"
    
    value_str = str(value)
    match = re.search(r'\d+', value_str)
    if match:
        return int(match.group())
    return "nan"


def compare_medium(row_value, pipe_class_value):
    """Compare medium values."""
    if pd.isna(row_value) or pd.isna(pipe_class_value):
        return 'nan'
    
    # Split the pipe class medium by comma and trim whitespace
    pipe_mediums = [m.strip() for m in str(pipe_class_value).split(',')]
    
    if str(row_value).strip() in pipe_mediums:
        return 'OK'
    else:
        return 'Medium is missing from the pipe class'


def compare_pressure(ps_value, pn_value):
    """Compare pressure values."""
    if pd.isna(ps_value) or pd.isna(pn_value):
        return 'nan'
    
    try:
        ps_numeric = float(ps_value)
        pn_numeric = float(pn_value)
        
        return 'OK' if ps_numeric <= pn_numeric else 'NOK'
    except (ValueError, TypeError):
        return 'nan'


def compare_temperature(ts_value, min_temp, max_temp):
    """Compare temperature values."""
    if pd.isna(ts_value) or pd.isna(min_temp) or pd.isna(max_temp):
        return 'nan'
    
    try:
        ts_numeric = float(ts_value)
        min_temp_numeric = float(min_temp)
        max_temp_numeric = float(max_temp)
        
        return 'OK' if min_temp_numeric <= ts_numeric <= max_temp_numeric else 'NOK'
    except (ValueError, TypeError):
        return 'nan'


def compare_diameter(dn_value, min_dn, max_dn):
    """Compare diameter values."""
    if pd.isna(dn_value) or pd.isna(min_dn) or pd.isna(max_dn):
        return 'nan'
    
    try:
        dn_numeric = extract_numeric_part(dn_value)
        if dn_numeric == 'nan':
            return 'nan'
        
        dn_numeric = float(dn_numeric)
        min_dn_numeric = float(min_dn)
        max_dn_numeric = float(max_dn)
        
        return 'OK' if min_dn_numeric <= dn_numeric <= max_dn_numeric else 'NOK'
    except (ValueError, TypeError):
        return 'nan'


def compare_pn(pn_value, summary_pn):
    """Compare PN values."""
    if pd.isna(pn_value) or pd.isna(summary_pn):
        return 'nan'
    
    try:
        pn_numeric = extract_numeric_part(pn_value)
        summary_pn_numeric = float(summary_pn)
        
        return 'OK' if pn_numeric == summary_pn_numeric else 'NOK'
    except (ValueError, TypeError):
        return 'nan'


def compare_material(material_value, summary_material):
    """Compare material values."""
    if pd.isna(material_value) or pd.isna(summary_material):
        return 'nan'
    
    return 'OK' if str(material_value).strip() == str(summary_material).strip() else 'NOK'


def validate_pipe_data(query_df, pipe_class_dict):
    """Validate pipe data against pipe class specifications."""
    print("Performing validation checks...")
    
    for index, row in query_df.iterrows():
        pipe_class = row.get('Pipe Class')
        
        # Handle missing pipe class
        if pd.isna(pipe_class) or pipe_class == '':
            query_df.at[index, 'Pipe Class_check'] = 'No pipe class assigned'
            continue
        
        # Check if pipe class exists in reference data
        if pipe_class not in pipe_class_dict:
            query_df.at[index, 'Pipe Class_check'] = 'Pipe class not found in summary'
            continue
        
        # Pipe class exists, perform detailed validation
        query_df.at[index, 'Pipe Class_check'] = 'OK'
        pipe_data = pipe_class_dict[pipe_class]
        
        # Validate Medium
        query_df.at[index, 'Medium_check'] = compare_medium(
            row.get('Medium'), pipe_data.get('Medium'))
        
        # Validate Pressure
        query_df.at[index, 'PS [bar(g)]_check'] = compare_pressure(
            row.get('PS [bar(g)]'), pipe_data.get('PN'))
        
        # Validate Temperature
        query_df.at[index, 'TS [°C]_check'] = compare_temperature(
            row.get('TS [°C]'), 
            pipe_data.get('Min temperature (°C)'), 
            pipe_data.get('Max temperature (°C)'))
        
        # Validate Diameter
        query_df.at[index, 'DN_check'] = compare_diameter(
            row.get('DN'),
            pipe_data.get('Diameter from [DN, NPS]'),
            pipe_data.get('Diameter to [DN, NPS]'))
        
        # Validate PN
        query_df.at[index, 'PN_check'] = compare_pn(
            row.get('PN'), pipe_data.get('PN'))
        
        # Validate Material
        query_df.at[index, 'EN No. Material_check'] = compare_material(
            row.get('EN No. Material'), pipe_data.get('EN No. Material'))
    
    # Add summary status column
    add_status_summary(query_df)
    
    print("Validation checks completed")
    return query_df


def add_status_summary(query_df):
    """Add a summary status column consolidating all check results."""
    # Find the position of 'Pipe Class_check' column
    pipe_class_check_idx = query_df.columns.get_loc('Pipe Class_check')
    
    # Initialize the new column with empty strings
    query_df.insert(pipe_class_check_idx + 1, 'Pipe Class status check', "")
    
    # Get all check columns
    check_columns = [col for col in query_df.columns if col.endswith('_check')]
    
    # For each row, populate the status check column
    for idx, row in query_df.iterrows():
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


def create_excel_formats(workbook):
    """Create Excel cell formats for different conditions."""
    formats = {
        # Basic formats
        'center_align': workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        }),
        
        # Regular cell formats
        'ok': workbook.add_format({
            'bg_color': '#C6EFCE', 'font_color': '#006100', 
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        }),
        'nok': workbook.add_format({
            'bg_color': '#FFC7CE', 'font_color': '#9C0006', 
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        }),
        'nan': workbook.add_format({
            'bg_color': '#FFEB9C', 'font_color': '#9C6500', 
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        }),
          # Header formats
        'header_red': workbook.add_format({
            'bg_color': '#FFC7CE', 'font_color': '#9C0006', 
            'align': 'center', 'valign': 'vcenter', 'bold': True, 'text_wrap': True
        }),
        'header_yellow': workbook.add_format({
            'bg_color': '#FFEB9C', 'font_color': '#9C6500', 
            'align': 'center', 'valign': 'vcenter', 'bold': True, 'text_wrap': True
        }),
        
        # Row number formats
        'row_number_red': workbook.add_format({
            'bg_color': '#FFC7CE', 'font_color': '#9C0006', 
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        }),
        'row_number_yellow': workbook.add_format({
            'bg_color': '#FFEB9C', 'font_color': '#9C6500', 
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        }),
        
        # Wrapped text formats for status column
        'ok_wrap': workbook.add_format({
            'bg_color': '#C6EFCE', 'font_color': '#006100', 
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        }),
        'nok_wrap': workbook.add_format({
            'bg_color': '#FFC7CE', 'font_color': '#9C0006', 
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        }),
        'nan_wrap': workbook.add_format({
            'bg_color': '#FFEB9C', 'font_color': '#9C6500', 
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        }),
        
        # Header formats with text wrapping
        'header_red_wrap': workbook.add_format({
            'bg_color': '#FFC7CE', 'font_color': '#9C0006', 
            'align': 'center', 'valign': 'vcenter', 'bold': True, 'text_wrap': True
        }),
        'header_yellow_wrap': workbook.add_format({
            'bg_color': '#FFEB9C', 'font_color': '#9C6500', 
            'align': 'center', 'valign': 'vcenter', 'bold': True, 'text_wrap': True
        }),
        'header_default_wrap': workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'bold': True, 'text_wrap': True
        })
    }
    
    return formats


def classify_cell_value(value):
    """Determine the formatting type for a cell value."""
    # Handle NaN values
    if pd.isna(value):
        return 'nan'
    
    str_val = str(value).strip().lower()
    if str_val == 'ok':
        return 'ok'
    elif str_val == 'nan':
        return 'nan'
    elif str_val:  # Any other non-empty value
        return 'nok'
    return None  # Empty value


def format_worksheet(worksheet, df, formats):
    """Apply formatting to Excel worksheet based on cell values."""
    # Apply center alignment, middle alignment and text wrapping to all cells
    for col in range(len(df.columns)):
        worksheet.set_column(col, col, None, formats['center_align'])
      # Apply center alignment to all data cells not otherwise formatted
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            # Get the cell value and handle NaN values
            cell_value = df.iloc[row_idx, col_idx]
            if pd.isna(cell_value):
                cell_value = ""  # Replace NaN with empty string
            
            # We'll override this for specific cells later
            worksheet.write(row_idx + 1, col_idx, cell_value, formats['center_align'])
    
    # Freeze panes at cell D2 (rows 0-1 and columns 0-3 will remain visible while scrolling)
    worksheet.freeze_panes(1, 3)  # Freeze at row 1 (after headers), column 3 (column D)
    
    # Get check columns
    check_columns = [col for col in df.columns if col.endswith('_check')]
    status_check_column = 'Pipe Class status check'
    
    # Format row numbers based on check values
    for row_idx in range(len(df)):
        row_has_red = False
        row_has_yellow = False
        
        for check_col in check_columns:
            val = str(df.loc[row_idx, check_col])
            format_type = classify_cell_value(val)
            
            if format_type == 'nok':
                row_has_red = True
                break
            elif format_type == 'nan':
                row_has_yellow = True
        
        # Apply formatting to Row Number cell
        if row_has_red:
            worksheet.write(row_idx + 1, 0, df.iloc[row_idx, 0], formats['row_number_red'])
        elif row_has_yellow:
            worksheet.write(row_idx + 1, 0, df.iloc[row_idx, 0], formats['row_number_yellow'])
    
    # Format check columns
    for col_name in check_columns:
        col_idx = df.columns.get_loc(col_name)
        col_has_red = False
        col_has_yellow = False
          # Format cells in this column
        for row_idx in range(len(df)):
            cell_value = df.loc[row_idx, col_name]
            # Handle NaN values
            if pd.isna(cell_value):
                cell_value = ""
                display_value = "nan"
                format_type = 'nan'
            else:
                display_value = str(cell_value)
                format_type = classify_cell_value(cell_value)
            
            if format_type == 'ok':
                worksheet.write(row_idx + 1, col_idx, display_value, formats['ok'])
            elif format_type == 'nan':
                worksheet.write(row_idx + 1, col_idx, display_value, formats['nan'])
                col_has_yellow = True
            elif format_type == 'nok':
                worksheet.write(row_idx + 1, col_idx, display_value, formats['nok'])
                col_has_red = True
        
        # Format column header
        if col_has_red:
            worksheet.write(0, col_idx, col_name, formats['header_red'])
        elif col_has_yellow:
            worksheet.write(0, col_idx, col_name, formats['header_yellow'])
    
    # Format status column specially with wrapped text
    if status_check_column in df.columns:
        format_status_column(worksheet, df, formats, status_check_column)
    
    # Auto-adjust column widths and add filter
    adjust_columns_and_add_filter(worksheet, df)


def format_status_column(worksheet, df, formats, status_col_name):
    """Apply special formatting to the status summary column."""
    status_col_idx = df.columns.get_loc(status_col_name)
    status_col_has_red = False
    status_col_has_yellow = False
      # Set fixed width for the status column - wider for status messages
    worksheet.set_column(status_col_idx, status_col_idx, 35)  # wider for better readability with wrapped text
    
    for row_idx in range(len(df)):
        status_value = df.loc[row_idx, status_col_name]
        
        # Handle NaN values
        if pd.isna(status_value):
            display_value = "nan"
            worksheet.write(row_idx + 1, status_col_idx, display_value, formats['nan_wrap'])
            status_col_has_yellow = True
            continue
        
        # Convert to string for processing
        status_value_str = str(status_value)
        
        if status_value_str == 'OK':
            worksheet.write(row_idx + 1, status_col_idx, status_value_str, formats['ok_wrap'])
        else:
            # Check if the status contains only 'nan' values or has other non-'OK' values
            only_nans = all("'nan'" in part or part == 'nan' for part in status_value_str.split(', ')) if status_value_str else True
            
            if only_nans:
                worksheet.write(row_idx + 1, status_col_idx, status_value_str, formats['nan_wrap'])
                status_col_has_yellow = True
            else:
                worksheet.write(row_idx + 1, status_col_idx, status_value_str, formats['nok_wrap'])
                status_col_has_red = True
    
    # Apply header formatting with text wrapping
    if status_col_has_red:
        worksheet.write(0, status_col_idx, status_col_name, formats['header_red_wrap'])
    elif status_col_has_yellow:
        worksheet.write(0, status_col_idx, status_col_name, formats['header_yellow_wrap'])
    else:
        worksheet.write(0, status_col_idx, status_col_name, formats['header_default_wrap'])


def adjust_columns_and_add_filter(worksheet, df):
    """Auto-adjust column widths and add filter to worksheet."""
    # Auto-adjust column widths based on content
    for col_num, value in enumerate(df.columns.values):
        # Find the maximum length in the column
        max_len = max(
            df[value].astype(str).map(len).max(),
            len(str(value))
        ) + 2  # Add extra space
        
        # For wrapped text, we want reasonable widths - not too narrow, not too wide
        col_width = min(max(max_len / 2, 10), 25)  # Between 10-25 characters
        
        # Set the column width
        worksheet.set_column(col_num, col_num, col_width)
    
    # Set row heights to be taller to accommodate wrapped text
    # Default Excel row height is 15, we'll make it slightly taller
    for row_num in range(len(df) + 1):  # +1 for header row
        worksheet.set_row(row_num, 20)  # 20 points high
    
    # Add filter to the header row
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)


def save_to_excel(df, output_file):
    """Save dataframe to formatted Excel file."""
    print(f"Saving results to '{output_file}'...")
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_file)
    os.makedirs(output_dir, exist_ok=True)
    
    # Handle NaN values by converting them to empty strings to avoid write_number() errors
    df = df.fillna("")
    
    try:
        # Use nan_inf_to_errors option to handle NaN/INF values
        with pd.ExcelWriter(output_file, engine='xlsxwriter', 
                          engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
            # Write dataframe to Excel
            df.to_excel(writer, sheet_name='Results', index=False)
            
            # Get workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Results']
            
            # Create cell formats
            formats = create_excel_formats(workbook)
            
            # Apply formatting to worksheet
            format_worksheet(worksheet, df, formats)
            
        print(f"File saved successfully with formatting applied")
        return True
    except Exception as e:
        print(f"Error saving file: {e}")
        return False


def generate_pipe_class_summary(query_df, pipe_class_dict, output_file):
    """
    Generate a summary of pipe classes from the query dataframe.
    For each pipe class, create separate columns for compliant and non-compliant values.
    Non-compliant columns are only created if there are actually non-compliant values.
    
    Args:
        query_df: DataFrame containing the query data with validation results
        pipe_class_dict: Dictionary containing the reference pipe class data
        output_file: Path to save the summary Excel file
    
    Returns:
        True if successful, False otherwise
    """
    print(f"Generating pipe class summary to '{output_file}'...")
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_file)
    os.makedirs(output_dir, exist_ok=True)
    
    # Base columns that will always be included
    base_columns = [
        'Pipe Class',
        'Medium',
        'PS [bar(g)]',
        'TS [°C]',
        'DN',
        'PN',
        'EN No. Material',
        'Count', 
        'In Reference Data',
        'Compliance Status'
    ]
    
    # Dictionary to store validation information
    validation_info = {}
    
    # Track which columns need non_compliant counterparts
    non_compliant_columns_needed = set()
    
    # Get all unique pipe classes
    pipe_classes = query_df['Pipe Class'].dropna().unique()
    
    # First pass - collect all validation info and determine which non_compliant columns are needed
    for pipe_class in pipe_classes:
        pipe_class_data = query_df[query_df['Pipe Class'] == pipe_class]
        
        # Initialize validation info for this pipe class
        validation_info[pipe_class] = {
            'Medium': {'valid': [], 'invalid': []},
            'PS [bar(g)]': {'valid': [], 'invalid': []},
            'TS [°C]': {'valid': [], 'invalid': []},
            'DN': {'valid': [], 'invalid': []},
            'PN': {'valid': [], 'invalid': []},
            'EN No. Material': {'valid': [], 'invalid': []}
        }
        
        # Group values based on validation results from check columns
        for column in validation_info[pipe_class].keys():
            check_col = f"{column}_check"
            
            # Create mapping of values to their validation results
            validation_map = {}
            for idx, row in pipe_class_data.iterrows():
                value = row[column]
                check_result = row[check_col] if check_col in pipe_class_data.columns else 'nan'
                
                # Skip NaN values
                if pd.isna(value):
                    continue
                
                # Add to validation map
                if value not in validation_map:
                    validation_map[value] = check_result
                    
            # Categorize values based on validation
            for value, check_result in validation_map.items():
                if check_result == 'OK':
                    validation_info[pipe_class][column]['valid'].append(value)
                else:
                    validation_info[pipe_class][column]['invalid'].append(value)
                    non_compliant_columns_needed.add(column)
        
        # Sort the values
        for column in validation_info[pipe_class].keys():
            if column in ['DN', 'PN']:
                key_func = lambda x: extract_numeric_part(x) if extract_numeric_part(x) != 'nan' else float('inf')
                validation_info[pipe_class][column]['valid'].sort(key=key_func)
                validation_info[pipe_class][column]['invalid'].sort(key=key_func)
            else:
                validation_info[pipe_class][column]['valid'].sort()
                validation_info[pipe_class][column]['invalid'].sort()
    
    # Create final column list including non_compliant columns where needed
    all_columns = base_columns.copy()
    for column in non_compliant_columns_needed:
        non_compliant_col = f"{column}_not_compliant"
        # Insert the non_compliant column right after its parent column
        parent_idx = all_columns.index(column)
        all_columns.insert(parent_idx + 1, non_compliant_col)
    
    # Create the summary dataframe with all necessary columns
    summary_df = pd.DataFrame(columns=all_columns)
    
    # Second pass - populate the dataframe with values
    for pipe_class in pipe_classes:
        pipe_class_data = query_df[query_df['Pipe Class'] == pipe_class]
        
        # Count the occurrences
        count = len(pipe_class_data)
        
        # Check if pipe class exists in reference data
        in_reference = 'Yes' if pipe_class in pipe_class_dict else 'No'
        
        # Overall compliance status
        has_invalid_values = any(len(validation_info[pipe_class][col]['invalid']) > 0 for col in validation_info[pipe_class])
        compliance_status = 'Non-Compliant' if has_invalid_values else 'Fully Compliant'
        
        # Create a row with just the basic info
        row_data = {
            'Pipe Class': pipe_class,
            'Count': count,
            'In Reference Data': in_reference,
            'Compliance Status': compliance_status
        }
        
        # Add compliant and non-compliant values to their respective columns
        for column in validation_info[pipe_class].keys():
            valid_values = validation_info[pipe_class][column]['valid']
            invalid_values = validation_info[pipe_class][column]['invalid']
            
            # Add valid values to the main column
            row_data[column] = ', '.join(map(str, valid_values)) if valid_values else ""
            
            # Add invalid values to the not_compliant column if it exists
            non_compliant_col = f"{column}_not_compliant"
            if non_compliant_col in all_columns:
                row_data[non_compliant_col] = ', '.join(map(str, invalid_values)) if invalid_values else ""
        
        # Add to summary dataframe
        summary_df = pd.concat([summary_df, pd.DataFrame([row_data])], ignore_index=True)    # Sort by Pipe Class
    summary_df = summary_df.sort_values('Pipe Class').reset_index(drop=True)
    
    try:
        # Create Excel file
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Write the dataframe to Excel
            summary_df.to_excel(writer, sheet_name='Pipe Class Summary', index=False)
            
            # Get workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Pipe Class Summary']
            
            # Create formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#D9D9D9',  # Light grey background
                'border': 1
            })
            
            # Header format for non-compliant columns
            non_compliant_header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#FFC7CE',  # Light red background
                'font_color': '#9C0006',  # Dark red text
                'border': 1
            })
            
            cell_format = workbook.add_format({
                'text_wrap': True,
                'valign': 'vcenter',
                'border': 1
            })
            
            # Format for compliant values
            compliant_value_format = workbook.add_format({
                'text_wrap': True,
                'valign': 'vcenter',
                'font_color': '#006100',  # Dark green
                'bold': True,
                'border': 1
            })
            
            # Format for non-compliant values
            non_compliant_value_format = workbook.add_format({
                'text_wrap': True,
                'valign': 'vcenter',
                'font_color': '#9C0006',  # Dark red
                'bold': True,
                'border': 1
            })
            
            count_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            
            yes_format = workbook.add_format({
                'bg_color': '#C6EFCE',  # Light green
                'font_color': '#006100',  # Dark green
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            
            no_format = workbook.add_format({
                'bg_color': '#FFC7CE',  # Light red
                'font_color': '#9C0006',  # Dark red
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            
            compliant_status_format = workbook.add_format({
                'bg_color': '#C6EFCE',  # Light green
                'font_color': '#006100',  # Dark green
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bold': True
            })
            
            non_compliant_status_format = workbook.add_format({
                'bg_color': '#FFC7CE',  # Light red
                'font_color': '#9C0006',  # Dark red
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bold': True
            })
            
            # Column name to index mapping for easier reference
            col_indices = {col_name: idx for idx, col_name in enumerate(summary_df.columns)}
            
            # Write headers with appropriate format
            for col_num, col_name in enumerate(summary_df.columns):
                if "_not_compliant" in col_name:
                    worksheet.write(0, col_num, col_name, non_compliant_header_format)
                else:
                    worksheet.write(0, col_num, col_name, header_format)
            
            # Format data cells
            base_columns = ['Medium', 'PS [bar(g)]', 'TS [°C]', 'DN', 'PN', 'EN No. Material']
            
            for row_num in range(len(summary_df)):
                for col_num, col_name in enumerate(summary_df.columns):
                    cell_value = summary_df.iloc[row_num, col_num]
                    
                    # Apply appropriate formatting based on column type
                    if col_name == 'Count':
                        worksheet.write(row_num + 1, col_num, cell_value, count_format)
                    
                    elif col_name == 'In Reference Data':
                        if cell_value == 'Yes':
                            worksheet.write(row_num + 1, col_num, cell_value, yes_format)
                        else:
                            worksheet.write(row_num + 1, col_num, cell_value, no_format)
                    
                    elif col_name == 'Compliance Status':
                        if cell_value == 'Fully Compliant':
                            worksheet.write(row_num + 1, col_num, cell_value, compliant_status_format)
                        else:
                            worksheet.write(row_num + 1, col_num, cell_value, non_compliant_status_format)
                    
                    elif col_name in base_columns:
                        # These are the compliant value columns
                        worksheet.write(row_num + 1, col_num, cell_value, compliant_value_format)
                    
                    elif "_not_compliant" in col_name:
                        # These are the non-compliant value columns
                        worksheet.write(row_num + 1, col_num, cell_value, non_compliant_value_format)
                    
                    else:
                        # Default formatting for any other columns
                        worksheet.write(row_num + 1, col_num, cell_value, cell_format)
            
            # Set column widths based on content
            for col_num, col_name in enumerate(summary_df.columns):
                # Find the maximum length in the column
                max_len = max(
                    summary_df[col_name].astype(str).map(len).max(),
                    len(str(col_name))
                ) + 2  # Add extra space
                
                # Set reasonable column widths based on column type
                if col_name == 'Pipe Class':
                    width = 15
                elif col_name == 'Count':
                    width = 10
                elif col_name == 'In Reference Data':
                    width = 15
                elif col_name == 'Compliance Status':
                    width = 20
                elif "_not_compliant" in col_name:
                    width = 30  # Wider for non-compliant values
                else:
                    width = 25  # For regular compliant value columns
                
                # Adjust width based on content, but within reasonable limits
                width = min(max(width, max_len / 2), 40)
                worksheet.set_column(col_num, col_num, width)
            
            # Set row heights
            for row_num in range(len(summary_df) + 1):
                worksheet.set_row(row_num, 30)
            
            # Add table with filter
            worksheet.add_table(0, 0, len(summary_df), len(summary_df.columns) - 1, {
                'columns': [{'header': col} for col in summary_df.columns],
                'style': 'Table Style Medium 2'
            })
        
        print(f"Pipe class summary saved successfully to '{output_file}'")
        return True
    
    except Exception as e:
        print(f"Error saving pipe class summary: {e}")
        return False


def generate_process_section_summary(query_df, output_file, pipe_class_dict=None):
    """
    Generate a summary of pipe classes and media used in each process section.
    Also generates a summary with pipe classes as the main column showing which process sections use each pipe class.
    Includes pipe class details from the reference data if available.
    
    Args:
        query_df: DataFrame containing the line list data with process section information
        output_file: Path to save the process section summary Excel file
        pipe_class_dict: Dictionary containing the reference pipe class data with detailed specifications
    
    Returns:
        True if successful, False otherwise
    """
    print(f"Generating process section summary to '{output_file}'...")
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_file)
    os.makedirs(output_dir, exist_ok=True)
    
    # Check if the required columns exist
    if "Description Process Section" not in query_df.columns:
        print("Error: 'Description Process Section' column not found in the dataframe")
        return False
    if "Pipe Class" not in query_df.columns:
        print("Error: 'Pipe Class' column not found in the dataframe")
        return False
    if "Medium" not in query_df.columns:
        print("Error: 'Medium' column not found in the dataframe")
        return False
      # -----------------------------
    # Sheet 1: Process Section View
    # -----------------------------
    # Group data by process section
    process_sections = query_df["Description Process Section"].dropna().unique()
    
    # Create a new dataframe to store the summary
    summary_data = []
    
    for section in process_sections:
        # Filter data for this process section
        section_data = query_df[query_df["Description Process Section"] == section]
        
        # Get unique pipe classes and media for this section
        section_pipe_classes = section_data["Pipe Class"].dropna().unique()
        media = section_data["Medium"].dropna().unique()
        
        # Initialize the row data
        row_data = {
            "Process Section": section,
            "Pipe Classes": ", ".join(map(str, sorted(section_pipe_classes))),
            "Number of Pipe Classes": len(section_pipe_classes),
            "Media": ", ".join(map(str, sorted(media))),
            "Number of Media": len(media),
            "Line Count": len(section_data)
        }
        
        # For each pipe class in this section, add its reference data details
        if pipe_class_dict:
            pipe_class_details = []
            for pc in sorted(section_pipe_classes):
                if pc in pipe_class_dict:
                    pipe_spec = pipe_class_dict[pc]
                    detail = f"{pc} [PN: {pipe_spec.get('PN', 'N/A')}, " \
                            f"Temp: {pipe_spec.get('Min temperature (°C)', 'N/A')}-{pipe_spec.get('Max temperature (°C)', 'N/A')}°C, " \
                            f"DN: {pipe_spec.get('Diameter from [DN, NPS]', 'N/A')}-{pipe_spec.get('Diameter to [DN, NPS]', 'N/A')}, " \
                            f"Material: {pipe_spec.get('EN No. Material', 'N/A')}]"
                    pipe_class_details.append(detail)
                else:
                    pipe_class_details.append(f"{pc} [Not in reference]")
            
            row_data["Pipe Class Details"] = "\n".join(pipe_class_details)
        else:
            row_data["Pipe Class Details"] = "Reference data not available"
        
        # Add to summary data
        summary_data.append(row_data)
    
    # Create dataframe from summary data
    summary_df = pd.DataFrame(summary_data)
    
    # Sort by process section name
    summary_df = summary_df.sort_values("Process Section").reset_index(drop=True)
      # -----------------------------
    # Sheet 2: Pipe Class View
    # -----------------------------
    # Group data by pipe class
    pipe_classes = query_df["Pipe Class"].dropna().unique()
    
    # Create a new dataframe to store the pipe class summary
    pipe_class_summary_data = []
    
    for pipe_class in pipe_classes:
        # Filter data for this pipe class
        pipe_class_data = query_df[query_df["Pipe Class"] == pipe_class]
        
        # Get unique process sections and media for this pipe class
        sections = pipe_class_data["Description Process Section"].dropna().unique()
        media = pipe_class_data["Medium"].dropna().unique()
        
        # Initialize the row data
        row_data = {
            "Pipe Class": pipe_class,
            "Process Sections": ", ".join(map(str, sorted(sections))),
            "Number of Process Sections": len(sections),
            "Media": ", ".join(map(str, sorted(media))),
            "Number of Media": len(media),
            "Line Count": len(pipe_class_data)
        }
        
        # Add pipe class details from reference data if available
        if pipe_class_dict and pipe_class in pipe_class_dict:
            pipe_spec = pipe_class_dict[pipe_class]
            
            # Add reference data columns
            row_data.update({
                "Ref: Medium": str(pipe_spec.get('Medium', '')),
                "Ref: PN": str(pipe_spec.get('PN', '')),
                "Ref: Min Temperature (°C)": str(pipe_spec.get('Min temperature (°C)', '')),
                "Ref: Max Temperature (°C)": str(pipe_spec.get('Max temperature (°C)', '')),
                "Ref: DN From": str(pipe_spec.get('Diameter from [DN, NPS]', '')),
                "Ref: DN To": str(pipe_spec.get('Diameter to [DN, NPS]', '')),
                "Ref: Material": str(pipe_spec.get('EN No. Material', ''))
            })
        else:
            # Add empty columns if reference data not available
            row_data.update({
                "Ref: Medium": "Not in reference",
                "Ref: PN": "Not in reference",
                "Ref: Min Temperature (°C)": "Not in reference",
                "Ref: Max Temperature (°C)": "Not in reference",
                "Ref: DN From": "Not in reference",
                "Ref: DN To": "Not in reference",
                "Ref: Material": "Not in reference"
            })
        
        # Add to pipe class summary data
        pipe_class_summary_data.append(row_data)
    
    # Create dataframe from pipe class summary data
    pipe_class_summary_df = pd.DataFrame(pipe_class_summary_data)
      # Sort by pipe class name
    pipe_class_summary_df = pipe_class_summary_df.sort_values("Pipe Class").reset_index(drop=True)
    
    try:
        # Create Excel file
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Get workbook object
            workbook = writer.book
            
            # Create common formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#D9D9D9',  # Light grey background
                'border': 1
            })
            
            cell_format = workbook.add_format({
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'left',
                'border': 1
            })
            
            count_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            
            # -----------------------------
            # Sheet 1: Process Section View
            # -----------------------------
            summary_df.to_excel(writer, sheet_name='Process Section View', index=False)
            worksheet = writer.sheets['Process Section View']
            
            # Write headers with appropriate format
            for col_num, col_name in enumerate(summary_df.columns):
                worksheet.write(0, col_num, col_name, header_format)
            
            # Format data cells
            for row_num in range(len(summary_df)):
                for col_num, col_name in enumerate(summary_df.columns):
                    cell_value = summary_df.iloc[row_num, col_num]
                    
                    # Apply appropriate formatting based on column type
                    if col_name in ["Number of Pipe Classes", "Number of Media", "Line Count"]:
                        worksheet.write(row_num + 1, col_num, cell_value, count_format)
                    else:
                        worksheet.write(row_num + 1, col_num, cell_value, cell_format)
              # Set column widths based on content
            column_widths = {
                "Process Section": 30,
                "Pipe Classes": 40,
                "Number of Pipe Classes": 15,
                "Media": 40,
                "Number of Media": 15,
                "Line Count": 15,
                "Pipe Class Details": 60
            }
            
            for col_num, col_name in enumerate(summary_df.columns):
                width = column_widths.get(col_name, 20)
                worksheet.set_column(col_num, col_num, width)
            
            # Set row heights to accommodate wrapped text
            for row_num in range(len(summary_df) + 1):
                worksheet.set_row(row_num, 30)
            
            # Add table with filter
            worksheet.add_table(0, 0, len(summary_df), len(summary_df.columns) - 1, {
                'columns': [{'header': col} for col in summary_df.columns],
                'style': 'Table Style Medium 2'
            })
            
            # -----------------------------
            # Sheet 2: Pipe Class View
            # -----------------------------
            pipe_class_summary_df.to_excel(writer, sheet_name='Pipe Class View', index=False)
            pipe_class_worksheet = writer.sheets['Pipe Class View']
            
            # Write headers with appropriate format
            for col_num, col_name in enumerate(pipe_class_summary_df.columns):
                pipe_class_worksheet.write(0, col_num, col_name, header_format)
            
            # Format data cells
            for row_num in range(len(pipe_class_summary_df)):
                for col_num, col_name in enumerate(pipe_class_summary_df.columns):
                    cell_value = pipe_class_summary_df.iloc[row_num, col_num]
                    
                    # Apply appropriate formatting based on column type
                    if col_name in ["Number of Process Sections", "Number of Media", "Line Count"]:
                        pipe_class_worksheet.write(row_num + 1, col_num, cell_value, count_format)
                    else:
                        pipe_class_worksheet.write(row_num + 1, col_num, cell_value, cell_format)
              # Set column widths based on content
            pipe_class_column_widths = {
                "Pipe Class": 15,
                "Process Sections": 40,
                "Number of Process Sections": 20,
                "Media": 40,
                "Number of Media": 15,
                "Line Count": 15,
                "Ref: Medium": 20,
                "Ref: PN": 15,
                "Ref: Min Temperature (°C)": 20,
                "Ref: Max Temperature (°C)": 20,
                "Ref: DN From": 15,
                "Ref: DN To": 15,
                "Ref: Material": 20
            }
            
            for col_num, col_name in enumerate(pipe_class_summary_df.columns):
                width = pipe_class_column_widths.get(col_name, 20)
                pipe_class_worksheet.set_column(col_num, col_num, width)
            
            # Set row heights to accommodate wrapped text
            for row_num in range(len(pipe_class_summary_df) + 1):
                pipe_class_worksheet.set_row(row_num, 30)
            
            # Add table with filter
            pipe_class_worksheet.add_table(0, 0, len(pipe_class_summary_df), len(pipe_class_summary_df.columns) - 1, {
                'columns': [{'header': col} for col in pipe_class_summary_df.columns],
                'style': 'Table Style Medium 2'
            })
        
        print(f"Process section and pipe class summary saved successfully to '{output_file}'")
        return True
    
    except Exception as e:
        print(f"Error saving process section summary: {e}")
        return False


def main():
    """Main function to orchestrate the pipe class validation process."""
    try:
        print("Starting Pipe Class Validation...")
        
        # Setup paths
        dirs, files = setup_paths()
        
        # Read input files
        query_df = read_line_list(files['line_list_file'])
        pipe_class_dict = read_pipe_class_summary(files['pipe_class_file'])
          # Validate data
        query_df = validate_pipe_data(query_df, pipe_class_dict)
        
        # Save results
        success = save_to_excel(query_df, files['output_file'])
        
        # Generate and save pipe class summary
        summary_success = generate_pipe_class_summary(query_df, pipe_class_dict, files['summary_file'])        # Generate and save process section summary with pipe class details
        process_section_success = generate_process_section_summary(query_df, files['process_section_file'], pipe_class_dict)
        
        if success and summary_success and process_section_success:
            print("\nProcess completed successfully")
            print(f"Results saved to '{files['output_file']}'")
            print(f"Pipe class summary saved to '{files['summary_file']}'")
            print(f"Process section summary saved to '{files['process_section_file']}'")
        else:
            print("\nProcess completed with errors")
    
    except Exception as e:
        print(f"Error in main process: {e}")


if __name__ == "__main__":
    main()
