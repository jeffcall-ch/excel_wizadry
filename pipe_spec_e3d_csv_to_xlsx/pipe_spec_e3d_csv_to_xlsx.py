import os
import pandas as pd

# Define the folder containing the CSV files
folder_path = r"C:\Users\szil\Repos\excel_wizadry\pipe_spec_e3d_csv_to_xlsx"
output_file = os.path.join(folder_path, "combined_output.xlsx")

# Initialize an empty list to store all rows
all_rows = []
max_columns = 0

# Function to split rows by semicolon (no CSV quoting rules)
def process_file(file_path, delimiter=';'):
    rows = []
    with open(file_path, 'r', encoding='utf-8-sig') as file:  # utf-8-sig removes BOM
        for line in file:
            line = line.strip()
            if line:  # Skip empty lines
                # Simple split by delimiter
                row = line.split(delimiter)
                rows.append(row)
    return rows

# Iterate through all files in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith(".csv"):
        file_path = os.path.join(folder_path, file_name)
        try:
            # Process the file
            rows = process_file(file_path)
            all_rows.extend(rows)
            
            # Track the maximum number of columns
            for row in rows:
                if len(row) > max_columns:
                    max_columns = len(row)
            
            print(f"Processed {file_name}: {len(rows)} rows")
        except Exception as e:
            print(f"Error processing {file_name}: {e}")

# Ensure all rows have the same number of columns by padding with empty strings
for row in all_rows:
    while len(row) < max_columns:
        row.append('')

# Create column names
column_names = [f"Column_{i+1}" for i in range(max_columns)]

# Create DataFrame
combined_data = pd.DataFrame(all_rows, columns=column_names)

# Write the combined data to an Excel file
try:
    combined_data.to_excel(output_file, index=False)
    print(f"Combined data written to {output_file}")
    print(f"Total rows: {len(combined_data)}, Total columns: {len(combined_data.columns)}")
except Exception as e:
    print(f"Error writing to Excel file: {e}")