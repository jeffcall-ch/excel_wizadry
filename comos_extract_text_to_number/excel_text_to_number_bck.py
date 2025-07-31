import pandas as pd
import os

# Hardcoded input file path
input_file = 'Export_pipe_list_30.07.2025_withfluid_code.xlsx'
input_path = os.path.join(os.path.dirname(__file__), input_file)

# Read the first sheet of the Excel file
try:
    df = pd.read_excel(input_path, dtype=str)
except Exception as e:
    print(f"Error reading file: {e}")
    exit(1)


# Convert all columns to numeric, preserving decimal places as in the original text
import re
def preserve_decimals(val):
    if isinstance(val, str):
        # Only extract number for 'DN <number>' or 'PN <number>' patterns
        match = re.match(r'^(DN|PN)\s*([-+]?\d*\.\d+|[-+]?\d+)$', val.strip(), re.IGNORECASE)
        if match:
            num_str = match.group(2)
            try:
                if '.' in num_str:
                    decimals = len(num_str.split('.')[-1])
                    num = float(num_str)
                    return float(f"{num:.{decimals}f}")
                else:
                    return int(num_str)
            except Exception:
                return val
        # Otherwise, keep original text
        return val
    return val



for col in df.columns:
    df[col] = df[col].apply(preserve_decimals)

# --- Categorization logic using rules from CSV ---
import numpy as np

# Load rules
rules_path = os.path.join(os.path.dirname(__file__), 'pipe categories according to EN13480-1 Table 5.1-1.csv')
df_rules = pd.read_csv(rules_path)

# Step 2: Preprocess Medium
def process_medium(medium):
    m = str(medium).lower()
    if 'water' in m or 'air' in m:
        return 'Liquids'
    return 'Gases'
df['Medium_processed'] = df['Medium'].apply(process_medium)

# Step 3: Categorize Each Pipe
def categorize_pipe(row):
    pipe_medium = row['Medium_processed']
    pipe_group = str(row.get('Fluid Group', '')).strip()
    try:
        pipe_dn = float(row['DN'])
    except Exception:
        pipe_dn = np.nan
    try:
        pipe_ps = float(row['PS [bar(g)]'])
    except Exception:
        pipe_ps = np.nan

    # Filter rules for medium and fluid group
    rules = df_rules[(df_rules['Medium'] == pipe_medium) & ((df_rules['Fluid group'] == pipe_group) | (df_rules['Fluid group'].str.lower() == 'any'))]
    for _, rule in rules.iterrows():
        criteria = rule['Criteria']
        category = rule['Category'].strip()
        # Prepare variables for eval
        PS = pipe_ps
        DN = pipe_dn
        try:
            # Replace logical operators for Python
            expr = criteria.replace('AND', 'and').replace('OR', 'or')
            # Evaluate
            if eval(expr, {"PS": PS, "DN": DN}):
                return category, criteria
        except Exception:
            continue
    return 'Not Categorized', 'No matching rule'

df[['Category', 'Category Details']] = df.apply(categorize_pipe, axis=1, result_type='expand')

# Create output filename with '_fixed' postfix

base, ext = os.path.splitext(input_file)
output_file = f"{base}_fixed{ext}"
output_path = os.path.join(os.path.dirname(__file__), output_file)

# Save the fixed DataFrame to Excel
try:
    df.to_excel(output_path, index=False)
    print(f"File saved: {output_path}")
except Exception as e:
    print(f"Error saving file: {e}")
