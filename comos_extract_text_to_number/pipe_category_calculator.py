import pandas as pd
import os
import numpy as np
import re

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

# Load categorization rules
rules_path = os.path.join(os.path.dirname(__file__), 'pipe categories according to EN13480-1 Table 5.1-1.csv')
df_rules = pd.read_csv(rules_path)

# Enhanced medium processing
def process_medium(medium):
    if pd.isna(medium):
        return 'Unknown'
    
    m = str(medium).lower().strip()
    
    # Liquids indicators
    liquid_indicators = ['water', 'liquid', 'oil', 'fuel', 'hydraulic', 'coolant']
    
    # Gas indicators  
    gas_indicators = ['air', 'gas', 'steam', 'vapor', 'nitrogen', 'oxygen', 'co2', 'methane']
    
    # Check for liquid indicators
    for indicator in liquid_indicators:
        if indicator in m:
            return 'Liquids'
    
    # Check for gas indicators
    for indicator in gas_indicators:
        if indicator in m:
            return 'Gases'
    
    # Default to Gases if uncertain (conservative approach)
    return 'Gases'

df['Medium_processed'] = df['Medium'].apply(process_medium)

# Enhanced categorization function with correct rule priority
def categorize_pipe(row):
    pipe_medium = row['Medium_processed']
    pipe_group = str(row.get('Fluid Group', '')).strip()
    
    # Handle data conversion with proper error handling
    try:
        pipe_dn = float(row['DN'])
        if pipe_dn < 0:
            return 'Invalid Data', 'Negative DN value'
    except (ValueError, TypeError):
        return 'Invalid Data', 'Invalid DN value'
    
    try:
        pipe_ps = float(row['PS [bar(g)]'])
        if pipe_ps < 0:
            return 'Invalid Data', 'Negative PS value'
    except (ValueError, TypeError):
        return 'Invalid Data', 'Invalid PS value'
    
    # Calculate PS*DN
    ps_dn = pipe_ps * pipe_dn
    
    # Handle unknown medium
    if pipe_medium == 'Unknown':
        return 'Invalid Data', 'Unknown medium type'
    
    # Filter rules: First try exact fluid group match, then 'any'
    exact_rules = df_rules[(df_rules['Medium'] == pipe_medium) & (df_rules['Fluid group'] == pipe_group)]
    any_rules = df_rules[(df_rules['Medium'] == pipe_medium) & (df_rules['Fluid group'] == 'any')]
    
    # Define category priority order (highest to lowest)
    category_priority = ['Category III', 'Category II', 'Category I', 'Category 0', 'See 5.3']
    
    # Collect all matching rules with their results
    matching_rules = []
    
    # Try exact match first, then 'any' rules
    for rules_set in [exact_rules, any_rules]:
        if len(rules_set) == 0:
            continue
            
        for _, rule in rules_set.iterrows():
            criteria = rule['Criteria'].strip()
            category = rule['Category'].strip()
            
            try:
                # Create safe evaluation environment
                safe_dict = {
                    "PS": pipe_ps,
                    "DN": pipe_dn,
                    "__builtins__": {}
                }
                
                # Replace logical operators for Python
                expr = criteria.replace('AND', 'and').replace('OR', 'or')
                
                # Evaluate the expression
                if eval(expr, safe_dict):
                    matching_rules.append({
                        'category': category,
                        'criteria': criteria,
                        'priority': category_priority.index(category) if category in category_priority else 99
                    })
                    
            except Exception as e:
                print(f"Error evaluating criteria '{criteria}': {e}")
                continue
    
    # If we have matching rules, return the highest priority one
    if matching_rules:
        # Sort by priority (lower index = higher priority)
        matching_rules.sort(key=lambda x: x['priority'])
        best_match = matching_rules[0]
        return best_match['category'], f"{best_match['criteria']} (PS={pipe_ps}, DN={pipe_dn}, PS*DN={ps_dn})"
    
    return 'Not Categorized', f'No matching rule found (Medium: {pipe_medium}, Group: {pipe_group}, PS: {pipe_ps}, DN: {pipe_dn})'

# Apply categorization
print("Starting categorization...")
df[['Category', 'Category Details']] = df.apply(categorize_pipe, axis=1, result_type='expand')

# Print summary statistics
print("\nCategorization Summary:")
print(df['Category'].value_counts())

# Create output filename with '_fixed' postfix
base, ext = os.path.splitext(input_file)
output_file = f"{base}_categorized{ext}"
output_path = os.path.join(os.path.dirname(__file__), output_file)

# Save the fixed DataFrame to Excel
try:
    df.to_excel(output_path, index=False)
    print(f"\nFile saved: {output_path}")
except Exception as e:
    print(f"Error saving file: {e}")

# Optional: Save detailed results for debugging
debug_file = f"{base}_debug.csv"
debug_path = os.path.join(os.path.dirname(__file__), debug_file)
try:
    df[['KKS', 'Medium', 'Medium_processed', 'Fluid Group', 'DN', 'PS [bar(g)]', 'Category', 'Category Details']].to_csv(debug_path, index=False)
    print(f"Debug file saved: {debug_path}")
except Exception as e:
    print(f"Error saving debug file: {e}")