import re
import csv
import pandas as pd
from pathlib import Path

def extract_components_from_branches(file_path):
    """
    Extract all piping components from branches in E3D database listing files.
    
    Components: FLANGE, ELBOW, TEE, REDUCER, VALVE, ATTACHMENT, etc.
    Each component has a SPRE SPCOMPONENT property.
    """
    kks_pattern = r'\d[A-Z]{3}\d{2}BR\d{3}'
    branch_pattern = r'/B\d+'
    
    components = []
    current_pipe = None
    current_branch = None
    current_component_type = None
    
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            # Check if line contains "NEW PIPE"
            if 'NEW PIPE' in line:
                match = re.search(kks_pattern, line)
                if match:
                    current_pipe = match.group()
                    current_branch = None
                    current_component_type = None
            
            # Check if line contains "NEW BRANCH"
            elif 'NEW BRANCH' in line and current_pipe:
                full_match = re.search(kks_pattern + branch_pattern, line)
                if full_match:
                    full_branch = full_match.group()
                    branch_match = re.search(branch_pattern, full_branch)
                    if branch_match:
                        current_branch = branch_match.group()
                        current_component_type = None
            
            # Check if line starts with "NEW " followed by a component type
            elif line.strip().startswith('NEW ') and current_branch:
                # Extract component type (e.g., "FLANGE", "ELBOW", "TEE")
                parts = line.strip().split()
                if len(parts) >= 2:
                    component_type = parts[1]
                    # Skip "BRANCH", "PIPE", "ZONE", "SITE", etc.
                    if component_type not in ['BRANCH', 'PIPE', 'ZONE', 'SITE', 'STRUCTURE', 'SUBSTRUCTURE', 'CYLINDER', 'CTORUS']:
                        current_component_type = component_type
            
            # Check if line contains "SPRE SPCOMPONENT"
            elif 'SPRE SPCOMPONENT' in line and current_component_type:
                # Extract the SPRE value after "SPRE SPCOMPONENT"
                spre_match = re.search(r'SPRE SPCOMPONENT\s+(\S+)', line)
                if spre_match:
                    spre_value = spre_match.group(1)
                    components.append({
                        'Pipe': current_pipe,
                        'Branch': current_branch,
                        'Component_Type': current_component_type,
                        'SPRE': spre_value
                    })
    
    return components

def lookup_and_merge_with_excel(components, excel_file):
    """
    Lookup SPRE values in Excel file and merge data.
    Add 'Found', 'P1 CONN', 'P2 CONN', and 'Welded' columns.
    """
    # Read Excel file
    df_excel = pd.read_excel(excel_file)
    
    # Create a dictionary for quick lookup
    spre_lookup = {}
    for _, row in df_excel.iterrows():
        spre = row['SPRE']
        spre_lookup[spre] = {
            'P1_CONN': row['P1 CONN'] if pd.notna(row['P1 CONN']) else '',
            'P2_CONN': row['P2 CONN'] if pd.notna(row['P2 CONN']) else ''
        }
    
    # Process components
    result = []
    for comp in components:
        spre = comp['SPRE']
        
        if spre in spre_lookup:
            found = 'X'
            p1_conn = spre_lookup[spre]['P1_CONN']
            p2_conn = spre_lookup[spre]['P2_CONN']
            welded = 'X' if (p1_conn == 'BWD' or p2_conn == 'BWD') else ''
        else:
            found = ''
            p1_conn = ''
            p2_conn = ''
            welded = ''
        
        result.append({
            'Pipe': comp['Pipe'],
            'Branch': comp['Branch'],
            'Component_Type': comp['Component_Type'],
            'SPRE': spre,
            'Found': found,
            'P1_CONN': p1_conn,
            'P2_CONN': p2_conn,
            'Welded': welded
        })
    
    return result

def save_to_csv(data, output_file):
    """
    Save components data to a CSV file.
    """
    if not data:
        print("No data to save.")
        return
    
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Pipe', 'Branch', 'Component_Type', 'SPRE', 'Found', 'P1_CONN', 'P2_CONN', 'Welded']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for row in data:
            writer.writerow(row)
    
    print(f"Data saved to: {output_file}")
    print(f"Total components written: {len(data)}")

if __name__ == "__main__":
    # Define file paths
    txt_files = [
        Path(__file__).parent / "TBY" / "TBY-0AUX-P.txt",
        Path(__file__).parent / "TBY" / "TBY-1AUX-P.txt"
    ]
    excel_file = Path(__file__).parent / "TBY" / "TBY_all_pspecs_wure_macro_08.12.2025.xlsx"
    output_csv = Path(__file__).parent / "TBY" / "components_with_welds.csv"
    
    all_components = []
    
    # Process each TXT file
    for txt_file in txt_files:
        print(f"Reading TXT file: {txt_file}")
        
        if not txt_file.exists():
            print(f"  Warning: File not found, skipping...")
            continue
        
        # Extract components from branches
        components = extract_components_from_branches(txt_file)
        print(f"  Extracted {len(components)} components")
        all_components.extend(components)
    
    print(f"\nTotal components from all files: {len(all_components)}")
    print(f"\nReading Excel file: {excel_file}")
    
    # Lookup and merge with Excel data
    result = lookup_and_merge_with_excel(all_components, excel_file)
    
    # Save to CSV
    save_to_csv(result, output_csv)
    
    # Display summary statistics
    found_count = sum(1 for r in result if r['Found'] == 'X')
    welded_count = sum(1 for r in result if r['Welded'] == 'X')
    
    print(f"\nSummary:")
    print(f"Total components: {len(result)}")
    print(f"Found in Excel: {found_count}")
    print(f"Not found in Excel: {len(result) - found_count}")
    print(f"Welded components (BWD): {welded_count}")
