import pandas as pd
import re
from pathlib import Path

# Read Excel
excel_file = Path(r"C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\TBY_all_pspecs_wure_macro_08.12.2025.xlsx")
df_excel = pd.read_excel(excel_file)

# Check VALVE data
valve_spre = '/HZI-GP-EN-VALV/Z-2JVVLG022FF'
valve_data = df_excel[df_excel['SPRE'] == valve_spre]
print("VALVE in Excel:")
print(valve_data[['SPRE', 'TYPE', 'PBOR', 'PBOR1', 'P1 CONN', 'P2 CONN']])

# Parse PBOR values
pbor_str = valve_data['PBOR'].values[0] if pd.notna(valve_data['PBOR'].values[0]) else '0mm'
pbor1_str = valve_data['PBOR1'].values[0] if pd.notna(valve_data['PBOR1'].values[0]) else '0mm'

try:
    pbor = float(str(pbor_str).replace('mm', ''))
except:
    pbor = 0.0

try:
    pbor1 = float(str(pbor1_str).replace('mm', ''))
except:
    pbor1 = 0.0

print(f"\nPBOR: {pbor}")
print(f"PBOR1: {pbor1}")

# Calculate length
if pbor > 0:
    length = 0.5 * pbor
    print(f"\nUsing PBOR: Length = 0.5 * {pbor} = {length}")
elif pbor1 > 0:
    length = 0.5 * pbor1
    print(f"\nUsing PBOR1 (fallback): Length = 0.5 * {pbor1} = {length}")
else:
    length = None
    print("\nNo valid length!")

# Check if welded
p1_conn = valve_data['P1 CONN'].values[0] if pd.notna(valve_data['P1 CONN'].values[0]) else ''
p2_conn = valve_data['P2 CONN'].values[0] if pd.notna(valve_data['P2 CONN'].values[0]) else ''
comp_type = valve_data['TYPE'].values[0] if pd.notna(valve_data['TYPE'].values[0]) else ''

is_welded = (p1_conn == 'BWD' or p2_conn == 'BWD' or comp_type == 'OLET')
print(f"Is welded: {is_welded} (P1={p1_conn}, P2={p2_conn}, TYPE={comp_type})")

# Check branch positions
df_branches = pd.read_csv(r'C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\connected_branches.csv')
branch_0QCQ = df_branches[(df_branches['Branch_A']=='0QCQ10BR100/B1') | (df_branches['Branch_B']=='0QCQ10BR100/B1')]
print(f"\n0QCQ10BR100/B1 in connected_branches.csv: {len(branch_0QCQ)} entries")

# Parse TXT file to find VALVE position
txt_file = Path(r"C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\TBY-0AUX-P.txt")
kks_pattern = r'\d[A-Z]{3}\d{2}BR\d{3}'
branch_pattern = r'/B\d+'

current_pipe = None
current_branch = None
in_0QCQ = False

with open(txt_file, 'r', encoding='utf-8') as f:
    for line in f:
        if 'NEW PIPE' in line:
            match = re.search(kks_pattern, line)
            if match:
                current_pipe = match.group()
                in_0QCQ = (current_pipe == '0QCQ10BR100')
        
        elif 'NEW BRANCH' in line and current_pipe:
            full_match = re.search(kks_pattern + branch_pattern, line)
            if full_match:
                full_branch = full_match.group()
                branch_match = re.search(branch_pattern, full_branch)
                if branch_match:
                    current_branch = branch_match.group()
                    if in_0QCQ and current_branch == '/B1':
                        print(f"\n=== Found branch {current_pipe}{current_branch} ===")
        
        elif in_0QCQ and current_branch == '/B1':
            if 'HPOS' in line:
                print(f"HPOS: {line.strip()}")
            elif 'SPRE SPCOMPONENT /HZI-GP-EN-VALV/Z-2JVVLG022FF' in line:
                print(f"Found VALVE SPRE!")
            elif line.strip().startswith('POS ') and 'VALVE' in str(locals().get('last_line', '')):
                print(f"VALVE POS: {line.strip()}")
        
        if in_0QCQ and current_branch == '/B1':
            last_line = line
