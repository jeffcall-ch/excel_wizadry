import re

# Simulate extract_branch_positions to see if HCON/TCON are being extracted
file_path = r'C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\TBY-0AUX-P.txt'
kks_pattern = r'\d[A-Z]{3}\d{2}BR\d{3}'
branch_pattern = r'/B\d+'

branches = []
current_pipe = None
current_branch = None
current_branch_info = {}
in_branch = False
count = 0

with open(file_path, 'r', encoding='utf-8') as f:
    for line in f:
        # Check if line contains "NEW PIPE"
        if 'NEW PIPE' in line:
            match = re.search(kks_pattern, line)
            if match:
                current_pipe = match.group()
                current_branch = None
                in_branch = False
        
        # Check if line contains "NEW BRANCH"
        elif 'NEW BRANCH' in line and current_pipe:
            # Save previous branch if exists
            if current_branch_info and current_branch_info.get('Pipe'):
                branches.append(current_branch_info.copy())
            
            full_match = re.search(kks_pattern + branch_pattern, line)
            if full_match:
                full_branch = full_match.group()
                branch_match = re.search(branch_pattern, full_branch)
                if branch_match:
                    current_branch = branch_match.group()
                    current_branch_info = {
                        'Pipe': current_pipe,
                        'Branch': current_branch,
                        'Full_Branch_ID': current_pipe + current_branch,
                        'HCON': '',
                        'TCON': '',
                        'HPOS_X': None,
                        'HPOS_Y': None,
                        'HPOS_Z': None,
                        'TPOS_X': None,
                        'TPOS_Y': None,
                        'TPOS_Z': None
                    }
                    in_branch = True
        
        # Extract HCON, TCON, HPOS, TPOS
        elif in_branch and current_branch_info:
            if line.strip().startswith('HCON '):
                hcon_match = re.search(r'HCON\s+(\S+)', line)
                if hcon_match:
                    current_branch_info['HCON'] = hcon_match.group(1)
                    count += 1
                    if count <= 5:
                        print(f"Found HCON: {current_branch_info['Full_Branch_ID']} = {current_branch_info['HCON']}")
            
            elif line.strip().startswith('TCON '):
                tcon_match = re.search(r'TCON\s+(\S+)', line)
                if tcon_match:
                    current_branch_info['TCON'] = tcon_match.group(1)
                    if count <= 5:
                        print(f"Found TCON: {current_branch_info['Full_Branch_ID']} = {current_branch_info['TCON']}")
            
            elif line.strip().startswith('HPOS '):
                hpos_match = re.search(r'HPOS X ([\d.-]+)mm Y ([\d.-]+)mm Z ([\d.-]+)mm', line)
                if hpos_match:
                    current_branch_info['HPOS_X'] = float(hpos_match.group(1))
                    current_branch_info['HPOS_Y'] = float(hpos_match.group(2))
                    current_branch_info['HPOS_Z'] = float(hpos_match.group(3))
            
            elif line.strip().startswith('TPOS '):
                tpos_match = re.search(r'TPOS X ([\d.-]+)mm Y ([\d.-]+)mm Z ([\d.-]+)mm', line)
                if tpos_match:
                    current_branch_info['TPOS_X'] = float(tpos_match.group(1))
                    current_branch_info['TPOS_Y'] = float(tpos_match.group(2))
                    current_branch_info['TPOS_Z'] = float(tpos_match.group(3))
                    in_branch = False

# Don't forget the last branch
if current_branch_info and current_branch_info.get('Pipe'):
    branches.append(current_branch_info)

print(f"\nTotal branches extracted: {len(branches)}")
print(f"Branches with HCON='BWD': {sum(1 for b in branches if b['HCON'] == 'BWD')}")
print(f"Branches with TCON='BWD': {sum(1 for b in branches if b['TCON'] == 'BWD')}")
print(f"Branches with both HPOS and TPOS: {sum(1 for b in branches if b['HPOS_X'] is not None and b['TPOS_X'] is not None)}")

print("\nFirst 5 branches with positions:")
for b in branches[:5]:
    if b['HPOS_X'] is not None:
        print(f"{b['Full_Branch_ID']}: HCON={b['HCON']}, TCON={b['TCON']}, Has positions: HPOS={b['HPOS_X'] is not None}, TPOS={b['TPOS_X'] is not None}")
