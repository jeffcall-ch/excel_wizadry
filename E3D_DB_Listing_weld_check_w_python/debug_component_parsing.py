import re
from pathlib import Path

# Quick test to see if we're parsing components with positions
file_path = Path(r"C:\Users\szil\Repos\excel_wizadry\E3D_DB_Listing_weld_check_w_python\TBY\TBY-0AUX-P.txt")
kks_pattern = r'\d[A-Z]{3}\d{2}BR\d{3}'
branch_pattern = r'/B\d+'

current_pipe = None
current_branch = None
current_component = None
count = 0

with open(file_path, 'r', encoding='utf-8') as f:
    for line in f:
        if 'NEW PIPE' in line:
            match = re.search(kks_pattern, line)
            if match:
                current_pipe = match.group()
                current_branch = None
                current_component = None
        
        elif 'NEW BRANCH' in line and current_pipe:
            full_match = re.search(kks_pattern + branch_pattern, line)
            if full_match:
                full_branch = full_match.group()
                branch_match = re.search(branch_pattern, full_branch)
                if branch_match:
                    current_branch = branch_match.group()
                    current_component = None
        
        elif current_pipe and current_branch:
            if line.strip().startswith('SPRE '):
                spre_match = re.search(r'SPRE\s+(\S+)', line)
                if spre_match:
                    spre = spre_match.group(1)
                    if current_component is None:
                        current_component = {'spre': spre, 'pos': None}
                    else:
                        current_component['spre'] = spre
            
            elif line.strip().startswith('POS '):
                pos_match = re.search(r'POS X ([\d.-]+)mm Y ([\d.-]+)mm Z ([\d.-]+)mm', line)
                if pos_match and current_component:
                    count += 1
                    if count <= 5:
                        print(f"Found component: {current_component['spre']} in {current_pipe}{current_branch}")
                        print(f"  Position: X={pos_match.group(1)}, Y={pos_match.group(2)}, Z={pos_match.group(3)}")
                    current_component = None

print(f"\nTotal components with positions found: {count}")
