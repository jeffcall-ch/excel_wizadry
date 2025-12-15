import re
import csv
import pandas as pd
from pathlib import Path

def extract_branch_positions(file_path):
    """
    Extract branch positions (HPOS and TPOS) from E3D database listing files.
    
    Returns list of branches with their head and tail positions.
    """
    kks_pattern = r'\d[A-Z]{3}\d{2}BR\d{3}'
    branch_pattern = r'/B\d+'
    
    branches = []
    current_pipe = None
    current_branch = None
    current_branch_info = {}
    in_branch = False
    
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
            
            # Extract HPOS, TPOS, HCON, TCON
            elif in_branch and current_branch_info:
                if line.strip().startswith('HCON '):
                    hcon_match = re.search(r'HCON\s+(\S+)', line)
                    if hcon_match:
                        current_branch_info['HCON'] = hcon_match.group(1)
                
                elif line.strip().startswith('TCON '):
                    tcon_match = re.search(r'TCON\s+(\S+)', line)
                    if tcon_match:
                        current_branch_info['TCON'] = tcon_match.group(1)
                
                elif line.strip().startswith('HPOS '):
                    # Extract coordinates from HPOS line
                    hpos_match = re.search(r'HPOS X ([\d.-]+)mm Y ([\d.-]+)mm Z ([\d.-]+)mm', line)
                    if hpos_match:
                        current_branch_info['HPOS_X'] = float(hpos_match.group(1))
                        current_branch_info['HPOS_Y'] = float(hpos_match.group(2))
                        current_branch_info['HPOS_Z'] = float(hpos_match.group(3))
                
                elif line.strip().startswith('TPOS '):
                    # Extract coordinates from TPOS line
                    tpos_match = re.search(r'TPOS X ([\d.-]+)mm Y ([\d.-]+)mm Z ([\d.-]+)mm', line)
                    if tpos_match:
                        current_branch_info['TPOS_X'] = float(tpos_match.group(1))
                        current_branch_info['TPOS_Y'] = float(tpos_match.group(2))
                        current_branch_info['TPOS_Z'] = float(tpos_match.group(3))
    
    # Don't forget the last branch
    if current_branch_info and current_branch_info.get('Pipe'):
        branches.append(current_branch_info)
    
    return branches

def find_connected_branches(branches, tolerance_tight=5.0, tolerance_loose=150.0):
    """
    Find branches that are connected based on axis-combination matching.
    Uses logic where two axes must be tight and third can be loose:
    1. X and Y within tolerance_tight, Z within tolerance_loose
    2. X and Z within tolerance_tight, Y within tolerance_loose
    3. Y and Z within tolerance_tight, X within tolerance_loose
    
    Only includes connections where at least one of TCON (of branch1) or HCON (of branch2) is 'BWD' (welded).
    
    Args:
        branches: List of branch dictionaries with position data
        tolerance_tight: Maximum offset in mm for two axes (default 5.0mm)
        tolerance_loose: Maximum offset in mm for third axis (default 150.0mm)
    
    Returns:
        List of connection dictionaries with paired branches side by side
    """
    connections = []
    processed_pairs = set()
    
    for i, branch1 in enumerate(branches):
        # Skip if branch doesn't have tail position
        if branch1['TPOS_X'] is None:
            continue
        
        tail_x = branch1['TPOS_X']
        tail_y = branch1['TPOS_Y']
        tail_z = branch1['TPOS_Z']
        
        for j, branch2 in enumerate(branches):
            # Skip same branch or if branch doesn't have head position
            if i == j or branch2['HPOS_X'] is None:
                continue
            
            head_x = branch2['HPOS_X']
            head_y = branch2['HPOS_Y']
            head_z = branch2['HPOS_Z']
            
            # Calculate individual axis offsets
            offset_x = abs(tail_x - head_x)
            offset_y = abs(tail_y - head_y)
            offset_z = abs(tail_z - head_z)
            
            # Calculate 3D distance
            distance = ((tail_x - head_x)**2 + 
                       (tail_y - head_y)**2 + 
                       (tail_z - head_z)**2)**0.5
            
            # Check all three axis combinations
            match_type = None
            
            # Combination 1: X and Y tight, Z loose
            if offset_x <= tolerance_tight and offset_y <= tolerance_tight and offset_z <= tolerance_loose:
                match_type = 'XY_tight_Z_loose'
            
            # Combination 2: X and Z tight, Y loose
            elif offset_x <= tolerance_tight and offset_z <= tolerance_tight and offset_y <= tolerance_loose:
                match_type = 'XZ_tight_Y_loose'
            
            # Combination 3: Y and Z tight, X loose
            elif offset_y <= tolerance_tight and offset_z <= tolerance_tight and offset_x <= tolerance_loose:
                match_type = 'YZ_tight_X_loose'
            
            # If any match found, record connection
            if match_type:
                # Check if at least one of the connection ends is welded (BWD)
                tcon_bwd = branch1.get('TCON', '') == 'BWD'
                hcon_bwd = branch2.get('HCON', '') == 'BWD'
                
                # Skip this connection if neither end is welded
                if not (tcon_bwd or hcon_bwd):
                    continue
                
                # Create a pair key to avoid duplicates (order-independent)
                pair_key = tuple(sorted([branch1['Full_Branch_ID'], branch2['Full_Branch_ID']]))
                
                # Only add if this pair hasn't been processed
                if pair_key not in processed_pairs:
                    processed_pairs.add(pair_key)
                    
                    # Calculate accuracy metric: maximum offset among the two tight axes
                    if match_type == 'XY_tight_Z_loose':
                        max_tight_offset = max(offset_x, offset_y)
                        loose_offset = offset_z
                    elif match_type == 'XZ_tight_Y_loose':
                        max_tight_offset = max(offset_x, offset_z)
                        loose_offset = offset_y
                    else:  # YZ_tight_X_loose
                        max_tight_offset = max(offset_y, offset_z)
                        loose_offset = offset_x
                    
                    # Accuracy percentage: how close are the tight axes to perfect (0mm)
                    # 100% = perfect (0mm), decreases as offset approaches tolerance_tight (5mm)
                    accuracy_pct = max(0, 100 * (1 - max_tight_offset / tolerance_tight))
                    
                    connections.append({
                        'Branch_A': branch1['Full_Branch_ID'],
                        'Branch_A_TCON': branch1.get('TCON', ''),
                        'Branch_B': branch2['Full_Branch_ID'],
                        'Branch_B_HCON': branch2.get('HCON', ''),
                        'Branch_A_Pipe': branch1['Pipe'],
                        'Branch_A_Branch': branch1['Branch'],
                        'Branch_B_Pipe': branch2['Pipe'],
                        'Branch_B_Branch': branch2['Branch'],
                        'Connection_X': tail_x,
                        'Connection_Y': tail_y,
                        'Connection_Z': tail_z,
                        'Distance_mm': round(distance, 3),
                        'Offset_X_mm': round(offset_x, 3),
                        'Offset_Y_mm': round(offset_y, 3),
                        'Offset_Z_mm': round(offset_z, 3),
                        'Max_Tight_Offset_mm': round(max_tight_offset, 3),
                        'Loose_Offset_mm': round(loose_offset, 3),
                        'Accuracy_Percent': round(accuracy_pct, 1),
                        'Match_Type': match_type
                    })
    
    return connections

def extract_branch_connections(file_path):
    """
    Extract branch connection information from E3D database listing files.
    
    Returns list of branches with their HCON, TCON, first/last components, and pipe lengths.
    """
    kks_pattern = r'\d[A-Z]{3}\d{2}BR\d{3}'
    branch_pattern = r'/B\d+'
    
    branches = []
    current_pipe = None
    current_branch = None
    current_branch_info = {}
    in_branch = False
    branch_depth = 0
    first_component_pos = None
    last_component_pos = None
    current_component_is_valid = False  # Track if current NEW component is valid
    
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            # Check if line contains "NEW PIPE"
            if 'NEW PIPE' in line:
                # Save previous branch before resetting (if exists)
                if current_branch_info and current_branch_info.get('Pipe'):
                    # Calculate pipe lengths before saving branch
                    if current_branch_info['HPOS_X'] is not None and first_component_pos is not None:
                        hpos = (current_branch_info['HPOS_X'], current_branch_info['HPOS_Y'], current_branch_info['HPOS_Z'])
                        head_pipe_length = ((hpos[0] - first_component_pos[0])**2 + 
                                           (hpos[1] - first_component_pos[1])**2 + 
                                           (hpos[2] - first_component_pos[2])**2)**0.5
                        current_branch_info['Head_Pipe_Length_mm'] = round(head_pipe_length, 2)
                    
                    if current_branch_info['TPOS_X'] is not None and last_component_pos is not None:
                        tpos = (current_branch_info['TPOS_X'], current_branch_info['TPOS_Y'], current_branch_info['TPOS_Z'])
                        tail_pipe_length = ((tpos[0] - last_component_pos[0])**2 + 
                                           (tpos[1] - last_component_pos[1])**2 + 
                                           (tpos[2] - last_component_pos[2])**2)**0.5
                        current_branch_info['Tail_Pipe_Length_mm'] = round(tail_pipe_length, 2)
                    
                    branches.append(current_branch_info.copy())
                    current_branch_info = {}  # Reset to prevent duplicate saves
                
                match = re.search(kks_pattern, line)
                if match:
                    current_pipe = match.group()
                    current_branch = None
                    in_branch = False
                    branch_depth = 0
                    first_component_pos = None
                    last_component_pos = None
                    current_component_is_valid = False
            
            # Check if line contains "NEW BRANCH"
            elif 'NEW BRANCH' in line and current_pipe:
                # Save previous branch if exists (with pipe length calculation)
                if current_branch_info and current_branch_info.get('Pipe'):
                    # Calculate pipe lengths before saving branch using current position variables
                    if current_branch_info['HPOS_X'] is not None and first_component_pos is not None:
                        hpos = (current_branch_info['HPOS_X'], current_branch_info['HPOS_Y'], current_branch_info['HPOS_Z'])
                        head_pipe_length = ((hpos[0] - first_component_pos[0])**2 + 
                                           (hpos[1] - first_component_pos[1])**2 + 
                                           (hpos[2] - first_component_pos[2])**2)**0.5
                        current_branch_info['Head_Pipe_Length_mm'] = round(head_pipe_length, 2)
                    
                    if current_branch_info['TPOS_X'] is not None and last_component_pos is not None:
                        tpos = (current_branch_info['TPOS_X'], current_branch_info['TPOS_Y'], current_branch_info['TPOS_Z'])
                        tail_pipe_length = ((tpos[0] - last_component_pos[0])**2 + 
                                           (tpos[1] - last_component_pos[1])**2 + 
                                           (tpos[2] - last_component_pos[2])**2)**0.5
                        current_branch_info['Tail_Pipe_Length_mm'] = round(tail_pipe_length, 2)
                    
                    branches.append(current_branch_info.copy())
                
                # Now reset for the new branch
                first_component_pos = None
                last_component_pos = None
                current_component_is_valid = False
                
                full_match = re.search(kks_pattern + branch_pattern, line)
                if full_match:
                    full_branch = full_match.group()
                    branch_match = re.search(branch_pattern, full_branch)
                    if branch_match:
                        current_branch = branch_match.group()
                        current_branch_info = {
                            'Pipe': current_pipe,
                            'Branch': current_branch,
                            'HCON': '',
                            'TCON': '',
                            'HSTU': '',
                            'First_Component': '',
                            'Last_Component': '',
                            'HPOS_X': None,
                            'HPOS_Y': None,
                            'HPOS_Z': None,
                            'TPOS_X': None,
                            'TPOS_Y': None,
                            'TPOS_Z': None,
                            'Head_Pipe_Length_mm': 0,
                            'Tail_Pipe_Length_mm': 0
                        }
                        in_branch = True
                        branch_depth = 0
            
            # Extract HCON, TCON, HSTU, HPOS, TPOS when inside a branch
            elif in_branch and current_branch_info:
                if line.strip().startswith('HPOS '):
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
                
                elif line.strip().startswith('HCON '):
                    hcon_match = re.search(r'HCON\s+(\S+)', line)
                    if hcon_match:
                        current_branch_info['HCON'] = hcon_match.group(1)
                
                elif line.strip().startswith('TCON '):
                    tcon_match = re.search(r'TCON\s+(\S+)', line)
                    if tcon_match:
                        current_branch_info['TCON'] = tcon_match.group(1)
                
                elif line.strip().startswith('HSTU SPCOMPONENT'):
                    hstu_match = re.search(r'HSTU SPCOMPONENT\s+(\S+)', line)
                    if hstu_match:
                        current_branch_info['HSTU'] = hstu_match.group(1)
                
                # Detect first component in branch (first NEW component after branch definition)
                elif line.strip().startswith('NEW '):
                    parts = line.strip().split()
                    if len(parts) >= 2:
                        component_type = parts[1]
                        if component_type not in ['BRANCH', 'PIPE', 'ZONE', 'SITE', 'STRUCTURE', 'SUBSTRUCTURE', 'CYLINDER', 'CTORUS', 'ATTACHMENT']:
                            current_component_is_valid = True
                            if not current_branch_info['First_Component']:
                                current_branch_info['First_Component'] = component_type
                            # Always update last component
                            current_branch_info['Last_Component'] = component_type
                        else:
                            current_component_is_valid = False
                
                # Extract component position (POS line) - only for valid components
                elif line.strip().startswith('POS ') and current_component_is_valid:
                    pos_match = re.search(r'POS X ([\d.-]+)mm Y ([\d.-]+)mm Z ([\d.-]+)mm', line)
                    if pos_match:
                        comp_x = float(pos_match.group(1))
                        comp_y = float(pos_match.group(2))
                        comp_z = float(pos_match.group(3))
                        
                        # Track first component position
                        if first_component_pos is None and current_branch_info['First_Component']:
                            first_component_pos = (comp_x, comp_y, comp_z)
                        
                        # Always update last component position for valid components
                        last_component_pos = (comp_x, comp_y, comp_z)
    
    # Don't forget the last branch (with pipe length calculation)
    if current_branch_info and current_branch_info.get('Pipe'):
        # Calculate pipe lengths before saving last branch
        if current_branch_info['HPOS_X'] is not None and first_component_pos is not None:
            hpos = (current_branch_info['HPOS_X'], current_branch_info['HPOS_Y'], current_branch_info['HPOS_Z'])
            head_pipe_length = ((hpos[0] - first_component_pos[0])**2 + 
                               (hpos[1] - first_component_pos[1])**2 + 
                               (hpos[2] - first_component_pos[2])**2)**0.5
            current_branch_info['Head_Pipe_Length_mm'] = round(head_pipe_length, 2)
        
        if current_branch_info['TPOS_X'] is not None and last_component_pos is not None:
            tpos = (current_branch_info['TPOS_X'], current_branch_info['TPOS_Y'], current_branch_info['TPOS_Z'])
            tail_pipe_length = ((tpos[0] - last_component_pos[0])**2 + 
                               (tpos[1] - last_component_pos[1])**2 + 
                               (tpos[2] - last_component_pos[2])**2)**0.5
            current_branch_info['Tail_Pipe_Length_mm'] = round(tail_pipe_length, 2)
        
        branches.append(current_branch_info)
    
    return branches

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
    Add 'Found', 'P1 CONN', 'P2 CONN', 'TYPE', 'Welded', and 'Weld_Count' columns.
    Welded components are those with:
    - P1 CONN = 'BWD'
    - P2 CONN = 'BWD'
    - TYPE = 'OLET'
    
    Weld count calculation:
    - P1 CONN = 'BWD' only: 1 weld
    - P2 CONN = 'BWD' only: 1 weld
    - Both P1 and P2 CONN = 'BWD': 2 welds
    - TYPE = 'OLET': 1 weld
    """
    # Read Excel file
    df_excel = pd.read_excel(excel_file)
    
    # Create a dictionary for quick lookup
    spre_lookup = {}
    for _, row in df_excel.iterrows():
        spre = row['SPRE']
        spre_lookup[spre] = {
            'P1_CONN': row['P1 CONN'] if pd.notna(row['P1 CONN']) else '',
            'P2_CONN': row['P2 CONN'] if pd.notna(row['P2 CONN']) else '',
            'TYPE': row['TYPE'] if pd.notna(row['TYPE']) else ''
        }
    
    # Process components
    result = []
    for comp in components:
        spre = comp['SPRE']
        
        if spre in spre_lookup:
            found = 'X'
            p1_conn = spre_lookup[spre]['P1_CONN']
            p2_conn = spre_lookup[spre]['P2_CONN']
            comp_type = spre_lookup[spre]['TYPE']
            
            # Calculate weld count
            weld_count = 0
            # Exclude WELD type components from weld counting
            if 'WELD' not in comp_type:
                if p1_conn == 'BWD':
                    weld_count += 1
                if p2_conn == 'BWD':
                    weld_count += 1
                if comp_type == 'OLET':
                    weld_count += 1
            
            # Mark as welded if BWD connection OR TYPE is OLET (but not WELD type)
            welded = 'X' if weld_count > 0 else ''
        else:
            found = ''
            p1_conn = ''
            p2_conn = ''
            comp_type = ''
            welded = ''
            weld_count = 0
        
        result.append({
            'Pipe': comp['Pipe'],
            'Branch': comp['Branch'],
            'Component_Type': comp['Component_Type'],
            'SPRE': spre,
            'Found': found,
            'P1_CONN': p1_conn,
            'P2_CONN': p2_conn,
            'TYPE': comp_type,
            'Welded': welded,
            'Weld_Count': weld_count
        })
    
    return result

def detect_components_at_branch_ends(file_path, excel_file, branch_positions, tolerance=5.0):
    """
    Detect welded components directly at branch heads (HPOS) or tails (TPOS).
    
    A component is considered at the head/tail if:
    - Distance from component center to HPOS/TPOS <= component_length/2 + tolerance
    
    Args:
        file_path: Path to E3D database listing file
        excel_file: Path to Excel file with component data
        branch_positions: List of branch positions from extract_branch_positions()
        tolerance: Distance tolerance in mm (default 5mm)
    
    Returns:
        Dictionary with:
        - 'components_at_ends': List of components at branch ends
        - 'stats': Statistics about components at ends vs HCON/TCON BWD
    """
    # Read Excel file and create lookup
    df_excel = pd.read_excel(excel_file)
    component_lookup = {}
    
    def calculate_component_length(comp_type, pbor, pbor1, form):
        """Calculate component length based on type."""
        if comp_type == 'ELBO':
            if form is not None and pbor > 0:
                try:
                    return float(form) * pbor
                except:
                    return None
            return None
        elif comp_type == 'TEE':
            return 0.90 * pbor if pbor > 0 else None
        elif comp_type == 'FLAN':
            return 0.4 * pbor if pbor > 0 else None
        elif comp_type == 'VALV':
            # Use PBOR if available, otherwise fallback to PBOR1
            if pbor > 0:
                return 0.5 * pbor
            elif pbor1 > 0:
                return 0.5 * pbor1
            else:
                return None
        elif comp_type == 'REDU':
            return 1.15 * pbor1 if pbor1 > 0 else None
        elif comp_type == 'CAP':
            return 0.0
        else:
            return None
    
    for _, row in df_excel.iterrows():
        spre = row['SPRE']
        p1_conn = row['P1 CONN'] if pd.notna(row['P1 CONN']) else ''
        p2_conn = row['P2 CONN'] if pd.notna(row['P2 CONN']) else ''
        comp_type = row['TYPE'] if pd.notna(row['TYPE']) else ''
        
        pbor_str = row['PBOR'] if pd.notna(row['PBOR']) else '0mm'
        try:
            pbor = float(str(pbor_str).replace('mm', ''))
        except:
            pbor = 0.0
        
        pbor1_str = row['PBOR1'] if pd.notna(row['PBOR1']) else '0mm'
        try:
            pbor1 = float(str(pbor1_str).replace('mm', ''))
        except:
            pbor1 = 0.0
        
        form_str = row['FORM'] if pd.notna(row['FORM']) else None
        try:
            form = float(form_str) if form_str is not None else None
        except:
            form = None
        
        comp_length = calculate_component_length(comp_type, pbor, pbor1, form)
        # Exclude WELD type components from welded classification
        is_welded = (p1_conn == 'BWD' or p2_conn == 'BWD' or comp_type == 'OLET') and 'WELD' not in comp_type
        
        component_lookup[spre] = {
            'welded': is_welded,
            'type': comp_type,
            'length': comp_length,
            'p1_conn': p1_conn,
            'p2_conn': p2_conn
        }
    
    # Create branch position lookup
    branch_lookup = {}
    for branch in branch_positions:
        branch_id = branch['Full_Branch_ID']
        branch_lookup[branch_id] = branch
    
    # Parse file to extract components with positions
    kks_pattern = r'\d[A-Z]{3}\d{2}BR\d{3}'
    branch_pattern = r'/B\d+'
    
    components_at_ends = []
    current_pipe = None
    current_branch = None
    current_component = None
    
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
                    spre_match = re.search(r'SPRE SPCOMPONENT\s+(\S+)', line)
                    if spre_match:
                        spre = spre_match.group(1)
                        if current_component is None:
                            current_component = {'spre': spre, 'pos': None}
                        else:
                            current_component['spre'] = spre
                        
                        # If we already have a position, process the component now
                        if current_component.get('pos') is not None:
                            spre = current_component['spre']
                            pos = current_component['pos']
                            if spre in component_lookup and component_lookup[spre]['welded']:
                                comp_data = component_lookup[spre]
                                comp_length = comp_data['length']
                                
                                if comp_length is not None:
                                    # Check against branch head and tail
                                    branch_id = current_pipe + current_branch
                                    if branch_id in branch_lookup:
                                        branch_info = branch_lookup[branch_id]
                                        hpos = (branch_info['HPOS_X'], branch_info['HPOS_Y'], branch_info['HPOS_Z'])
                                        tpos = (branch_info['TPOS_X'], branch_info['TPOS_Y'], branch_info['TPOS_Z'])
                                        
                                        # Calculate distances
                                        if all(coord is not None for coord in hpos):
                                            dist_to_head = ((pos[0] - hpos[0])**2 + (pos[1] - hpos[1])**2 + (pos[2] - hpos[2])**2)**0.5
                                            threshold_head = comp_length / 2.0 + tolerance
                                            
                                            if dist_to_head <= threshold_head:
                                                components_at_ends.append({
                                                    'KKS_Pipe': current_pipe,
                                                    'Branch': current_branch,
                                                    'Full_Branch_ID': branch_id,
                                                    'Component_Name': spre,
                                                    'Component_Type': comp_data['type'],
                                                    'Component_Length': round(comp_length, 2),
                                                    'Position': 'HEAD',
                                                    'HCON': branch_info['HCON'],
                                                    'Component_X': pos[0],
                                                    'Component_Y': pos[1],
                                                    'Component_Z': pos[2],
                                                    'Branch_End_X': hpos[0],
                                                    'Branch_End_Y': hpos[1],
                                                    'Branch_End_Z': hpos[2],
                                                    'Distance_mm': round(dist_to_head, 2),
                                                    'Threshold_mm': round(threshold_head, 2)
                                                })
                                        
                                        if all(coord is not None for coord in tpos):
                                            dist_to_tail = ((pos[0] - tpos[0])**2 + (pos[1] - tpos[1])**2 + (pos[2] - tpos[2])**2)**0.5
                                            threshold_tail = comp_length / 2.0 + tolerance
                                            
                                            if dist_to_tail <= threshold_tail:
                                                components_at_ends.append({
                                                    'KKS_Pipe': current_pipe,
                                                    'Branch': current_branch,
                                                    'Full_Branch_ID': branch_id,
                                                    'Component_Name': spre,
                                                    'Component_Type': comp_data['type'],
                                                    'Component_Length': round(comp_length, 2),
                                                    'Position': 'TAIL',
                                                    'TCON': branch_info['TCON'],
                                                    'Component_X': pos[0],
                                                    'Component_Y': pos[1],
                                                    'Component_Z': pos[2],
                                                    'Branch_End_X': tpos[0],
                                                    'Branch_End_Y': tpos[1],
                                                    'Branch_End_Z': tpos[2],
                                                    'Distance_mm': round(dist_to_tail, 2),
                                                    'Threshold_mm': round(threshold_tail, 2)
                                                })
                            # Reset for next component
                            current_component = None
                
                elif line.strip().startswith('POS '):
                    pos_match = re.search(r'POS X ([\d.-]+)mm Y ([\d.-]+)mm Z ([\d.-]+)mm', line)
                    if pos_match:
                        pos = (float(pos_match.group(1)), float(pos_match.group(2)), float(pos_match.group(3)))
                        if current_component is None:
                            # POS came before SPRE, initialize component with position
                            current_component = {'spre': None, 'pos': pos}
                        else:
                            current_component['pos'] = pos
                        
                        # If we already have both spre and pos, process the component
                        if current_component.get('spre') is not None:
                            spre = current_component['spre']
                            if spre in component_lookup and component_lookup[spre]['welded']:
                                comp_data = component_lookup[spre]
                                comp_length = comp_data['length']
                                
                                if comp_length is not None:
                                    # Check against branch head and tail
                                    branch_id = current_pipe + current_branch
                                    if branch_id in branch_lookup:
                                        branch_info = branch_lookup[branch_id]
                                        hpos = (branch_info['HPOS_X'], branch_info['HPOS_Y'], branch_info['HPOS_Z'])
                                        tpos = (branch_info['TPOS_X'], branch_info['TPOS_Y'], branch_info['TPOS_Z'])
                                        
                                        # Calculate distances
                                        if all(coord is not None for coord in hpos):
                                            dist_to_head = ((pos[0] - hpos[0])**2 + (pos[1] - hpos[1])**2 + (pos[2] - hpos[2])**2)**0.5
                                            threshold_head = comp_length / 2.0 + tolerance
                                            
                                            if dist_to_head <= threshold_head:
                                                components_at_ends.append({
                                                    'KKS_Pipe': current_pipe,
                                                    'Branch': current_branch,
                                                    'Full_Branch_ID': branch_id,
                                                    'Component_Name': spre,
                                                    'Component_Type': comp_data['type'],
                                                    'Component_Length': round(comp_length, 2),
                                                    'Position': 'HEAD',
                                                    'HCON': branch_info['HCON'],
                                                    'Component_X': pos[0],
                                                    'Component_Y': pos[1],
                                                    'Component_Z': pos[2],
                                                    'Branch_End_X': hpos[0],
                                                    'Branch_End_Y': hpos[1],
                                                    'Branch_End_Z': hpos[2],
                                                    'Distance_mm': round(dist_to_head, 2),
                                                    'Threshold_mm': round(threshold_head, 2)
                                                })
                                        
                                        if all(coord is not None for coord in tpos):
                                            dist_to_tail = ((pos[0] - tpos[0])**2 + (pos[1] - tpos[1])**2 + (pos[2] - tpos[2])**2)**0.5
                                            threshold_tail = comp_length / 2.0 + tolerance
                                            
                                            if dist_to_tail <= threshold_tail:
                                                components_at_ends.append({
                                                    'KKS_Pipe': current_pipe,
                                                    'Branch': current_branch,
                                                    'Full_Branch_ID': branch_id,
                                                    'Component_Name': spre,
                                                    'Component_Type': comp_data['type'],
                                                    'Component_Length': round(comp_length, 2),
                                                    'Position': 'TAIL',
                                                    'TCON': branch_info['TCON'],
                                                    'Component_X': pos[0],
                                                    'Component_Y': pos[1],
                                                    'Component_Z': pos[2],
                                                    'Branch_End_X': tpos[0],
                                                    'Branch_End_Y': tpos[1],
                                                    'Branch_End_Z': tpos[2],
                                                    'Distance_mm': round(dist_to_tail, 2),
                                                    'Threshold_mm': round(threshold_tail, 2)
                                                })
                            # Reset for next component
                            current_component = None
    
    # Calculate statistics
    total_hcon_bwd = sum(1 for b in branch_positions if b['HCON'] == 'BWD')
    total_tcon_bwd = sum(1 for b in branch_positions if b['TCON'] == 'BWD')
    components_at_head = sum(1 for c in components_at_ends if c['Position'] == 'HEAD')
    components_at_tail = sum(1 for c in components_at_ends if c['Position'] == 'TAIL')
    components_at_head_bwd = sum(1 for c in components_at_ends if c['Position'] == 'HEAD' and c.get('HCON') == 'BWD')
    components_at_tail_bwd = sum(1 for c in components_at_ends if c['Position'] == 'TAIL' and c.get('TCON') == 'BWD')
    
    stats = {
        'total_hcon_bwd': total_hcon_bwd,
        'total_tcon_bwd': total_tcon_bwd,
        'total_bwd_connections': total_hcon_bwd + total_tcon_bwd,
        'components_at_head': components_at_head,
        'components_at_tail': components_at_tail,
        'components_at_head_bwd': components_at_head_bwd,
        'components_at_tail_bwd': components_at_tail_bwd,
        'total_components_at_ends': len(components_at_ends)
    }
    
    return {
        'components_at_ends': components_at_ends,
        'stats': stats
    }

def extract_component_adjacency(file_path, excel_file, distance_threshold_close=50.0, distance_threshold_near=150.0):
    """
    Extract component pairs that are close to each other (potentially touching/welded).
    Only processes components that have welds (BWD connections or OLET type).
    Only checks specific component type pairs (order-independent):
    - ELBO-ELBO, ELBO-REDU, ELBO-FLAN, ELBO-TEE, ELBO-CAP, ELBO-VALV
    - TEE-REDU, TEE-TEE, TEE-CAP, TEE-FLAN
    - REDU-FLAN, REDU-CAP, REDU-REDU, REDU-VALV
    
    Uses PBOR (Pipe Bore) values to determine appropriate distance thresholds.
    Threshold logic:
    - Touching: PBOR + 50mm (accounts for component size + small gap)
    - Near: PBOR * 2 + 100mm (accounts for component size + moderate gap)
    
    Args:
        file_path: Path to E3D database listing file
        excel_file: Path to Excel file with component data (P1 CONN, P2 CONN, TYPE, PBOR)
        distance_threshold_close: Not used (kept for backward compatibility)
        distance_threshold_near: Not used (kept for backward compatibility)
    
    Returns:
        List of component pairs with distance information
    """
    # Define valid component type pairs (order-independent)
    VALID_PAIRS = {
        frozenset(['ELBO', 'ELBO']),
        frozenset(['ELBO', 'REDU']),
        frozenset(['ELBO', 'FLAN']),
        frozenset(['ELBO', 'TEE']),
        frozenset(['ELBO', 'CAP']),
        frozenset(['ELBO', 'VALV']),
        frozenset(['TEE', 'REDU']),
        frozenset(['TEE', 'TEE']),
        frozenset(['TEE', 'CAP']),
        frozenset(['TEE', 'FLAN']),
        frozenset(['REDU', 'FLAN']),
        frozenset(['REDU', 'CAP']),
        frozenset(['REDU', 'REDU']),
        frozenset(['REDU', 'VALV'])
    }
    
    # Read Excel file and create lookup for welded components with PBOR, PBOR1, FORM
    df_excel = pd.read_excel(excel_file)
    welded_components = {}  # Dictionary: SPRE -> {'welded': bool, 'pbor': float, 'pbor1': float, 'type': str, 'form': float, 'length': float}
    
    def calculate_component_length(comp_type, pbor, pbor1, form):
        """
        Calculate component length based on type and dimensions.
        
        Returns:
            float: Component length in mm, or None if cannot be calculated
        """
        if comp_type == 'ELBO':
            # ELBO: FORM * PBOR
            # FORM must be numeric (3 or 5)
            if form is not None and pbor > 0:
                try:
                    form_val = float(form)
                    return form_val * pbor
                except:
                    return None
            return None
        elif comp_type == 'TEE':
            # TEE: 0.90 * PBOR
            return 0.90 * pbor if pbor > 0 else None
        elif comp_type == 'FLAN':
            # FLAN: 0.4 * PBOR
            return 0.4 * pbor if pbor > 0 else None
        elif comp_type == 'VALV':
            # VALV: 0.5 * PBOR (use PBOR1 as fallback)
            if pbor > 0:
                return 0.5 * pbor
            elif pbor1 > 0:
                return 0.5 * pbor1
            else:
                return None
        elif comp_type == 'REDU':
            # REDU: 1.15 * PBOR1
            return 1.15 * pbor1 if pbor1 > 0 else None
        elif comp_type == 'CAP':
            # CAP: endpoint component, length = 0
            return 0.0
        else:
            return None
    
    for _, row in df_excel.iterrows():
        spre = row['SPRE']
        p1_conn = row['P1 CONN'] if pd.notna(row['P1 CONN']) else ''
        p2_conn = row['P2 CONN'] if pd.notna(row['P2 CONN']) else ''
        comp_type = row['TYPE'] if pd.notna(row['TYPE']) else ''
        
        # Parse PBOR (e.g., "100mm" -> 100.0)
        pbor_str = row['PBOR'] if pd.notna(row['PBOR']) else '0mm'
        try:
            pbor = float(str(pbor_str).replace('mm', ''))
        except:
            pbor = 0.0
        
        # Parse PBOR1 (e.g., "15mm" -> 15.0)
        pbor1_str = row['PBOR1'] if pd.notna(row['PBOR1']) else '0mm'
        try:
            pbor1 = float(str(pbor1_str).replace('mm', ''))
        except:
            pbor1 = 0.0
        
        # Parse FORM (could be numeric like '3', '5' or text like 'SWF/SWF')
        form_str = row['FORM'] if pd.notna(row['FORM']) else None
        try:
            form = float(form_str) if form_str is not None else None
        except:
            form = None
        
        # Calculate component length
        comp_length = calculate_component_length(comp_type, pbor, pbor1, form)
        
        # Include only welded components (BWD or OLET), exclude WELD type
        is_welded = (p1_conn == 'BWD' or p2_conn == 'BWD' or comp_type == 'OLET') and 'WELD' not in comp_type
        welded_components[spre] = {
            'welded': is_welded, 
            'pbor': pbor, 
            'pbor1': pbor1,
            'type': comp_type,
            'form': form,
            'length': comp_length
        }
    
    kks_pattern = r'\d[A-Z]{3}\d{2}BR\d{3}'
    branch_pattern = r'/B\d+'
    
    component_pairs = []
    current_pipe = None
    current_branch = None
    branch_components = []  # List of (component_type, spre, position, pbor) for current branch - only welded ones
    current_component = None  # Track the component being built
    
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            # Check if line contains "NEW PIPE"
            if 'NEW PIPE' in line:
                match = re.search(kks_pattern, line)
                if match:
                    # Process previous branch's components before resetting
                    if current_pipe and current_branch and len(branch_components) > 1:
                        # Analyze pairs of consecutive components
                        for i in range(len(branch_components) - 1):
                            comp1 = branch_components[i]
                            comp2 = branch_components[i + 1]
                            
                            # Check if this is a valid component type pair
                            type_pair = frozenset([comp1['type'], comp2['type']])
                            if type_pair not in VALID_PAIRS:
                                continue  # Skip this pair
                            
                            # Skip if either component has no valid length
                            if comp1['length'] is None or comp2['length'] is None:
                                continue
                            
                            # Calculate 3D distance between components
                            distance = ((comp1['pos'][0] - comp2['pos'][0])**2 + 
                                       (comp1['pos'][1] - comp2['pos'][1])**2 + 
                                       (comp1['pos'][2] - comp2['pos'][2])**2)**0.5
                            
                            # Calculate expected touching distance based on component types
                            if comp1['type'] == 'ELBO' and comp2['type'] == 'ELBO':
                                # ELBO to ELBO: sum of both lengths
                                expected_distance = comp1['length'] + comp2['length']
                            elif comp1['type'] == 'ELBO':
                                # ELBO to other: just ELBO length
                                expected_distance = comp1['length']
                            elif comp2['type'] == 'ELBO':
                                # Other to ELBO: just ELBO length
                                expected_distance = comp2['length']
                            elif comp1['type'] == 'CAP' or comp2['type'] == 'CAP':
                                # If CAP is involved, use the other component's length
                                expected_distance = comp1['length'] if comp2['type'] == 'CAP' else comp2['length']
                            else:
                                # Other combinations: sum of both lengths
                                expected_distance = comp1['length'] + comp2['length']
                            
                            # Apply margins: touching ±10%, near ±50%, separated ±100%
                            threshold_touching = expected_distance * 1.10  # +10% margin
                            threshold_near = expected_distance * 1.50  # +50% margin
                            threshold_separated = expected_distance * 2.00  # +100% margin
                            
                            # Determine relationship
                            if distance <= threshold_touching:
                                relationship = 'Touching'
                            elif distance <= threshold_near:
                                relationship = 'Near'
                            else:
                                relationship = 'Separated'
                            
                            component_pairs.append({
                                'KKS_Pipe': current_pipe,
                                'Branch': current_branch,
                                'Component_1_Name': comp1['spre'],
                                'Component_1_Type': comp1['type'],
                                'Component_1_PBOR': comp1['pbor'],
                                'Component_1_Length': round(comp1['length'], 2),
                                'Component_1_X': comp1['pos'][0],
                                'Component_1_Y': comp1['pos'][1],
                                'Component_1_Z': comp1['pos'][2],
                                'Component_2_Name': comp2['spre'],
                                'Component_2_Type': comp2['type'],
                                'Component_2_PBOR': comp2['pbor'],
                                'Component_2_Length': round(comp2['length'], 2),
                                'Component_2_X': comp2['pos'][0],
                                'Component_2_Y': comp2['pos'][1],
                                'Component_2_Z': comp2['pos'][2],
                                'Distance_mm': round(distance, 2),
                                'Expected_Distance_mm': round(expected_distance, 2),
                                'Threshold_Touching': round(threshold_touching, 2),
                                'Threshold_Near': round(threshold_near, 2),
                                'Relationship': relationship
                            })
                    
                    current_pipe = match.group()
                    current_branch = None
                    branch_components = []
            
            # Check if line contains "NEW BRANCH"
            elif 'NEW BRANCH' in line and current_pipe:
                # Process previous branch's components before resetting
                if current_branch and len(branch_components) > 1:
                    for i in range(len(branch_components) - 1):
                        comp1 = branch_components[i]
                        comp2 = branch_components[i + 1]
                        
                        # Check if this is a valid component type pair
                        type_pair = frozenset([comp1['type'], comp2['type']])
                        if type_pair not in VALID_PAIRS:
                            continue  # Skip this pair
                        
                        # Skip if either component has no valid length
                        if comp1['length'] is None or comp2['length'] is None:
                            continue
                        
                        distance = ((comp1['pos'][0] - comp2['pos'][0])**2 + 
                                   (comp1['pos'][1] - comp2['pos'][1])**2 + 
                                   (comp1['pos'][2] - comp2['pos'][2])**2)**0.5
                        
                        # Calculate expected touching distance based on component types
                        if comp1['type'] == 'ELBO' and comp2['type'] == 'ELBO':
                            expected_distance = comp1['length'] + comp2['length']
                        elif comp1['type'] == 'ELBO':
                            expected_distance = comp1['length']
                        elif comp2['type'] == 'ELBO':
                            expected_distance = comp2['length']
                        elif comp1['type'] == 'CAP' or comp2['type'] == 'CAP':
                            expected_distance = comp1['length'] if comp2['type'] == 'CAP' else comp2['length']
                        else:
                            expected_distance = comp1['length'] + comp2['length']
                        
                        threshold_touching = expected_distance * 1.10
                        threshold_near = expected_distance * 1.50
                        
                        if distance <= threshold_touching:
                            relationship = 'Touching'
                        elif distance <= threshold_near:
                            relationship = 'Near'
                        else:
                            relationship = 'Separated'
                        
                        component_pairs.append({
                            'KKS_Pipe': current_pipe,
                            'Branch': current_branch,
                            'Component_1_Name': comp1['spre'],
                            'Component_1_Type': comp1['type'],
                            'Component_1_PBOR': comp1['pbor'],
                            'Component_1_Length': round(comp1['length'], 2),
                            'Component_1_X': comp1['pos'][0],
                            'Component_1_Y': comp1['pos'][1],
                            'Component_1_Z': comp1['pos'][2],
                            'Component_2_Name': comp2['spre'],
                            'Component_2_Type': comp2['type'],
                            'Component_2_PBOR': comp2['pbor'],
                            'Component_2_Length': round(comp2['length'], 2),
                            'Component_2_X': comp2['pos'][0],
                            'Component_2_Y': comp2['pos'][1],
                            'Component_2_Z': comp2['pos'][2],
                            'Distance_mm': round(distance, 2),
                            'Expected_Distance_mm': round(expected_distance, 2),
                            'Threshold_Touching': round(threshold_touching, 2),
                            'Threshold_Near': round(threshold_near, 2),
                            'Relationship': relationship
                        })
                
                full_match = re.search(kks_pattern + branch_pattern, line)
                if full_match:
                    full_branch = full_match.group()
                    branch_match = re.search(branch_pattern, full_branch)
                    if branch_match:
                        current_branch = branch_match.group()
                        branch_components = []
                        current_component = None
            
            # Check if line starts with "NEW " followed by a component type
            elif line.strip().startswith('NEW ') and current_branch:
                parts = line.strip().split()
                if len(parts) >= 2:
                    component_type = parts[1]
                    # Exclude ATTACHMENT and structural elements  
                    # Include only piping components: FLANGE, ELBOW, TEE, REDUCER, VALVE, etc.
                    if component_type not in ['BRANCH', 'PIPE', 'ZONE', 'SITE', 'STRUCTURE', 
                                              'SUBSTRUCTURE', 'CYLINDER', 'CTORUS', 'ATTACHMENT']:
                        current_component = {
                            'type': component_type, 
                            'spre': None, 
                            'pos': None, 
                            'pbor': 0.0,
                            'pipe': current_pipe,
                            'branch': current_branch
                        }
            
            # Extract SPRE for current component
            elif 'SPRE SPCOMPONENT' in line and current_component:
                if current_component['spre'] is None:
                    spre_match = re.search(r'SPRE SPCOMPONENT\s+(\S+)', line)
                    if spre_match:
                        current_component['spre'] = spre_match.group(1)
                        # Component is complete if both spre and pos are set
                        # Only add if it's a welded component
                        if current_component['pos'] is not None:
                            spre = current_component['spre']
                            if spre in welded_components and welded_components[spre]['welded']:
                                current_component['pbor'] = welded_components[spre]['pbor']
                                current_component['pbor1'] = welded_components[spre]['pbor1']
                                current_component['type'] = welded_components[spre]['type']  # Use type from Excel
                                current_component['length'] = welded_components[spre]['length']
                                branch_components.append(current_component)
                            current_component = None
            
            # Extract position for current component
            elif line.strip().startswith('POS ') and current_component:
                if current_component['pos'] is None:
                    pos_match = re.search(r'POS X ([\d.-]+)mm Y ([\d.-]+)mm Z ([\d.-]+)mm', line)
                    if pos_match:
                        current_component['pos'] = (
                            float(pos_match.group(1)),
                            float(pos_match.group(2)),
                            float(pos_match.group(3))
                        )
                        # Component is complete if both spre and pos are set
                        # Only add if it's a welded component
                        if current_component['spre'] is not None:
                            spre = current_component['spre']
                            if spre in welded_components and welded_components[spre]['welded']:
                                current_component['pbor'] = welded_components[spre]['pbor']
                                current_component['pbor1'] = welded_components[spre]['pbor1']
                                current_component['type'] = welded_components[spre]['type']  # Use type from Excel
                                current_component['length'] = welded_components[spre]['length']
                                branch_components.append(current_component)
                            current_component = None
    
    # Don't forget the last branch
    if current_branch and len(branch_components) > 1:
        for i in range(len(branch_components) - 1):
            comp1 = branch_components[i]
            comp2 = branch_components[i + 1]
            
            # Check if this is a valid component type pair
            type_pair = frozenset([comp1['type'], comp2['type']])
            if type_pair not in VALID_PAIRS:
                continue  # Skip this pair
            
            # Skip if either component has no valid length
            if comp1['length'] is None or comp2['length'] is None:
                continue
            
            distance = ((comp1['pos'][0] - comp2['pos'][0])**2 + 
                       (comp1['pos'][1] - comp2['pos'][1])**2 + 
                       (comp1['pos'][2] - comp2['pos'][2])**2)**0.5
            
            # Calculate expected touching distance based on component types
            if comp1['type'] == 'ELBO' and comp2['type'] == 'ELBO':
                expected_distance = comp1['length'] + comp2['length']
            elif comp1['type'] == 'ELBO':
                expected_distance = comp1['length']
            elif comp2['type'] == 'ELBO':
                expected_distance = comp2['length']
            elif comp1['type'] == 'CAP' or comp2['type'] == 'CAP':
                expected_distance = comp1['length'] if comp2['type'] == 'CAP' else comp2['length']
            else:
                expected_distance = comp1['length'] + comp2['length']
            
            threshold_touching = expected_distance * 1.10
            threshold_near = expected_distance * 1.50
            
            if distance <= threshold_touching:
                relationship = 'Touching'
            elif distance <= threshold_near:
                relationship = 'Near'
            else:
                relationship = 'Separated'
            
            component_pairs.append({
                'KKS_Pipe': current_pipe,
                'Branch': current_branch,
                'Component_1_Name': comp1['spre'],
                'Component_1_Type': comp1['type'],
                'Component_1_PBOR': comp1['pbor'],
                'Component_1_Length': round(comp1['length'], 2),
                'Component_1_X': comp1['pos'][0],
                'Component_1_Y': comp1['pos'][1],
                'Component_1_Z': comp1['pos'][2],
                'Component_2_Name': comp2['spre'],
                'Component_2_Type': comp2['type'],
                'Component_2_PBOR': comp2['pbor'],
                'Component_2_Length': round(comp2['length'], 2),
                'Component_2_X': comp2['pos'][0],
                'Component_2_Y': comp2['pos'][1],
                'Component_2_Z': comp2['pos'][2],
                'Distance_mm': round(distance, 2),
                'Expected_Distance_mm': round(expected_distance, 2),
                'Threshold_Touching': round(threshold_touching, 2),
                'Threshold_Near': round(threshold_near, 2),
                'Relationship': relationship
            })
    
    return component_pairs

def save_to_csv(data, output_file):
    """
    Save components data to a CSV file.
    """
    if not data:
        print("No data to save.")
        return
    
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Pipe', 'Branch', 'Component_Type', 'SPRE', 'Found', 'P1_CONN', 'P2_CONN', 'TYPE', 'Welded', 'Weld_Count']
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
    output_branches_csv = Path(__file__).parent / "TBY" / "branch_connections.csv"
    output_branch_coord_csv = Path(__file__).parent / "TBY" / "connected_branches.csv"
    output_adjacency_csv = Path(__file__).parent / "TBY" / "component_adjacency.csv"
    
    all_components = []
    all_branches = []
    all_branch_positions = []
    all_component_pairs = []
    
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
        
        # Extract branch connection information
        branches = extract_branch_connections(txt_file)
        print(f"  Extracted {len(branches)} branch connections")
        all_branches.extend(branches)
        
        # Extract branch positions
        branch_positions = extract_branch_positions(txt_file)
        print(f"  Extracted {len(branch_positions)} branch positions")
        all_branch_positions.extend(branch_positions)
        
        # Extract component adjacency (touching components) - only welded components
        component_pairs = extract_component_adjacency(txt_file, excel_file, distance_threshold_close=50.0, distance_threshold_near=150.0)
        print(f"  Extracted {len(component_pairs)} component pairs")
        all_component_pairs.extend(component_pairs)
    
    print(f"\nTotal components from all files: {len(all_components)}")
    print(f"Total branches from all files: {len(all_branches)}")
    print(f"Total branch positions: {len(all_branch_positions)}")
    print(f"Total component pairs: {len(all_component_pairs)}")
    
    # Detect components at branch ends
    print(f"\nDetecting components at branch ends...")
    all_components_at_ends = []
    for txt_file in txt_files:
        result = detect_components_at_branch_ends(txt_file, excel_file, all_branch_positions, tolerance=100.0)
        all_components_at_ends.extend(result['components_at_ends'])
    
    # Get combined stats
    end_detection_stats = detect_components_at_branch_ends(txt_files[0], excel_file, all_branch_positions, tolerance=100.0)['stats']
    
    # Find connected branches based on coordinates
    print(f"\nFinding connected branches based on coordinates...")
    connected_branches = find_connected_branches(all_branch_positions, tolerance_tight=5.0, tolerance_loose=150.0)
    print(f"Found {len(connected_branches)} branch connections")
    
    # Count match types
    xy_count = sum(1 for c in connected_branches if c['Match_Type'] == 'XY_tight_Z_loose')
    xz_count = sum(1 for c in connected_branches if c['Match_Type'] == 'XZ_tight_Y_loose')
    yz_count = sum(1 for c in connected_branches if c['Match_Type'] == 'YZ_tight_X_loose')
    print(f"  - XY tight (<=5mm), Z loose (<=150mm): {xy_count}")
    print(f"  - XZ tight (<=5mm), Y loose (<=150mm): {xz_count}")
    print(f"  - YZ tight (<=5mm), X loose (<=150mm): {yz_count}")
    
    # Save connected branches to CSV FIRST (before Excel read which might fail)
    print(f"\nSaving connected branches to: {output_branch_coord_csv}")
    with open(output_branch_coord_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Branch_A', 'Branch_A_TCON', 'Branch_B', 'Branch_B_HCON',
                     'Branch_A_Pipe', 'Branch_A_Branch',
                     'Branch_B_Pipe', 'Branch_B_Branch', 
                     'Connection_X', 'Connection_Y', 'Connection_Z', 
                     'Distance_mm', 'Offset_X_mm', 'Offset_Y_mm', 'Offset_Z_mm', 
                     'Max_Tight_Offset_mm', 'Loose_Offset_mm', 'Accuracy_Percent', 'Match_Type']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in connected_branches:
            writer.writerow(row)
    print(f"Connected branches written: {len(connected_branches)}")
    
    print(f"\nReading Excel file: {excel_file}")
    
    # Lookup and merge with Excel data
    result = lookup_and_merge_with_excel(all_components, excel_file)
    
    # Save components to CSV
    save_to_csv(result, output_csv)
    
    # Save branches to CSV
    print(f"\nSaving branch connections to: {output_branches_csv}")
    with open(output_branches_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Pipe', 'Branch', 'HCON', 'TCON', 'HSTU', 'First_Component', 'Last_Component', 'Head_Pipe_Length_mm', 'Tail_Pipe_Length_mm']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        for row in all_branches:
            writer.writerow(row)
    print(f"Branch connections written: {len(all_branches)}")
    
    # Save components at branch ends to CSV
    output_ends_csv = Path(__file__).parent / "TBY" / 'components_at_branch_ends.csv'
    print(f"\nSaving components at branch ends to: {output_ends_csv}")
    with open(output_ends_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['KKS_Pipe', 'Branch', 'Full_Branch_ID', 'Component_Name', 'Component_Type', 
                     'Component_Length', 'Position', 'HCON', 'TCON',
                     'Component_X', 'Component_Y', 'Component_Z',
                     'Branch_End_X', 'Branch_End_Y', 'Branch_End_Z',
                     'Distance_mm', 'Threshold_mm']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        for row in all_components_at_ends:
            writer.writerow(row)
    print(f"Components at branch ends written: {len(all_components_at_ends)}")
    
    # Create BWD connections report
    output_bwd_report_csv = Path(__file__).parent / "TBY" / 'bwd_connections_report.csv'
    print(f"\nGenerating BWD connections report...")
    
    # Create a lookup for components at branch ends
    components_at_head_lookup = {}  # branch_id -> list of components
    components_at_tail_lookup = {}  # branch_id -> list of components
    
    for comp in all_components_at_ends:
        branch_id = comp['Full_Branch_ID']
        if comp['Position'] == 'HEAD':
            if branch_id not in components_at_head_lookup:
                components_at_head_lookup[branch_id] = []
            components_at_head_lookup[branch_id].append({
                'name': comp['Component_Name'],
                'type': comp['Component_Type'],
                'distance': comp['Distance_mm']
            })
        elif comp['Position'] == 'TAIL':
            if branch_id not in components_at_tail_lookup:
                components_at_tail_lookup[branch_id] = []
            components_at_tail_lookup[branch_id].append({
                'name': comp['Component_Name'],
                'type': comp['Component_Type'],
                'distance': comp['Distance_mm']
            })
    
    # Build BWD connection report
    bwd_connections_report = []
    
    for branch in all_branch_positions:
        branch_id = branch['Full_Branch_ID']
        pipe = branch['Pipe']
        branch_name = branch['Branch']
        hcon = branch['HCON']
        tcon = branch['TCON']
        
        # Check if HCON is BWD
        if hcon == 'BWD':
            head_components = components_at_head_lookup.get(branch_id, [])
            if head_components:
                # Sort by distance and get closest
                head_components.sort(key=lambda x: x['distance'])
                closest = head_components[0]
                has_component = 'Yes'
                component_name = closest['name']
                component_type = closest['type']
                component_distance = closest['distance']
            else:
                has_component = 'No'
                component_name = ''
                component_type = ''
                component_distance = None
            
            bwd_connections_report.append({
                'KKS_Pipe': pipe,
                'Branch': branch_name,
                'Full_Branch_ID': branch_id,
                'End_Type': 'HEAD',
                'Connection_Type': hcon,
                'Has_Component_At_End': has_component,
                'Component_Name': component_name,
                'Component_Type': component_type,
                'Distance_To_End_mm': component_distance
            })
        
        # Check if TCON is BWD
        if tcon == 'BWD':
            tail_components = components_at_tail_lookup.get(branch_id, [])
            if tail_components:
                # Sort by distance and get closest
                tail_components.sort(key=lambda x: x['distance'])
                closest = tail_components[0]
                has_component = 'Yes'
                component_name = closest['name']
                component_type = closest['type']
                component_distance = closest['distance']
            else:
                has_component = 'No'
                component_name = ''
                component_type = ''
                component_distance = None
            
            bwd_connections_report.append({
                'KKS_Pipe': pipe,
                'Branch': branch_name,
                'Full_Branch_ID': branch_id,
                'End_Type': 'TAIL',
                'Connection_Type': tcon,
                'Has_Component_At_End': has_component,
                'Component_Name': component_name,
                'Component_Type': component_type,
                'Distance_To_End_mm': component_distance
            })
    
    # Save BWD connections report to CSV
    print(f"Saving BWD connections report to: {output_bwd_report_csv}")
    with open(output_bwd_report_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['KKS_Pipe', 'Branch', 'Full_Branch_ID', 'End_Type', 'Connection_Type',
                     'Has_Component_At_End', 'Component_Name', 'Component_Type', 'Distance_To_End_mm']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in bwd_connections_report:
            writer.writerow(row)
    print(f"BWD connections report written: {len(bwd_connections_report)} entries")
    
    # Count summary for BWD report
    bwd_with_component = sum(1 for r in bwd_connections_report if r['Has_Component_At_End'] == 'Yes')
    bwd_without_component = sum(1 for r in bwd_connections_report if r['Has_Component_At_End'] == 'No')
    print(f"  - BWD connections WITH component at end: {bwd_with_component}")
    print(f"  - BWD connections WITHOUT component at end: {bwd_without_component}")
    
    # Save component adjacency to CSV
    print(f"\nSaving component adjacency to: {output_adjacency_csv}")
    with open(output_adjacency_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['KKS_Pipe', 'Branch', 
                     'Component_1_Name', 'Component_1_Type', 'Component_1_PBOR', 'Component_1_Length',
                     'Component_1_X', 'Component_1_Y', 'Component_1_Z',
                     'Component_2_Name', 'Component_2_Type', 'Component_2_PBOR', 'Component_2_Length',
                     'Component_2_X', 'Component_2_Y', 'Component_2_Z',
                     'Distance_mm', 'Expected_Distance_mm', 'Threshold_Touching', 'Threshold_Near', 'Relationship']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in all_component_pairs:
            writer.writerow(row)
    print(f"Component pairs written: {len(all_component_pairs)}")
    
    # Count relationships
    touching_count_all = sum(1 for p in all_component_pairs if p['Relationship'] == 'Touching')
    near_count = sum(1 for p in all_component_pairs if p['Relationship'] == 'Near')
    separated_count = sum(1 for p in all_component_pairs if p['Relationship'] == 'Separated')
    
    # Filter touching pairs: exclude OLETs and FLANGE-FLANGE pairs
    touching_count_filtered = sum(1 for p in all_component_pairs 
                                  if p['Relationship'] == 'Touching'
                                  and p['Component_1_Type'] != 'OLET' 
                                  and p['Component_2_Type'] != 'OLET'
                                  and not (p['Component_1_Type'] == 'FLANGE' and p['Component_2_Type'] == 'FLANGE'))
    
    print(f"  - Touching (PBOR-based threshold): {touching_count_all}")
    print(f"    * Excluding OLETs and FLANGE-FLANGE pairs: {touching_count_filtered}")
    print(f"  - Near (PBOR-based threshold): {near_count}")
    print(f"  - Separated (>threshold): {separated_count}")
    
    # Display summary statistics
    found_count = sum(1 for r in result if r['Found'] == 'X')
    welded_count = sum(1 for r in result if r['Welded'] == 'X')
    total_welds = sum(r['Weld_Count'] for r in result)
    
    # Count BWD branch ends with and without components
    total_bwd_branch_ends = end_detection_stats['total_bwd_connections']
    components_at_bwd_ends = end_detection_stats['components_at_head_bwd'] + end_detection_stats['components_at_tail_bwd']
    
    # Calculate final weld count
    # Start with component welds
    # Add BWD branch ends (each BWD connection needs a weld)
    #   Note: connected_branches are already included in BWD branch ends count, so we don't add them separately
    # Subtract touching component pairs (already counted in component welds)
    # Subtract components at BWD ends (already counted in component welds)
    # Exclude OLETs and FLANGE-FLANGE pairs from touching count
    final_weld_count = total_welds + total_bwd_branch_ends - touching_count_filtered - components_at_bwd_ends
    
    print(f"\nSummary:")
    print(f"Total components: {len(result)}")
    print(f"Found in Excel: {found_count}")
    print(f"Not found in Excel: {len(result) - found_count}")
    print(f"Welded components (BWD or OLET): {welded_count}")
    print(f"Total weld count (from components): {total_welds}")
    print(f"Connected branches (branch-to-branch welds): {len(connected_branches)}")
    print(f"Touching component pairs (component-to-component welds, excl. OLETs & FLANGE-FLANGE): {touching_count_filtered}")
    print(f"BWD branch ends (require welds): {total_bwd_branch_ends}")
    print(f"  Note: Connected branches ({len(connected_branches)}) are already included in BWD branch ends")
    print(f"Components at BWD branch ends (already in component welds): {components_at_bwd_ends}")
    print(f"Final weld count: {total_welds} + {total_bwd_branch_ends} - {touching_count_filtered} - {components_at_bwd_ends} = {final_weld_count}")
    
    print(f"\n{'='*80}")
    print(f"BRANCH END ANALYSIS:")
    print(f"{'='*80}")
    print(f"Total HCON BWD connections in project: {end_detection_stats['total_hcon_bwd']}")
    print(f"Total TCON BWD connections in project: {end_detection_stats['total_tcon_bwd']}")
    print(f"Total BWD connections (HCON + TCON): {end_detection_stats['total_bwd_connections']}")
    print(f"\nComponents directly at branch HEADS: {end_detection_stats['components_at_head']}")
    print(f"  - At HEADS with HCON=BWD: {end_detection_stats['components_at_head_bwd']}")
    print(f"Components directly at branch TAILS: {end_detection_stats['components_at_tail']}")
    print(f"  - At TAILS with TCON=BWD: {end_detection_stats['components_at_tail_bwd']}")
    print(f"\nTotal components at branch ends (HEAD or TAIL): {end_detection_stats['total_components_at_ends']}")
    print(f"{'='*80}")


