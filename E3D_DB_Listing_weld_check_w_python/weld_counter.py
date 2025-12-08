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
                            'HPOS_X': None,
                            'HPOS_Y': None,
                            'HPOS_Z': None,
                            'TPOS_X': None,
                            'TPOS_Y': None,
                            'TPOS_Z': None
                        }
                        in_branch = True
            
            # Extract HPOS and TPOS coordinates
            elif in_branch and current_branch_info:
                if line.strip().startswith('HPOS '):
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
                        # Once we have TPOS, we can consider branch parsing complete
                        in_branch = False
    
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
                        'Branch_B': branch2['Full_Branch_ID'],
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
            if p1_conn == 'BWD':
                weld_count += 1
            if p2_conn == 'BWD':
                weld_count += 1
            if comp_type == 'OLET':
                weld_count += 1
            
            # Mark as welded if BWD connection OR TYPE is OLET
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

def extract_component_adjacency(file_path, excel_file, distance_threshold_close=50.0, distance_threshold_near=150.0):
    """
    Extract component pairs that are close to each other (potentially touching/welded).
    Only processes components that have welds (BWD connections or OLET type).
    
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
    # Read Excel file and create lookup for welded components with PBOR
    df_excel = pd.read_excel(excel_file)
    welded_components = {}  # Dictionary: SPRE -> {'welded': bool, 'pbor': float}
    
    for _, row in df_excel.iterrows():
        spre = row['SPRE']
        p1_conn = row['P1 CONN'] if pd.notna(row['P1 CONN']) else ''
        p2_conn = row['P2 CONN'] if pd.notna(row['P2 CONN']) else ''
        comp_type = row['TYPE'] if pd.notna(row['TYPE']) else ''
        pbor_str = row['PBOR'] if pd.notna(row['PBOR']) else '0mm'
        
        # Parse PBOR (e.g., "100mm" -> 100.0)
        try:
            pbor = float(pbor_str.replace('mm', ''))
        except:
            pbor = 0.0
        
        # Include only welded components (BWD or OLET)
        is_welded = (p1_conn == 'BWD' or p2_conn == 'BWD' or comp_type == 'OLET')
        welded_components[spre] = {'welded': is_welded, 'pbor': pbor}
    
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
                            
                            # Calculate 3D distance between components
                            distance = ((comp1['pos'][0] - comp2['pos'][0])**2 + 
                                       (comp1['pos'][1] - comp2['pos'][1])**2 + 
                                       (comp1['pos'][2] - comp2['pos'][2])**2)**0.5
                            
                            # Use PBOR-based thresholds (average of both components)
                            avg_pbor = (comp1['pbor'] + comp2['pbor']) / 2.0
                            threshold_touching = avg_pbor + 50.0  # PBOR + 50mm
                            threshold_near = avg_pbor * 2 + 100.0  # PBOR * 2 + 100mm
                            
                            # Determine relationship
                            if distance <= threshold_touching:
                                relationship = 'Touching'
                            elif distance <= threshold_near:
                                relationship = 'Near'
                            else:
                                relationship = 'Separated'
                            
                            component_pairs.append({
                                'Pipe': current_pipe,
                                'Branch': current_branch,
                                'Component_1_Type': comp1['type'],
                                'Component_1_SPRE': comp1['spre'],
                                'Component_1_PBOR': comp1['pbor'],
                                'Component_1_X': comp1['pos'][0],
                                'Component_1_Y': comp1['pos'][1],
                                'Component_1_Z': comp1['pos'][2],
                                'Component_2_Type': comp2['type'],
                                'Component_2_SPRE': comp2['spre'],
                                'Component_2_PBOR': comp2['pbor'],
                                'Component_2_X': comp2['pos'][0],
                                'Component_2_Y': comp2['pos'][1],
                                'Component_2_Z': comp2['pos'][2],
                                'Distance_mm': round(distance, 2),
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
                        
                        distance = ((comp1['pos'][0] - comp2['pos'][0])**2 + 
                                   (comp1['pos'][1] - comp2['pos'][1])**2 + 
                                   (comp1['pos'][2] - comp2['pos'][2])**2)**0.5
                        
                        # Use PBOR-based thresholds (average of both components)
                        avg_pbor = (comp1['pbor'] + comp2['pbor']) / 2.0
                        threshold_touching = avg_pbor + 50.0  # PBOR + 50mm
                        threshold_near = avg_pbor * 2 + 100.0  # PBOR * 2 + 100mm
                        
                        if distance <= threshold_touching:
                            relationship = 'Touching'
                        elif distance <= threshold_near:
                            relationship = 'Near'
                        else:
                            relationship = 'Separated'
                        
                        component_pairs.append({
                            'Pipe': current_pipe,
                            'Branch': current_branch,
                            'Component_1_Type': comp1['type'],
                            'Component_1_SPRE': comp1['spre'],
                            'Component_1_PBOR': comp1['pbor'],
                            'Component_1_X': comp1['pos'][0],
                            'Component_1_Y': comp1['pos'][1],
                            'Component_1_Z': comp1['pos'][2],
                            'Component_2_Type': comp2['type'],
                            'Component_2_SPRE': comp2['spre'],
                            'Component_2_PBOR': comp2['pbor'],
                            'Component_2_X': comp2['pos'][0],
                            'Component_2_Y': comp2['pos'][1],
                            'Component_2_Z': comp2['pos'][2],
                            'Distance_mm': round(distance, 2),
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
                        current_component = {'type': component_type, 'spre': None, 'pos': None, 'pbor': 0.0}
            
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
                                branch_components.append(current_component)
                            current_component = None
    
    # Don't forget the last branch
    if current_branch and len(branch_components) > 1:
        for i in range(len(branch_components) - 1):
            comp1 = branch_components[i]
            comp2 = branch_components[i + 1]
            
            distance = ((comp1['pos'][0] - comp2['pos'][0])**2 + 
                       (comp1['pos'][1] - comp2['pos'][1])**2 + 
                       (comp1['pos'][2] - comp2['pos'][2])**2)**0.5
            
            # Use PBOR-based thresholds (average of both components)
            avg_pbor = (comp1['pbor'] + comp2['pbor']) / 2.0
            threshold_touching = avg_pbor + 50.0  # PBOR + 50mm
            threshold_near = avg_pbor * 2 + 100.0  # PBOR * 2 + 100mm
            
            if distance <= threshold_touching:
                relationship = 'Touching'
            elif distance <= threshold_near:
                relationship = 'Near'
            else:
                relationship = 'Separated'
            
            component_pairs.append({
                'Pipe': current_pipe,
                'Branch': current_branch,
                'Component_1_Type': comp1['type'],
                'Component_1_SPRE': comp1['spre'],
                'Component_1_PBOR': comp1['pbor'],
                'Component_1_X': comp1['pos'][0],
                'Component_1_Y': comp1['pos'][1],
                'Component_1_Z': comp1['pos'][2],
                'Component_2_Type': comp2['type'],
                'Component_2_SPRE': comp2['spre'],
                'Component_2_PBOR': comp2['pbor'],
                'Component_2_X': comp2['pos'][0],
                'Component_2_Y': comp2['pos'][1],
                'Component_2_Z': comp2['pos'][2],
                'Distance_mm': round(distance, 2),
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
    
    # Find connected branches based on coordinates
    print(f"\nFinding connected branches based on coordinates...")
    connected_branches = find_connected_branches(all_branch_positions, tolerance_tight=5.0, tolerance_loose=150.0)
    print(f"Found {len(connected_branches)} branch connections")
    
    # Count match types
    xy_count = sum(1 for c in connected_branches if c['Match_Type'] == 'XY_tight_Z_loose')
    xz_count = sum(1 for c in connected_branches if c['Match_Type'] == 'XZ_tight_Y_loose')
    yz_count = sum(1 for c in connected_branches if c['Match_Type'] == 'YZ_tight_X_loose')
    print(f"  - XY tight (≤5mm), Z loose (≤150mm): {xy_count}")
    print(f"  - XZ tight (≤5mm), Y loose (≤150mm): {xz_count}")
    print(f"  - YZ tight (≤5mm), X loose (≤150mm): {yz_count}")
    
    # Save connected branches to CSV FIRST (before Excel read which might fail)
    print(f"\nSaving connected branches to: {output_branch_coord_csv}")
    with open(output_branch_coord_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Branch_A', 'Branch_B', 'Branch_A_Pipe', 'Branch_A_Branch',
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
    
    # Save component adjacency to CSV
    print(f"\nSaving component adjacency to: {output_adjacency_csv}")
    with open(output_adjacency_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Pipe', 'Branch', 'Component_1_Type', 'Component_1_SPRE', 'Component_1_PBOR',
                     'Component_1_X', 'Component_1_Y', 'Component_1_Z',
                     'Component_2_Type', 'Component_2_SPRE', 'Component_2_PBOR',
                     'Component_2_X', 'Component_2_Y', 'Component_2_Z',
                     'Distance_mm', 'Relationship']
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
    
    # Calculate final weld count
    # Add connected branches (branch-to-branch welds)
    # Subtract touching component pairs (component-to-component welds that replace pipe-to-component welds)
    # Exclude OLETs and FLANGE-FLANGE pairs from touching count
    final_weld_count = total_welds + len(connected_branches) - touching_count_filtered
    
    print(f"\nSummary:")
    print(f"Total components: {len(result)}")
    print(f"Found in Excel: {found_count}")
    print(f"Not found in Excel: {len(result) - found_count}")
    print(f"Welded components (BWD or OLET): {welded_count}")
    print(f"Total weld count (from components): {total_welds}")
    print(f"Connected branches (branch-to-branch welds): {len(connected_branches)}")
    print(f"Touching component pairs (component-to-component welds, excl. OLETs & FLANGE-FLANGE): {touching_count_filtered}")
    print(f"Final weld count: {total_welds} + {len(connected_branches)} - {touching_count_filtered} = {final_weld_count}")

