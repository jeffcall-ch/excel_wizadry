import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import uuid
import re
from pathlib import Path
from collections import defaultdict

# File paths
EXCEL_FILE = r"C:\Users\szil\Repos\excel_wizadry\Navisworks_viewpoint_generation_from_support_list\GP_BQ_30.01.2026_CA.xlsx"
TEMPLATE_XML = r"C:\Users\szil\Repos\excel_wizadry\Navisworks_viewpoint_generation_from_support_list\TBY_Thameside_ERF_20251117_viewpoints_251125_.xml"
OUTPUT_XML = r"C:\Users\szil\Repos\excel_wizadry\Navisworks_viewpoint_generation_from_support_list\generated_viewpoints.xml"

# Camera offset from target point (adjusted for closer zoom)
CAMERA_OFFSET_X = 1.5  # meters - reduced from 3.45 for closer view
CAMERA_OFFSET_Y = 1.5  # meters - reduced from 2.98 for closer view
CAMERA_OFFSET_Z = 1.5  # meters - reduced from 2.58 for closer view

# Reference quaternion for camera rotation (looking at target from oblique angle)
CAMERA_ROTATION = {
    'a': 0.1759198966,
    'b': 0.4247082003,
    'c': 0.8204732386,
    'd': 0.3398511430
}

# Bounding box dimensions (half-widths from center)
BOX_HALF_WIDTH_X = 1.0  # 2m total width
BOX_HALF_WIDTH_Y = 1.0  # 2m total depth
BOX_HALF_WIDTH_Z = 1.0  # 2m total height


def parse_coordinates(coord_str):
    """
    Parse coordinate string like 'X 18040mm Y 94404.038mm Z 186.7mm'
    Returns tuple (X, Y, Z) as integers (stripped of decimals and 'mm')
    """
    if pd.isna(coord_str):
        return None, None, None
    
    # Extract X, Y, Z values
    x_match = re.search(r'X\s*([\d.]+)mm', str(coord_str))
    y_match = re.search(r'Y\s*([\d.]+)mm', str(coord_str))
    z_match = re.search(r'Z\s*([\d.]+)mm', str(coord_str))
    
    x = int(float(x_match.group(1))) if x_match else None
    y = int(float(y_match.group(1))) if y_match else None
    z = int(float(z_match.group(1))) if z_match else None
    
    return x, y, z


def extract_subfolder_name(su_ref):
    """Extract first 7 characters from SU/STEEL ref for subfolder name."""
    if pd.isna(su_ref):
        return ""
    return str(su_ref)[:7]


def extract_bq_subfolder_name(bq_name):
    """Extract first 6 characters from BQ Name for subfolder name."""
    if pd.isna(bq_name):
        return ""
    return str(bq_name)[:6]


def strip_numeric_prefix(zone_name):
    """Strip /[0-9] prefix from zone name to create parent folder.
    E.g., '/0AUX-P/PW' -> 'AUX-P/PW'
    """
    if pd.isna(zone_name):
        return ""
    zone_str = str(zone_name)
    # Check if starts with /[0-9]
    if zone_str.startswith('/') and len(zone_str) > 1 and zone_str[1].isdigit():
        return zone_str[2:]  # Strip /[0-9]
    return zone_str


def handle_bq_duplicates(df):
    """Handle duplicate BQ Names by appending _1, _2, etc. to viewpoint names."""
    # Count occurrences of each BQ Name
    bq_counts = defaultdict(int)
    unique_names = []
    
    for bq_name in df['BQ Name']:
        if pd.isna(bq_name):
            unique_names.append(bq_name)
            continue
        
        bq_counts[bq_name] += 1
        if bq_counts[bq_name] == 1:
            unique_names.append(bq_name)
        else:
            unique_names.append(f"{bq_name}_{bq_counts[bq_name]}")
    
    return unique_names


def generate_deterministic_guid(seed_string):
    """
    Generate a deterministic GUID based on a seed string.
    This ensures the same input always generates the same GUID.
    Uses UUID5 (SHA-1 hash) with a custom namespace.
    """
    # Custom namespace for this project
    namespace = uuid.UUID('6ba7b810-9dad-11d1-80b4-00c04fd430c8')  # DNS namespace
    return str(uuid.uuid5(namespace, seed_string))


def create_viewpoint_element(name, x, y, z, guid_seed):
    """
    Create a viewpoint XML element with camera and clipping planes.
    Coordinates are in millimeters, will be converted to meters.
    """
    # Convert mm to meters
    x_m = x / 1000.0
    y_m = y / 1000.0
    z_m = z / 1000.0
    
    # Calculate camera position
    cam_x = x_m + CAMERA_OFFSET_X
    cam_y = y_m + CAMERA_OFFSET_Y
    cam_z = z_m + CAMERA_OFFSET_Z
    
    # Calculate bounding box
    box_min_x = x_m - BOX_HALF_WIDTH_X
    box_min_y = y_m - BOX_HALF_WIDTH_Y
    box_min_z = z_m - BOX_HALF_WIDTH_Z
    
    box_max_x = x_m + BOX_HALF_WIDTH_X
    box_max_y = y_m + BOX_HALF_WIDTH_Y
    box_max_z = z_m + BOX_HALF_WIDTH_Z
    
    # Generate GUID
    view_guid = generate_deterministic_guid(guid_seed)
    
    # Create XML structure
    view = ET.Element('view', name=name, guid=view_guid)
    
    # Viewpoint
    viewpoint = ET.SubElement(view, 'viewpoint', 
                              tool="autocam_orbit", 
                              render="shaded", 
                              lighting="headlight",
                              focal="5.2995378803",
                              linear="34.9251909750",
                              angular="0.7853981634")
    
    # Camera
    camera = ET.SubElement(viewpoint, 'camera',
                          projection="persp",
                          near="4.0231288732",
                          far="45.0566477257",
                          aspect="1.6675496689",
                          height="0.8975073621")
    
    position = ET.SubElement(camera, 'position')
    ET.SubElement(position, 'pos3f', 
                  x=f"{cam_x:.10f}",
                  y=f"{cam_y:.10f}",
                  z=f"{cam_z:.10f}")
    
    rotation = ET.SubElement(camera, 'rotation')
    ET.SubElement(rotation, 'quaternion',
                  a=f"{CAMERA_ROTATION['a']:.10f}",
                  b=f"{CAMERA_ROTATION['b']:.10f}",
                  c=f"{CAMERA_ROTATION['c']:.10f}",
                  d=f"{CAMERA_ROTATION['d']:.10f}")
    
    # Viewer
    ET.SubElement(viewpoint, 'viewer',
                  radius="0.3000000000",
                  height="1.8000000000",
                  actual_height="1.8000000000",
                  eye_height="0.1500000000",
                  avatar="construction_worker",
                  camera_mode="first",
                  first_to_third_angle="0.0000000000",
                  first_to_third_distance="3.0000000000",
                  first_to_third_param="1.0000000000",
                  first_to_third_correction="1",
                  collision_detection="0",
                  auto_crouch="0",
                  gravity="0",
                  gravity_value="9.8000000000",
                  terminal_velocity="50.0000000000")
    
    # Up vector
    up = ET.SubElement(viewpoint, 'up')
    ET.SubElement(up, 'vec3f', x="0.0000000000", y="0.0000000000", z="1.0000000000")
    
    # Clip plane set
    clipplaneset = ET.SubElement(view, 'clipplaneset',
                                  linked="0",
                                  current="0",
                                  mode="box",
                                  enabled="1")
    
    # Range
    range_elem = ET.SubElement(clipplaneset, 'range')
    range_box = ET.SubElement(range_elem, 'box3f')
    range_min = ET.SubElement(range_box, 'min')
    ET.SubElement(range_min, 'pos3f', x="1.0000000000", y="1.0000000000", z="1.0000000000")
    range_max = ET.SubElement(range_box, 'max')
    ET.SubElement(range_max, 'pos3f', x="0.0000000000", y="0.0000000000", z="0.0000000000")
    
    # Clip planes (6 planes defining the box)
    clipplanes = ET.SubElement(clipplaneset, 'clipplanes')
    
    # Top plane
    cp_top = ET.SubElement(clipplanes, 'clipplane', state="default", distance="0.0000000000", alignment="top")
    plane_top = ET.SubElement(cp_top, 'plane', distance=f"{-box_max_z:.10f}")
    ET.SubElement(plane_top, 'vec3f', x="-0.0000000000", y="-0.0000000000", z="-1.0000000000")
    
    # Bottom plane
    cp_bottom = ET.SubElement(clipplanes, 'clipplane', state="default", distance="0.0000000000", alignment="bottom")
    plane_bottom = ET.SubElement(cp_bottom, 'plane', distance="0.0000000000")
    ET.SubElement(plane_bottom, 'vec3f', x="0.0000000000", y="1.0000000000", z="0.0000000000")
    
    # Front plane
    cp_front = ET.SubElement(clipplanes, 'clipplane', state="default", distance="0.0000000000", alignment="front")
    plane_front = ET.SubElement(cp_front, 'plane', distance=f"{box_max_y:.10f}")
    ET.SubElement(plane_front, 'vec3f', x="0.0000000000", y="1.0000000000", z="0.0000000000")
    
    # Back plane
    cp_back = ET.SubElement(clipplanes, 'clipplane', state="default", distance="0.0000000000", alignment="back")
    plane_back = ET.SubElement(cp_back, 'plane', distance=f"{-box_min_y:.10f}")
    ET.SubElement(plane_back, 'vec3f', x="-0.0000000000", y="-1.0000000000", z="-0.0000000000")
    
    # Left plane
    cp_left = ET.SubElement(clipplanes, 'clipplane', state="default", distance="0.0000000000", alignment="left")
    plane_left = ET.SubElement(cp_left, 'plane', distance=f"{box_max_x:.10f}")
    ET.SubElement(plane_left, 'vec3f', x="1.0000000000", y="0.0000000000", z="0.0000000000")
    
    # Right plane
    cp_right = ET.SubElement(clipplanes, 'clipplane', state="default", distance="0.0000000000", alignment="right")
    plane_right = ET.SubElement(cp_right, 'plane', distance=f"{-box_min_x:.10f}")
    ET.SubElement(plane_right, 'vec3f', x="-1.0000000000", y="0.0000000000", z="0.0000000000")
    
    # Box
    box_elem = ET.SubElement(clipplaneset, 'box')
    box_3f = ET.SubElement(box_elem, 'box3f')
    box_min = ET.SubElement(box_3f, 'min')
    ET.SubElement(box_min, 'pos3f',
                  x=f"{box_min_x:.10f}",
                  y=f"{box_min_y:.10f}",
                  z=f"{box_min_z:.10f}")
    box_max = ET.SubElement(box_3f, 'max')
    ET.SubElement(box_max, 'pos3f',
                  x=f"{box_max_x:.10f}",
                  y=f"{box_max_y:.10f}",
                  z=f"{box_max_z:.10f}")
    
    # Box rotation
    box_rotation = ET.SubElement(clipplaneset, 'box-rotation')
    box_rot = ET.SubElement(box_rotation, 'rotation')
    ET.SubElement(box_rot, 'quaternion', a="0.0000000000", b="0.0000000000", c="0.0000000000", d="1.0000000000")
    
    return view


def prettify_xml(elem):
    """Return a pretty-printed XML string for the Element."""
    rough_string = ET.tostring(elem, encoding='unicode')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")


def main():
    print("Reading Excel file...")
    df = pd.read_excel(EXCEL_FILE)
    
    print(f"Total rows: {len(df)}")
    
    # Separate into two groups: with SU and without SU
    df_with_su = df[df['SU/STEEL ref'] != '-'].copy()
    df_without_su = df[df['SU/STEEL ref'] == '-'].copy()
    
    print(f"Rows with SU/STEEL ref (not '-'): {len(df_with_su)}")
    print(f"Rows without SU/STEEL ref ('-'): {len(df_without_su)}")
    
    # ==================== PROCESS GP_SU GROUP (EXISTING LOGIC) ====================
    print("\n=== Processing GP_SU group ===")
    
    # Filter to unique combinations of SU/STEEL ref + BQ Zone
    # (Same SU can appear in multiple zones)
    df_su_unique = df_with_su.drop_duplicates(subset=['SU/STEEL ref', 'BQ Zone']).copy()
    print(f"Unique SU/STEEL ref + BQ Zone combinations: {len(df_su_unique)}")
    
    # Parse coordinates
    print("Parsing coordinates...")
    df_su_unique[['X', 'Y', 'Z']] = df_su_unique['POS WRT /*'].apply(
        lambda x: pd.Series(parse_coordinates(x))
    )
    
    # Remove rows with missing coordinates
    df_su_clean = df_su_unique.dropna(subset=['X', 'Y', 'Z']).copy()
    print(f"GP_SU rows with valid coordinates: {len(df_su_clean)}")
    
    # Extract subfolder names (first 7 chars)
    df_su_clean['Subfolder'] = df_su_clean['SU/STEEL ref'].apply(extract_subfolder_name)
    
    # Extract parent folder from BQ Zone
    df_su_clean['ParentFolder'] = df_su_clean['BQ Zone'].apply(strip_numeric_prefix)
    
    # Sort
    df_su_sorted = df_su_clean.sort_values(by=['ParentFolder', 'BQ Zone', 'Subfolder', 'SU/STEEL ref'])
    print(f"Sorted {len(df_su_sorted)} GP_SU items")
    
    # Build hierarchy for GP_SU
    hierarchy_su = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    
    for _, row in df_su_sorted.iterrows():
        parent = row['ParentFolder']
        zone = row['BQ Zone']
        subfolder = row['Subfolder']
        su_ref = row['SU/STEEL ref']
        x, y, z = int(row['X']), int(row['Y']), int(row['Z'])
        
        hierarchy_su[parent][zone][subfolder].append({
            'name': su_ref,
            'x': x,
            'y': y,
            'z': z
        })
    
    # ==================== PROCESS GP_BQ_without_SU GROUP (NEW LOGIC) ====================
    print("\n=== Processing GP_BQ_without_SU group ===")
    
    # Filter to unique combinations of BQ Name + BQ Zone
    # (Same BQ Name can appear in multiple zones)
    df_bq_unique = df_without_su.drop_duplicates(subset=['BQ Name', 'BQ Zone']).copy()
    print(f"Unique BQ Name + BQ Zone combinations: {len(df_bq_unique)}")
    
    # Parse coordinates
    print("Parsing coordinates...")
    df_bq_unique[['X', 'Y', 'Z']] = df_bq_unique['POS WRT /*'].apply(
        lambda x: pd.Series(parse_coordinates(x))
    )
    
    # Remove rows with missing coordinates
    df_bq_clean = df_bq_unique.dropna(subset=['X', 'Y', 'Z']).copy()
    print(f"GP_BQ rows with valid coordinates: {len(df_bq_clean)}")
    
    # Handle duplicates in BQ Name by appending _1, _2, etc.
    # Note: This handles cases where same BQ Name appears multiple times within same processing batch
    df_bq_clean['UniqueViewpointName'] = handle_bq_duplicates(df_bq_clean)
    
    # Extract subfolder names (first 6 chars of BQ Name)
    df_bq_clean['Subfolder'] = df_bq_clean['BQ Name'].apply(extract_bq_subfolder_name)
    
    # Extract parent folder from BQ Zone
    df_bq_clean['ParentFolder'] = df_bq_clean['BQ Zone'].apply(strip_numeric_prefix)
    
    # Sort
    df_bq_sorted = df_bq_clean.sort_values(by=['ParentFolder', 'BQ Zone', 'Subfolder', 'BQ Name'])
    print(f"Sorted {len(df_bq_sorted)} GP_BQ items")
    
    # Build hierarchy for GP_BQ
    hierarchy_bq = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
    
    for _, row in df_bq_sorted.iterrows():
        parent = row['ParentFolder']
        zone = row['BQ Zone']
        subfolder = row['Subfolder']
        viewpoint_name = row['UniqueViewpointName']
        x, y, z = int(row['X']), int(row['Y']), int(row['Z'])
        
        hierarchy_bq[parent][zone][subfolder].append({
            'name': viewpoint_name,
            'x': x,
            'y': y,
            'z': z
        })
    
    # ==================== CREATE XML WITH BOTH HIERARCHIES ====================
    print("\n=== Generating XML ===")
    root = ET.Element('exchange')
    root.set('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')
    root.set('xsi:noNamespaceSchemaLocation', 'http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd')
    root.set('units', 'm')
    root.set('filename', 'generated_viewpoints.nwd')
    root.set('filepath', r'C:\Users\szil\Repos\excel_wizadry\Navisworks_viewpoint_generation_from_support_list')
    
    viewpoints = ET.SubElement(root, 'viewpoints')
    
    # ==================== CREATE GP_SU FOLDER ====================
    gp_su_guid = generate_deterministic_guid("GP_SU_top_folder")
    gp_su_folder = ET.SubElement(viewpoints, 'viewfolder', name="GP_SU", guid=gp_su_guid)
    
    for parent in sorted(hierarchy_su.keys()):
        parent_guid = generate_deterministic_guid(f"su_parent_{parent}")
        parent_folder = ET.SubElement(gp_su_folder, 'viewfolder', name=str(parent), guid=parent_guid)
        
        for zone in sorted(hierarchy_su[parent].keys()):
            zone_guid = generate_deterministic_guid(f"su_zone_{parent}_{zone}")
            zone_folder = ET.SubElement(parent_folder, 'viewfolder', name=str(zone), guid=zone_guid)
            
            for subfolder in sorted(hierarchy_su[parent][zone].keys()):
                subfolder_guid = generate_deterministic_guid(f"su_subfolder_{parent}_{zone}_{subfolder}")
                subfolder_elem = ET.SubElement(zone_folder, 'viewfolder', name=str(subfolder), guid=subfolder_guid)
                
                for item in hierarchy_su[parent][zone][subfolder]:
                    guid_seed = f"su_view_{parent}_{zone}_{subfolder}_{item['name']}_{item['x']}_{item['y']}_{item['z']}"
                    view = create_viewpoint_element(
                        item['name'],
                        item['x'],
                        item['y'],
                        item['z'],
                        guid_seed
                    )
                    subfolder_elem.append(view)
    
    # ==================== CREATE GP_BQ_without_SU FOLDER ====================
    gp_bq_guid = generate_deterministic_guid("GP_BQ_without_SU_top_folder")
    gp_bq_folder = ET.SubElement(viewpoints, 'viewfolder', name="GP_BQ_without_SU", guid=gp_bq_guid)
    
    for parent in sorted(hierarchy_bq.keys()):
        parent_guid = generate_deterministic_guid(f"bq_parent_{parent}")
        parent_folder = ET.SubElement(gp_bq_folder, 'viewfolder', name=str(parent), guid=parent_guid)
        
        for zone in sorted(hierarchy_bq[parent].keys()):
            zone_guid = generate_deterministic_guid(f"bq_zone_{parent}_{zone}")
            zone_folder = ET.SubElement(parent_folder, 'viewfolder', name=str(zone), guid=zone_guid)
            
            for subfolder in sorted(hierarchy_bq[parent][zone].keys()):
                subfolder_guid = generate_deterministic_guid(f"bq_subfolder_{parent}_{zone}_{subfolder}")
                subfolder_elem = ET.SubElement(zone_folder, 'viewfolder', name=str(subfolder), guid=subfolder_guid)
                
                for item in hierarchy_bq[parent][zone][subfolder]:
                    guid_seed = f"bq_view_{parent}_{zone}_{subfolder}_{item['name']}_{item['x']}_{item['y']}_{item['z']}"
                    view = create_viewpoint_element(
                        item['name'],
                        item['x'],
                        item['y'],
                        item['z'],
                        guid_seed
                    )
                    subfolder_elem.append(view)
    
    # ==================== WRITE XML ====================
    print(f"Writing XML to {OUTPUT_XML}...")
    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ")
    
    with open(OUTPUT_XML, 'wb') as f:
        f.write(b'<?xml version="1.0" encoding="UTF-8" ?>\n\n')
        tree.write(f, encoding='utf-8', xml_declaration=False)
    
    # ==================== STATISTICS ====================
    total_su_parents = len(hierarchy_su)
    total_su_zones = sum(len(zones) for zones in hierarchy_su.values())
    total_bq_parents = len(hierarchy_bq)
    total_bq_zones = sum(len(zones) for zones in hierarchy_bq.values())
    
    print(f"\n✓ Successfully generated viewpoints:")
    print(f"  - GP_SU: {len(df_su_sorted)} viewpoints in {total_su_parents} parent folders, {total_su_zones} zone folders")
    print(f"  - GP_BQ_without_SU: {len(df_bq_sorted)} viewpoints in {total_bq_parents} parent folders, {total_bq_zones} zone folders")
    print(f"✓ Output file: {OUTPUT_XML}")


if __name__ == "__main__":
    main()
