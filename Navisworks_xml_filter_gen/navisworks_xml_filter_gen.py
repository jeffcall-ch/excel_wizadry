def get_today_str():
    return datetime.datetime.now().strftime('%y%m%d')
import os
import datetime
import xml.etree.ElementTree as ET
import openpyxl

def read_kks_codes(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    ws = wb.active
    codes = []
    for row in ws.iter_rows(min_row=1, max_col=1, values_only=True):
        code = row[0]
        if code:
            # Remove all whitespace (including spaces within the KKS code)
            clean_code = str(code).replace(' ', '').strip()
            if clean_code:  # Only add non-empty codes
                codes.append(clean_code)
    return codes

def create_navisworks_xml(kks_codes, output_path):
    # Create XML with minimal metadata (leaving most blank as requested)
    exchange = ET.Element('exchange', {
        'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        'xsi:noNamespaceSchemaLocation': 'http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd',
        'units': 'm',
        'filename': '',
        'filepath': ''
    })
    findspec = ET.SubElement(exchange, 'findspec', {'mode': 'all', 'disjoint': '0'})
    conditions = ET.SubElement(findspec, 'conditions')
    
    for i, code in enumerate(kks_codes):
        # First condition uses flags="10", subsequent ones use flags="74" (OR logic)
        flags = "10" if i == 0 else "74"
        condition = ET.SubElement(conditions, 'condition', {'test': 'contains', 'flags': flags})
        prop = ET.SubElement(condition, 'property')
        name = ET.SubElement(prop, 'name', {'internal': 'LcOaSceneBaseUserName'})
        name.text = 'Name'
        value = ET.SubElement(condition, 'value')
        data = ET.SubElement(value, 'data', {'type': 'wstring'})
        data.text = code
    
    locator = ET.SubElement(findspec, 'locator')
    locator.text = '/'
    
    # Format for pretty printing
    ET.indent(exchange, space="  ", level=0)
    
    # Write XML with validation
    tree = ET.ElementTree(exchange)
    tree.write(output_path, encoding='utf-8', xml_declaration=True)
    
    # Validate XML structure
    try:
        # Re-read to validate
        ET.parse(output_path)
        print(f'XML validation: PASSED')
        return True
    except ET.ParseError as e:
        print(f'XML validation: FAILED - {e}')
        return False

def main():
    folder = os.path.dirname(os.path.abspath(__file__))
    xlsx_path = os.path.join(folder, 'Navis_KKS_List.xlsx')
    
    if not os.path.exists(xlsx_path):
        print(f'Error: Excel file not found: {xlsx_path}')
        return
    
    kks_codes = read_kks_codes(xlsx_path)
    
    if not kks_codes:
        print('Warning: No KKS codes found in Excel file')
        return
    
    print(f'Found {len(kks_codes)} KKS codes')
    
    today = get_today_str()
    output_name = f'{today}_navis_filter.xml'
    output_path = os.path.join(folder, output_name)
    
    success = create_navisworks_xml(kks_codes, output_path)
    
    if success:
        print(f'XML file generated successfully: {output_path}')
        print(f'Total KKS codes processed: {len(kks_codes)}')
    else:
        print('Failed to generate valid XML file')

if __name__ == '__main__':
    main()
