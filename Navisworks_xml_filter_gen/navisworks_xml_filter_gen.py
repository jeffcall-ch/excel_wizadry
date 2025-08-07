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
            codes.append(str(code).strip())
    return codes

def create_navisworks_xml(kks_codes, output_path):
    # Metadata from example
    exchange = ET.Element('exchange', {
        'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        'xsi:noNamespaceSchemaLocation': 'http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd',
        'units': 'm',
        'filename': 'ADI_Abu_Dhabi_20250806.nwd',
        'filepath': r'C:\Users\szil\OneDrive - Kanadevia Inova\Desktop\Projects\03 AbuDhabi\02 Navisworks'
    })
    findspec = ET.SubElement(exchange, 'findspec', {'mode': 'all', 'disjoint': '0'})
    conditions = ET.SubElement(findspec, 'conditions')
    for code in kks_codes:
        condition = ET.SubElement(conditions, 'condition', {'test': 'contains', 'flags': '10'})
        prop = ET.SubElement(condition, 'property')
        name = ET.SubElement(prop, 'name', {'internal': 'LcOaSceneBaseUserName'})
        name.text = 'Name'
        value = ET.SubElement(condition, 'value')
        data = ET.SubElement(value, 'data', {'type': 'wstring'})
        data.text = code
    locator = ET.SubElement(findspec, 'locator')
    locator.text = '/'
    tree = ET.ElementTree(exchange)
    tree.write(output_path, encoding='utf-8', xml_declaration=True)

def main():
    folder = os.path.dirname(os.path.abspath(__file__))
    xlsx_path = os.path.join(folder, 'Navis_KKS_List.xlsx')
    kks_codes = read_kks_codes(xlsx_path)
    today = get_today_str()
    output_name = f'{today}_navis_filter.xml'
    output_path = os.path.join(folder, output_name)
    create_navisworks_xml(kks_codes, output_path)
    print(f'XML file generated: {output_path}')

if __name__ == '__main__':
    main()
