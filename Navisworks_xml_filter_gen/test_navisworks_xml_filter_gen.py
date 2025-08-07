import os
import tempfile
import shutil
import datetime
import xml.etree.ElementTree as ET
import openpyxl
import pytest
from navisworks_xml_filter_gen import read_kks_codes, create_navisworks_xml

def create_sample_xlsx(path, codes):
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, code in enumerate(codes, start=1):
        ws[f'A{i}'] = code
    wb.save(path)

def test_read_kks_codes(tmp_path):
    codes = ['2HGC60BR100', '3ABC12DE34', 'XKKS000001']
    xlsx_path = tmp_path / 'Navis_KKS_List.xlsx'
    create_sample_xlsx(xlsx_path, codes)
    result = read_kks_codes(str(xlsx_path))
    assert result == codes

def test_create_navisworks_xml(tmp_path):
    codes = ['2HGC60BR100', '3ABC12DE34']
    output_path = tmp_path / 'test_output.xml'
    create_navisworks_xml(codes, str(output_path))
    tree = ET.parse(str(output_path))
    root = tree.getroot()
    # Check root and metadata
    assert root.tag == 'exchange'
    assert root.attrib['units'] == 'm'
    assert root.attrib['filename'] == 'ADI_Abu_Dhabi_20250806.nwd'
    # Check findspec and conditions
    findspec = root.find('findspec')
    assert findspec is not None
    conditions = findspec.find('conditions')
    assert conditions is not None
    condition_list = conditions.findall('condition')
    assert len(condition_list) == len(codes)
    for cond, code in zip(condition_list, codes):
        assert cond.attrib['test'] == 'contains'
        value = cond.find('value')
        data = value.find('data')
        assert data.text == code
    locator = findspec.find('locator')
    assert locator is not None
    assert locator.text == '/'

def test_main_script(tmp_path, monkeypatch):
    # Simulate running main()
    import navisworks_xml_filter_gen as script
    codes = ['A', 'B', 'C']
    xlsx_path = tmp_path / 'Navis_KKS_List.xlsx'
    create_sample_xlsx(xlsx_path, codes)
    # Patch __file__ to tmp_path
    monkeypatch.setattr(script, '__file__', str(tmp_path / 'navisworks_xml_filter_gen.py'))
    # Patch get_today_str to return fixed date string
    monkeypatch.setattr(script, 'get_today_str', lambda: '250806')
    script.main()
    output_name = '250806_navis_filter.xml'
    output_path = tmp_path / output_name
    assert output_path.exists()
    tree = ET.parse(str(output_path))
    root = tree.getroot()
    findspec = root.find('findspec')
    conditions = findspec.find('conditions')
    assert len(conditions.findall('condition')) == len(codes)
