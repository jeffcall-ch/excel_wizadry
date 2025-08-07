import os
import tempfile
import pandas as pd
import openpyxl
import csv
from excel_file_search import search_excel_for_text, search_directory_for_excels

def create_test_xlsx(path, sheet_name, data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row in data:
        ws.append(row)
    wb.save(path)

def create_test_csv(path, data):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerows(data)

def test_search_excel_for_text():
    with tempfile.TemporaryDirectory() as tmpdir:
        # Create XLSX file
        xlsx_path = os.path.join(tmpdir, 'test.xlsx')
        create_test_xlsx(xlsx_path, 'Sheet1', [['foo', 'bar'], ['baz', 'searchPhrase']])
        # Create CSV file
        csv_path = os.path.join(tmpdir, 'test.csv')
        create_test_csv(csv_path, [['hello', 'searchPhrase'], ['world', 'test']])
        # Test XLSX
        matches_xlsx = search_excel_for_text(xlsx_path, 'searchPhrase')
        assert any('searchPhrase' in m['match_context'] for m in matches_xlsx)
        # Test CSV
        matches_csv = search_excel_for_text(csv_path, 'searchPhrase')
        assert any('searchPhrase' in m['match_context'] for m in matches_csv)
        # Test directory search
        matches_dir = search_directory_for_excels(tmpdir, 'searchPhrase')
        assert len(matches_dir) == 2
        print('All tests passed.')

if __name__ == "__main__":
    test_search_excel_for_text()
