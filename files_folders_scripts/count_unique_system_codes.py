import csv

csv_path = r"C:\Users\szil\Repos\excel_wizadry\files_folders_scripts\pdf_file_list.csv"
system_codes = set()

with open(csv_path, newline='', encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        system_codes.add(row["System_code"])

print(f"Unique system codes: {len(system_codes)}")
