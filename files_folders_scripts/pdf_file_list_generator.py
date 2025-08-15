import os
import csv
import re

source_dir = r"C:\temp\rev 0 supports"
output_csv = r"C:\Users\szil\Repos\excel_wizadry\files_folders_scripts\pdf_file_list_generator.py"

def extract_kks(filename):
    match = re.search(r'^[^-]+-[^-]+-([A-Z0-9]+)-', filename)
    return match.group(1) if match else ""

def extract_system_code(kks):
    match = re.search(r'([A-Z]{3}\d{2})', kks)
    return match.group(1) if match else ""

def extract_line_number(kks):
    return kks[0] if kks else ""

def extract_system_code_short(system_code):
    return system_code[:-2] if system_code and len(system_code) > 2 else system_code

with open(output_csv.replace('_generator.py', '.csv'), "w", newline='', encoding="utf-8") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["filename", "full_path", "KKS", "System_code", "Line_number", "System_code_short"])
    for root, _, files in os.walk(source_dir):
        for file in files:
            if file.lower().endswith(".pdf"):
                full_path = os.path.join(root, file)
                kks = extract_kks(file)
                system_code = extract_system_code(kks)
                line_number = extract_line_number(kks)
                system_code_short = extract_system_code_short(system_code)
                writer.writerow([file, full_path, kks, system_code, line_number, system_code_short])

print("Done! CSV created.")
