import csv
import shutil
import os

csv_path = r"C:\Users\szil\Repos\excel_wizadry\files_folders_scripts\pdf_file_list_unique_short_system_code.csv"
dest_dir = r"C:\temp\support_check_14.08.2025"

os.makedirs(dest_dir, exist_ok=True)

with open(csv_path, newline='', encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        src = row["full_path"]
        if src and os.path.isfile(src):
            shutil.copy2(src, dest_dir)

print("Done! Files copied.")
