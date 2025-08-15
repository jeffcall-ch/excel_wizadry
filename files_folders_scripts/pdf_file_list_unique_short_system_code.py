import csv
import random

csv_path = r"C:\Users\szil\Repos\excel_wizadry\files_folders_scripts\pdf_file_list.csv"
output_csv = r"C:\Users\szil\Repos\excel_wizadry\files_folders_scripts\pdf_file_list_unique_short_system_code.csv"

rows_by_short_code = {}

with open(csv_path, newline='', encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        short_code = row["System_code_short"]
        if short_code not in rows_by_short_code:
            rows_by_short_code[short_code] = []
        rows_by_short_code[short_code].append(row)

unique_rows = [random.choice(rows) for rows in rows_by_short_code.values()]

with open(output_csv, "w", newline='', encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=unique_rows[0].keys())
    writer.writeheader()
    writer.writerows(unique_rows)

print("Done! Unique short system code CSV created.")
