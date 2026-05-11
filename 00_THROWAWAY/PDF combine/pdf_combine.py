import re
from pathlib import Path
from pypdf import PdfWriter

FOLDER = Path(r"C:\Users\szil\OneDrive - Kanadevia Inova\Downloads\TFM")
OUTPUT = FOLDER / "_combined.pdf"

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
    "jan": 1, "feb": 2, "mar": 3, "apr": 4,
    "jun": 6, "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

# Pattern 1: TFM_2024_03March_Newsletter_...  -> year=2024, month=03
_PAT1 = re.compile(r"TFM_(\d{4})_(\d{2})\w+_", re.IGNORECASE)
# Pattern 2: TFM_April23_...  -> month=April, year=20YY
_PAT2 = re.compile(r"TFM_([A-Za-z]+)(\d{2})_", re.IGNORECASE)


def date_key(path: Path):
    name = path.name
    m = _PAT1.search(name)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    m = _PAT2.search(name)
    if m:
        month_str = m.group(1).lower()
        year = 2000 + int(m.group(2))
        month = MONTHS.get(month_str, 0)
        return (year, month)
    print(f"  WARNING: could not parse date from '{name}', placing at end.")
    return (0, 0)


pdfs = [p for p in FOLDER.glob("*.pdf") if p.name != OUTPUT.name]
pdfs.sort(key=date_key, reverse=True)  # most recent first

if not pdfs:
    print("No PDF files found.")
else:
    writer = PdfWriter()
    for pdf in pdfs:
        key = date_key(pdf)
        print(f"Adding [{key[0]}-{key[1]:02d}]: {pdf.name}")
        writer.append(str(pdf))
    with open(OUTPUT, "wb") as f:
        writer.write(f)
    print(f"\nDone. Combined {len(pdfs)} files -> {OUTPUT}")
