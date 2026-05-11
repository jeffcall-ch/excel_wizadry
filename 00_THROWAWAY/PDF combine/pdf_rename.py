import re
from pathlib import Path

FOLDER = Path(r"C:\Users\szil\OneDrive - Kanadevia Inova\Downloads\TFM")

DRY_RUN = False  # Set to False to actually rename files

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}

# Pattern 1: TFM_2024_03March_Newsletter_...
_PAT1 = re.compile(r"^TFM_(\d{4})_(\d{2})[A-Za-z]+[_ ]", re.IGNORECASE)
# Pattern 2: TFM_April23_... or TFM_July17_...
_PAT2 = re.compile(r"^TFM_([A-Za-z]+)(\d{2})_", re.IGNORECASE)


def parse_date(name: str):
    m = _PAT1.match(name)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = _PAT2.match(name)
    if m:
        month = MONTHS.get(m.group(1).lower(), 0)
        year = 2000 + int(m.group(2))
        return year, month
    return None


pdfs = [p for p in FOLDER.glob("*.pdf")]
errors = []
renames = []

for pdf in sorted(pdfs):
    result = parse_date(pdf.name)
    if result is None:
        errors.append(pdf.name)
        continue
    year, month = result
    new_name = f"TFM_{year}_{month:02d}.pdf"
    new_path = pdf.parent / new_name
    renames.append((pdf, new_path))

# Check for duplicate targets
targets = [new for _, new in renames]
duplicates = {t for t in targets if targets.count(t) > 1}

print(f"{'DRY RUN - ' if DRY_RUN else ''}Renaming {len(renames)} files\n")
for old, new in renames:
    flag = " *** DUPLICATE TARGET ***" if new in duplicates else ""
    print(f"  {old.name}")
    print(f"    -> {new.name}{flag}")

if errors:
    print(f"\nWARNING: Could not parse date from {len(errors)} file(s):")
    for name in errors:
        print(f"  {name}")

if duplicates:
    print(f"\nERROR: Duplicate targets detected. Fix before renaming.")
elif not DRY_RUN:
    for old, new in renames:
        if new.exists() and new != old:
            print(f"SKIP (target exists): {new.name}")
        else:
            old.rename(new)
    print(f"\nDone. Renamed {len(renames)} files.")
else:
    print(f"\nDry run complete. Set DRY_RUN = False to apply.")
