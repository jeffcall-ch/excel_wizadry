"""Audit workbook footers in the adjacent 01 Lists directory."""

import argparse
from pathlib import Path

from openpyxl import load_workbook


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
SKIP_SHEETS = {"coversheet", "instructions"}


def check_footer(footer, filename, sheet_name):
    left = footer.left.text or ""
    right = footer.right.text or ""
    missing = []
    if "&F" not in left and filename not in left:
        missing.append("file name on left")
    if "&A" not in left and sheet_name not in left:
        missing.append("sheet name on left")
    if "&P" not in right:
        missing.append("page number on right")
    if "&N" not in right:
        missing.append("total pages on right")
    return missing


def audit(root):
    file_count = 0
    sheet_count = 0
    problems = 0
    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in EXCEL_EXTENSIONS or path.name.startswith("~$"):
            continue
        file_count += 1
        try:
            workbook = load_workbook(path, read_only=False, keep_links=False)
            try:
                for sheet in workbook.worksheets:
                    if sheet.title.casefold() in SKIP_SHEETS:
                        continue
                    sheet_count += 1
                    footers = [("odd/default", sheet.oddFooter)]
                    if sheet.HeaderFooter.differentOddEven:
                        footers.append(("even", sheet.evenFooter))
                    if sheet.HeaderFooter.differentFirst:
                        footers.append(("first", sheet.firstFooter))
                    for label, footer in footers:
                        missing = check_footer(footer, path.name, sheet.title)
                        if missing:
                            problems += 1
                            print(
                                f"{path.relative_to(root)} | {sheet.title} | {label}: "
                                f"{', '.join(missing)} (left={footer.left.text!r}, "
                                f"right={footer.right.text!r})"
                            )
            finally:
                workbook.close()
        except Exception as error:
            problems += 1
            print(f"{path.relative_to(root)} | COULD NOT READ: {error}")
    print(f"Checked {sheet_count} sheets in {file_count} files; {problems} problem(s).")
    return problems


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", nargs="?", type=Path, default=Path(__file__).resolve().parent / "01 Lists")
    root = parser.parse_args().directory
    if not root.is_dir():
        parser.error(f"not a directory: {root}")
    raise SystemExit(1 if audit(root) else 0)


if __name__ == "__main__":
    main()