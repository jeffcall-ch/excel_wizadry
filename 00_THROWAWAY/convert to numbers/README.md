# Convert Excel Cells to Numbers

This utility scans a folder tree, finds every `.xlsx` or `.xlsm` workbook, and converts cells that contain numeric-looking text into true Excel numbers while preserving the visible decimal precision. The conversion is performed in place, so keep a backup of the files you plan to process.

## Requirements

- Python 3.10+
- Packages: `pandas`, `openpyxl`

Install or update dependencies via:

```bash
pip install -r requirements.txt
```

## Usage

```bash
python "convert_excel_cells_to_numbers.py" <folder> [--dry-run] [--log-level LEVEL]
```

Arguments:
- `folder`: Root directory to scan recursively for `.xlsx` and `.xlsm` files.
- `--dry-run`: Process files but skip saving changes; useful for verification.
- `--log-level`: Standard Python logging level (e.g., `DEBUG`, `INFO`, `WARNING`). Defaults to `INFO`.

Example:

```bash
python "convert_excel_cells_to_numbers.py" "d:/Projects/Reports" --log-level DEBUG
```

## What the Script Does

1. Walks every subfolder under the provided root to find Excel files.
2. Opens each workbook with `pandas` to ensure every sheet is touched.
3. Loads the workbook with `openpyxl` for precise cell-level edits.
4. Detects strings that match `^[+-]?\d+(?:\.\d+)?$`.
5. Converts the cell value to a real number and assigns an Excel number format (`0`, `0.0`, `0.00`, etc.) that mirrors the original decimal count.
6. Saves the workbook in place (unless `--dry-run` is used).

## Notes and Limitations

- Only `.` is treated as the decimal separator, and thousands separators are ignored.
- Hidden sheets and cells are processed just like visible ones.
- Macros are preserved when converting `.xlsm` files (`keep_vba=True`).
- Formatting other than the number format is left untouched.
- If a cell contains non-numeric text (e.g., `5.0 kg`), it is skipped.

## Troubleshooting

- Use `--log-level DEBUG` to see which cells and sheets are being processed.
- If a file fails to open, the error is logged and the script continues with the next workbook.
- On Windows paths with spaces, wrap the arguments in quotes as shown in the example above.
