# comos_extract_text_to_number

## Purpose
This folder contains scripts for processing and converting pipe data from Excel files. The main scripts are:
- `excel_text_to_number_bck.py`: Reads an Excel file and converts columns containing pipe dimensions (e.g., 'DN', 'PN') to numeric values, preserving decimal places.
- `pipe_category_calculator.py`: Reads an Excel file, extracts numeric values from pipe dimension columns, and performs category calculations for further analysis.

**Input:**
- Excel file (e.g., `Export_pipe_list_30.07.2025_withfluid_code.xlsx`) in the same folder.

**Output:**
- Processed DataFrame (in-memory) or output file, depending on script modifications.

## Usage
1. Place the required Excel file in this folder.
2. Install dependencies:
   ```powershell
   pip install pandas
   ```
3. Run the script from the command line:
   ```powershell
   python excel_text_to_number_bck.py
   python pipe_category_calculator.py
   ```

## Examples
```powershell
python excel_text_to_number_bck.py
python pipe_category_calculator.py
```

## Known Limitations
- Input file name is hardcoded; change the script if your file is named differently.
- Only processes the first sheet of the Excel file.
- Assumes specific column formats (e.g., 'DN <number>', 'PN <number>').
- No output file is written by default; modify the script to save results if needed.
- Large files may impact performance.
