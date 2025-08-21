# Navisworks XML Filter Generator

This tool generates Navisworks XML filter files from Excel lists of KKS codes. It reads KKS codes from an Excel file and creates a properly formatted XML filter that can be imported into Navisworks for selective object visibility.

## What it does

The script automatically:
- Reads KKS codes from an Excel file (`Navis_KKS_List.xlsx`)
- Cleans and validates the codes (removes spaces, filters empty entries)
- Generates a Navisworks-compatible XML filter file
- Uses proper filter logic (first condition AND, subsequent conditions OR)
- Validates the generated XML structure
- Names the output file with today's date (e.g., `250811_navis_filter.xml`)

## Requirements

- Python 3.8+ with virtual environment
- openpyxl library (for Excel file reading)
- Excel file named `Navis_KKS_List.xlsx` in the same directory

## File Structure

```
Navisworks_xml_filter_gen/
├── navisworks_xml_filter_gen.py          # Main Python script
├── run_navisworks_xml_generator.bat      # Windows batch file (recommended)
├── run_navisworks_xml_generator.ps1      # PowerShell script (alternative)
├── Navis_KKS_List.xlsx                   # Input Excel file with KKS codes
├── navis_example.xml                     # Reference XML structure
└── README.md                             # This file
```

## Usage

### Method 1: Batch File (Recommended)
Simply double-click `run_navisworks_xml_generator.bat` or run from command line:
```cmd
run_navisworks_xml_generator.bat
```

### Method 2: PowerShell Script
```powershell
.\run_navisworks_xml_generator.ps1
```

### Method 3: Direct Python Execution
```bash
# Activate virtual environment first
..\.venv\Scripts\Activate.ps1
python navisworks_xml_filter_gen.py
```

## Input File Format

**Excel File**: `Navis_KKS_List.xlsx`
## Known Limitations

- Only supports Excel files with KKS codes in the first column and no header.
- Output XML structure is fixed for Navisworks; custom requirements may need script changes.
- Large lists of KKS codes may impact performance.

**Example Excel content:**
```
/0AUX-P-INN/PW
/0AUX-P-INN/TW
/0AUX-P-INN/WW
/1AUX-P-INN/CW-FH
```

## Output

**Generated XML file**: `YYMMDD_navis_filter.xml` (e.g., `250811_navis_filter.xml`)

The XML file contains:
- Proper Navisworks XML schema declarations
- Filter conditions for each KKS code using "contains" logic
- First condition with `flags="10"` (AND logic)
- Subsequent conditions with `flags="74"` (OR logic)
- Clean, formatted XML structure

**Example output structure:**
```xml
<?xml version='1.0' encoding='utf-8'?>
<exchange xmlns:xsi="..." units="m" filename="" filepath="">
  <findspec mode="all" disjoint="0">
    <conditions>
      <condition test="contains" flags="10">
        <property><name internal="LcOaSceneBaseUserName">Name</name></property>
        <value><data type="wstring">/0AUX-P-INN/PW</data></value>
      </condition>
      <condition test="contains" flags="74">
        <property><name internal="LcOaSceneBaseUserName">Name</name></property>
        <value><data type="wstring">/0AUX-P-INN/TW</data></value>
      </condition>
      <!-- Additional conditions... -->
    </conditions>
    <locator>/</locator>
  </findspec>
</exchange>
```

## Using in Navisworks

1. Open your Navisworks file
2. Go to **Home → Find Items** (or press Ctrl+F)
3. In the Find Items dialog, click **Import**
4. Select the generated XML file (`YYMMDD_navis_filter.xml`)
5. Click **Find All** to apply the filter
6. Objects matching the KKS codes will be selected/highlighted

## Error Handling

The script includes comprehensive error checking:
- **Missing virtual environment**: Checks for `..\.venv\Scripts\Activate.ps1`
- **Missing Excel file**: Verifies `Navis_KKS_List.xlsx` exists
- **Empty Excel file**: Warns if no KKS codes found
- **XML validation**: Confirms generated XML is well-formed
- **File generation**: Confirms output file was created successfully

## Features

### Data Processing
- ✅ Removes all whitespace from KKS codes (including internal spaces)
- ✅ Filters out empty/null entries
- ✅ Processes unlimited number of KKS codes
- ✅ Preserves exact KKS code formatting after cleaning

### XML Generation
- ✅ Follows Navisworks XML schema specification
- ✅ Implements proper filter logic (AND/OR conditions)
- ✅ Generates formatted, readable XML
- ✅ Validates XML structure automatically

### User Experience
- ✅ Automatic virtual environment activation
- ✅ Clear success/failure messages
- ✅ Progress reporting (number of codes processed)
- ✅ Date-stamped output files
- ✅ "Press any key to exit" for batch file usage

## Troubleshooting

**"Virtual environment not found"**
- Ensure the virtual environment is set up in `..\.venv\`
- Run from the correct directory (Navisworks_xml_filter_gen)

**"Excel file not found"**
- Ensure `Navis_KKS_List.xlsx` is in the same directory as the script
- Check file name spelling and extension

**"No KKS codes found"**
- Check that KKS codes are in the first column (Column A)
- Ensure cells contain text/values (not formulas)
- Verify the Excel file isn't corrupted

**"XML validation failed"**
- Contact support - this indicates an internal script issue

## Examples

### Sample Input (Excel)
```
Column A:
/0AUX-P-INN/PW
/0 AUX-P-INN/TW    (spaces will be removed)
/0AUX-P-INN/WW

(empty row - will be ignored)
/1AUX-P-INN/CW-FH
```

### Sample Output (Console)
```
========================================
Navisworks XML Filter Generator
========================================

Activating virtual environment...

Found 4 KKS codes
XML validation: PASSED
XML file generated successfully: C:\...\250811_navis_filter.xml
Total KKS codes processed: 4

========================================
Script execution completed
========================================

SUCCESS: XML file generated - 250811_navis_filter.xml

Press any key to exit...
```

## Version History

- **v1.0**: Initial release with basic XML generation
- **v1.1**: Added virtual environment support and error handling
- **v1.2**: Added batch file automation and improved user experience
- **v1.3**: Added whitespace cleaning and XML validation
