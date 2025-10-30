# QR Label Splitter

A Python tool that extracts individual labels from a multi-label PDF document, categorizes them based on search criteria, and outputs organized, optimized PDF files with searchable text.

## Overview

This application processes a source PDF containing multiple labels (e.g., QR codes, part labels, equipment tags) and:
- Extracts each label based on user-defined coordinates
- Categorizes labels based on configurable search strings
- Arranges labels in a clean 2-column layout on A4 pages
- Sorts labels naturally by BoM Item Number
- Removes invisible text that extends beyond label boundaries
- Outputs compressed, optimized PDFs with full text searchability

## Features

- ✅ **Smart Extraction**: Uses precise coordinates to extract labels from source PDF
- ✅ **Categorization**: Automatically sorts labels into separate PDFs based on filter strings
- ✅ **Clean Output**: Removes invisible text outside label boundaries while preserving searchability
- ✅ **Natural Sorting**: Sorts labels by BoM Item No. with natural number ordering
- ✅ **Optimized PDFs**: Uses compression and garbage collection for small file sizes
- ✅ **Verification**: Hashes all labels to ensure complete extraction without duplicates or missing items

## Requirements

### Software
- Python 3.9 or higher
- Virtual environment (recommended)

### Python Packages
- PyMuPDF (v1.26.4 or higher) - for PDF manipulation
- pandas - for Excel coordinate handling
- openpyxl - for reading .xlsx files

Install dependencies:
```bash
pip install PyMuPDF>=1.26.4 pandas openpyxl
```

## Required Input Files

### 1. Configuration File (`config.ini`)

Create a `config.ini` file in the same directory as the script with the following structure:

```ini
[Paths]
input_pdf = C:\path\to\your\source_labels.pdf
excel_coordinates = C:\path\to\your\coordinates.xlsx
output_dir = C:\path\to\output\directory

[Filters]
search_strings = FILTER1, FILTER2, FILTER3
```

**Configuration Parameters:**

- **`input_pdf`**: Full path to the source PDF containing multiple labels
- **`excel_coordinates`**: Full path to the Excel file with label coordinates
- **`output_dir`**: Directory where output PDFs will be saved
- **`search_strings`**: Comma-separated list of search terms for categorization
  - Labels containing these strings will be saved to separate PDFs
  - Labels not matching any filter go to `rest.pdf`
  - Example: `COR1_MK, COR1_ME` creates `COR1_MK.pdf`, `COR1_ME.pdf`, and `rest.pdf`

### 2. Coordinates Excel File (`coordinates.xlsx`)

An Excel file defining the rectangular boundaries of each label in the source PDF. The user must extract these coordinates manually from the PDF beforehand.

**Required columns:**
- **`points`**: Unique identifier for each coordinate point (e.g., 1001, 1002, 1101, 1102)
- **`X`**: X-coordinate in PDF points (1 point = 1/72 inch)
- **`Y`**: Y-coordinate in PDF points

**Structure:**
Each label requires **2 points** to define a rectangle (top-left and bottom-right corners):
- Points with matching prefix define one rectangle (e.g., 1001 & 1002 = Label 10)
- First 2 digits = label number, last digit = point number

**Example:**
```
points    X       Y
1001     22.32   22.32    # Label 10 - Top-left corner
1002    283.82  181.77    # Label 10 - Bottom-right corner
1101     22.32  182.77    # Label 11 - Top-left corner
1102    283.82  340.50    # Label 11 - Bottom-right corner
```

**How to Extract Coordinates:**
1. Open the source PDF in a PDF editor (Adobe Acrobat, PDF-XChange Editor, etc.)
2. Use measurement or annotation tools to identify label boundaries
3. Record the X,Y coordinates of top-left and bottom-right corners
4. Create the Excel file with the required format

### 3. Source PDF File

The input PDF containing all labels to be extracted. This file should contain multiple labels arranged on one or more pages.

## Usage

### Basic Usage

1. **Prepare your files:**
   - Source PDF with labels
   - Excel file with coordinates (extracted manually from PDF)
   - `config.ini` with paths and filter settings

2. **Run the script:**
   ```bash
   python extract_label.py
   ```

3. **Check output:**
   - Output PDFs will be created in the specified output directory
   - Each category gets its own PDF (e.g., `COR1_MK.pdf`, `COR1_ME.pdf`)
   - Uncategorized labels go to `rest.pdf`

### Example Workflow

```bash
# 1. Activate virtual environment (if using one)
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # Linux/Mac

# 2. Install dependencies
pip install PyMuPDF>=1.26.4 pandas openpyxl

# 3. Prepare config.ini and coordinates.xlsx

# 4. Run the extraction
python extract_label.py
```

## Output

The script generates categorized PDF files with the following characteristics:

- **Layout**: 2-column A4 format with 20-point margins
- **Sorting**: Labels sorted naturally by BoM Item Number
- **Optimization**: 
  - Garbage collection removes duplicate objects
  - Deflate compression reduces file size
  - Content streams cleaned and optimized
- **Quality**: Full text searchability preserved, invisible text removed

### Output Files

Based on filter configuration:
- `{FILTER_NAME}.pdf` - Labels matching each filter string
- `rest.pdf` - All labels not matching any filter

### Console Output

```
--- Hashing Source Labels ---
Found and hashed 380 unique labels from the source document.
Successfully created C:\output\rest.pdf
Successfully created C:\output\COR1_MK.pdf
Successfully created C:\output\COR1_ME.pdf

--- Verification ---
Success: All 380 labels from the source document have been successfully extracted and processed.
Verification complete.
```

## How It Works

1. **Configuration Loading**: Reads paths and filter settings from `config.ini`

2. **Coordinate Loading**: Parses Excel file to build label rectangles
   - Groups coordinate points by label ID
   - Creates Rect objects for each label boundary

3. **Source Hashing**: Creates SHA-256 hashes of all source labels for verification

4. **Extraction & Categorization**:
   - Reads text from each label using defined coordinates
   - Searches for filter strings in label text
   - Categorizes labels (matched filters → separate PDFs, rest → `rest.pdf`)

5. **Natural Sorting**: Sorts labels by BoM Item No. with intelligent number ordering

6. **Label Processing**:
   - Creates temporary page for each label
   - Clips content to exact label boundaries
   - **Uses `clip_to_rect()` to permanently remove text/graphics outside boundaries**
   - Stamps cleaned label onto output page in 2-column layout

7. **PDF Generation**:
   - Saves with garbage=4 (remove duplicates)
   - Applies deflate compression
   - Cleans content streams

8. **Verification**: Compares source and output hashes to ensure completeness

## Technical Details

### Text Boundary Cleaning

The script uses PyMuPDF's `clip_to_rect()` method (v1.26.4+) to permanently remove:
- Text characters extending beyond label boundaries
- Vector graphics outside the defined rectangle
- Images not fully contained in the label area

This ensures no "invisible" search hits appear when users search the output PDFs.

### File Size Optimization

Three optimization techniques are used:
- **`garbage=4`**: Highest level garbage collection, removes all duplicate/unused objects
- **`deflate=True`**: Compresses streams using deflate algorithm
- **`clean=True`**: Optimizes and reformats content streams

Typical results: 90-95% file size reduction compared to unoptimized output.

### Natural Sorting

Labels are sorted using a natural sorting algorithm that handles:
- Alphanumeric combinations (e.g., "Item10" comes after "Item2")
- Hyphenated numbers (e.g., "A-1-10" comes after "A-1-2")
- Mixed case strings

## Troubleshooting

### "AttributeError: 'Page' object has no attribute 'clip_to_rect'"
- **Cause**: PyMuPDF version too old
- **Solution**: Upgrade to v1.26.4 or higher: `pip install --upgrade PyMuPDF`

### Large Output File Sizes
- **Cause**: Missing optimization parameters in save()
- **Solution**: Ensure script uses `garbage=4, deflate=True, clean=True` in save()

### Missing Labels in Output
- **Cause**: Incorrect coordinates in Excel file
- **Solution**: Verify coordinates match actual label positions in source PDF

### Labels Not Categorizing Correctly
- **Cause**: Filter strings don't match label content
- **Solution**: Check search strings in `config.ini` match text in labels

### Invisible Text Still Present
- **Cause**: PyMuPDF version < 1.26.4
- **Solution**: Update PyMuPDF to latest version

## Version History

- **v1.0**: Initial release with clip_to_rect() method for clean text boundaries
- Optimized file sizes with garbage collection and compression
- Natural sorting by BoM Item Number
- Hash-based verification system

## License

[Add your license information here]

## Author

[Add your information here]
