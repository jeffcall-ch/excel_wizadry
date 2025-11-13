# PDF Image Remover

Remove images from PDF files while preserving text, tables, and document structure. Perfect for reducing PDF file sizes before feeding them to LLMs.

## Features

- ✅ Removes all images from PDF files
- ✅ Preserves all text content
- ✅ Maintains document structure
- ✅ Keeps tables and formatting intact
- ✅ Retains hyperlinks and annotations
- ✅ Significant file size reduction (typically 50-90% for image-heavy PDFs)

## Installation

Install the required dependency:

```powershell
pip install PyMuPDF
```

## Usage

### Basic Usage

```powershell
python pdf_image_remover.py input.pdf output.pdf
```

### Example

For your piping catalogue:

```powershell
python pdf_image_remover.py piping_catalogue.pdf piping_catalogue_text_only.pdf
```

### Expected Results

For an 80MB+ PDF with images:
- **Original size**: 80+ MB
- **New size**: 5-20 MB (depending on text/image ratio)
- **Size reduction**: 70-95%

## How It Works

The script uses PyMuPDF (fitz) to:

1. Open the PDF file
2. Iterate through all pages
3. Identify and remove all embedded images
4. Clean and compress the PDF
5. Save the optimized version

All text content, including:
- Regular text
- Tables
- Headers and footers
- Annotations
- Hyperlinks
- Bookmarks

...remains completely intact.

## Notes

- The output PDF will be significantly smaller but may look different visually (no images)
- Text extraction by LLMs will work perfectly with the optimized PDF
- The original file is never modified - a new file is created
- Processing time depends on the number of pages and images

## Troubleshooting

If you encounter any issues:

1. **Import Error**: Make sure PyMuPDF is installed: `pip install PyMuPDF`
2. **File Not Found**: Check that the input path is correct
3. **Permission Error**: Ensure you have write access to the output location
4. **Memory Issues**: For very large PDFs, you may need to close other applications

## Alternative: Batch Processing

If you need to process multiple PDFs, you can modify the script or run it in a loop:

```powershell
# Process all PDFs in a folder
Get-ChildItem -Filter *.pdf | ForEach-Object {
    python pdf_image_remover.py $_.Name "optimized_$($_.Name)"
}
```
