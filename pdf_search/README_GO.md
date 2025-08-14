# PDF Search Tool (Go Version)

This is a Go rewrite of the Python PDF search tool. It searches for text in PDF files within a directory and generates an Excel report with the results.

## Features

- Recursive search through directories for PDF files
- Case-sensitive and case-insensitive search options
- Context extraction around matches (50 characters before and after)
- Excel report generation with hyperlinks to source files
- Summary sheet with search statistics
- Command-line interface with flags

## Prerequisites

- Go 1.21 or later
- Internet connection for downloading dependencies

## Installation

1. Clone or download the files to your local machine
2. Navigate to the directory containing the Go files
3. Initialize the module and download dependencies:

```bash
go mod tidy
```

## Usage

### Basic Usage

```bash
go run multi_pdf_full_text_search.go -search="text to find" -dir="/path/to/pdfs"
```

### Advanced Usage

```bash
# Case-sensitive search with custom output file
go run multi_pdf_full_text_search.go -search="Error" -dir="./documents" -case-sensitive -output="my_results.xlsx"

# Search in a specific directory
go run multi_pdf_full_text_search.go -search="failure" -dir="C:\temp\rev1_isos"
```

### Command Line Options

- `-search` (required): Text to search for in PDF files
- `-dir` (required): Directory containing PDF files to search
- `-output` (optional): Output file path for Excel report (default: auto-generated with timestamp)
- `-case-sensitive` (optional): Enable case-sensitive search
- `-help` (optional): Show help message

### Help

```bash
go run multi_pdf_full_text_search.go -help
```

## Building

To create a standalone executable:

```bash
# For Windows
go build -o pdf_search.exe multi_pdf_full_text_search.go

# For Linux/Mac
go build -o pdf_search multi_pdf_full_text_search.go
```

Then run the executable directly:

```bash
# Windows
.\pdf_search.exe -search="text" -dir="./pdfs"

# Linux/Mac
./pdf_search -search="text" -dir="./pdfs"
```

## Output

The tool generates an Excel file with two sheets:

1. **Search Results**: Contains all matches with:
   - File Path (with hyperlink to original file)
   - File Name
   - Page Number
   - Match Context (text surrounding the match)

2. **Summary**: Contains search statistics:
   - Date of search
   - Search term used
   - Total PDF files searched
   - Total matches found
   - Report generator information

## Differences from Python Version

The Go version maintains the same functionality as the Python version with these improvements:

1. **Performance**: Generally faster execution due to Go's compiled nature
2. **Dependencies**: Self-contained executable after building (no need for Python runtime)
3. **Flag-based CLI**: Uses Go's standard flag package for cleaner command-line interface
4. **Memory efficiency**: Better memory management for large PDF files
5. **Cross-platform**: Easy to build for different operating systems

## Dependencies

- `github.com/ledongthuc/pdf`: For PDF text extraction
- `github.com/xuri/excelize/v2`: For Excel file generation

## Troubleshooting

1. **Module not found errors**: Run `go mod tidy` to download dependencies
2. **Permission errors**: Ensure you have read access to PDF files and write access to output directory
3. **Large files**: The tool loads entire PDF pages into memory; very large files may cause memory issues
4. **Corrupted PDFs**: Some damaged PDF files may cause the tool to skip pages or files entirely

## Performance Notes

- The tool processes PDFs sequentially
- Large directories with many PDFs will take time to process
- Memory usage scales with PDF file sizes
- Progress is displayed as files are being processed
