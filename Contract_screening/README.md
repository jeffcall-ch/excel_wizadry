# Contract Screening Tool

## Overview
This tool automatically screens large PDF contracts (1000+ pages) for piping-related keywords and generates comprehensive Excel reports. Perfect for EPC projects where you need to quickly identify all piping-related content scattered throughout massive contract documents.

## Features
- **PDF Text Extraction**: Processes native PDF files efficiently
- **Keyword Matching**: Finds keywords as substrings (case-insensitive)
- **Context Capture**: Extracts 5 words before and after each match
- **Smart Scoring**: Pages scored based on keyword frequency + diversity bonus
- **Multiple Excel Sheets**:
  - All Results: Complete list of matches with context
  - Category Summary: Statistics by keyword category
  - High-Scoring Pages: Top 20% of pages or those with score ≥ 5
  - Statistics: Overall analysis and top keywords

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Command Line
```bash
python contract_screening.py <pdf_file> <keywords_csv> [-o output.xlsx]
```

### Example
```bash
python contract_screening.py "contract.pdf" "piping_contract_search_phrases_for_piping_and_supports.txt" -o "screening_results.xlsx"
```

### Keywords CSV Format
Your keywords file should have this structure:
```csv
Category,Keyword
General Piping & Scope,piping
General Piping & Scope,piping system
Supports & Hangers,pipe support
Materials & Components,valves
...
```

## Scoring System
- **Base Score**: 1 point per keyword match
- **Diversity Bonus**: 0.5 points per unique keyword on the page
- **Example**: Page with "piping" (3 times) + "valves" (2 times) = 5 + 1 = 6 points

## Output Excel Structure

### Sheet 1: All Results
| Page | Keyword | Category | Context |
|------|---------|----------|---------|
| 45 | piping | General Piping & Scope | "system shall include piping layout drawings and specifications" |

### Sheet 2: Category Summary
| Category | Total Matches | Pages with Matches | Avg Matches per Page |
|----------|---------------|-------------------|---------------------|
| General Piping & Scope | 156 | 78 | 2.0 |

### Sheet 3: High-Scoring Pages
| Page | Score | Total Matches | Unique Keywords | Categories Found |
|------|-------|---------------|-----------------|------------------|
| 234 | 12.5 | 10 | 5 | General Piping & Scope, Materials & Components |

### Sheet 4: Statistics
- Total pages with matches
- Total keyword matches
- Unique keywords found
- Top 10 most frequent keywords

## Performance Notes
- Processes ~50-100 pages per minute (depending on PDF complexity)
- Memory usage scales with PDF size
- Progress logging every 50 pages during extraction, 100 pages during search

## Troubleshooting

### Common Issues
1. **"PDF file not found"**: Check file path and ensure PDF exists
2. **"Keywords CSV file not found"**: Verify CSV file path and format
3. **Memory errors**: For very large PDFs, consider processing in batches

### PDF Requirements
- Native PDF files with selectable text (not scanned images)
- If you have scanned PDFs, convert them first using OCR tools

## Customization Options

You can modify the script to:
- Change context window size (currently 5 words before/after)
- Adjust scoring algorithm
- Add fuzzy matching for keyword variations
- Include/exclude specific categories
- Change Excel formatting and styling

## Example Use Case
Perfect for annual contract reviews in EPC projects where you need to:
- Identify all piping scope mentions
- Find specification requirements
- Locate interface points and tie-ins
- Review quality and inspection requirements
- Extract maintenance and commissioning details

---
*Tool developed for piping engineers working on large EPC energy-to-waste projects*