# Excel Line List Comparator

A high-performance Python tool for comparing Excel piping line lists, generating detailed change analysis with visual formatting.

## 🚀 Quick Start

```bash
# Run comparison with auto-generated timestamped filename
python excel_compare.py compare.xlsx

# Or specify custom output filename
python excel_compare.py compare.xlsx my_result.xlsx
```

## 📁 Core Files

- **`excel_compare.py`** - Single file with complete comparison engine and CLI
- **`compare.xlsx`** - Sample input file with 'old' and 'new' sheets
- **`VBA_compare.vba`** - Original VBA reference code
- **`README.md`** - This documentation

## ✨ Features

- **KKS-based comparison** - Uses KKS column as unique identifier
- **Auto-generated filenames** - Timestamped outputs prevent conflicts
- **High performance** - Processes 1,600+ rows in ~8 seconds
- **Visual change tracking** - Colors, comments, and change markers
- **Single output sheet** - 'new_color_coded' sheet with NEW as base

## 📊 Output Format

The tool generates a single Excel sheet (`new_color_coded`) containing:
- **NEW table as base** - All new piping specifications
- **Color-coded changes** - Visual indicators for modifications
- **Change comments** - Detailed change descriptions
- **Added rows section** - New piping lines at bottom
- **Change markers** - Summary indicators per row

## 🎯 Example Output Filename

Input: `compare.xlsx`  
Output: `compare_20251007_143820_comparison.xlsx`

Pattern: `{filename}_{YYYYMMDD_HHMMSS}_comparison.xlsx`

## 📋 Requirements

- Python 3.12+
- pandas, openpyxl (see requirements.txt in parent directory)
- Input Excel file must contain 'old' and 'new' sheets

## 🔧 Technical Details

- **Single file solution**: All functionality in one Python file for simplicity
- **Algorithm**: Optimized O(n) comparison using vectorized pandas operations
- **Memory efficient**: Processes large datasets without memory issues
- **VBA compatible**: Produces identical results to original VBA code
- **Error handling**: Robust validation and cleanup
- **Dependencies**: pandas, openpyxl (standard Excel processing libraries)

## 📈 Performance

- **~1,600 rows**: 8 seconds
- **Small datasets**: <1 second
- **Memory usage**: Efficient for large piping datasets
- **Optimization**: 12x faster than original implementation

---

**Note**: This tool replaces the original VBA Excel comparison with a high-performance Python implementation while maintaining identical output results.