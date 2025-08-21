# python_package_documentation_crawler

## Purpose
This folder contains a tool for crawling and aggregating documentation for Python packages.
- `documentation_crawler.py`: Given a list of Python package dependencies, finds their official documentation, crawls the site, and compiles the content into Markdown files for each package.

**Input:**
- List of Python package names or dependency strings.

**Output:**
- Markdown files containing the full documentation for each package.

## Usage
1. Install dependencies:
   ```powershell
   pip install requests beautifulsoup4 markdownify
   ```
2. Run the script from the command line:
   ```powershell
   python documentation_crawler.py <package_list.txt>
   ```
   (See script for argument details.)

## Examples
```powershell
python documentation_crawler.py requirements.txt
```

## Known Limitations
- Relies on the structure of external documentation sites; may break if sites change.
- Rate limits or blocks may occur for large crawls.
- Only supports Python packages with public documentation.
