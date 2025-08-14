import PyPDF2
import sys

if len(sys.argv) < 2:
    print("Usage: python debug_pdf.py <pdf_file>")
    sys.exit(1)

pdf_path = sys.argv[1]

try:
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        num_pages = len(reader.pages)
        
        print(f"PDF: {pdf_path}")
        print(f"Number of pages: {num_pages}")
        
        if num_pages > 0:
            page = reader.pages[0]
            text = page.extract_text()
            
            print(f"\nRaw text length: {len(text)} characters")
            print(f"First 500 characters:")
            print(repr(text[:500]))
            
            # Check for "pip" (case insensitive)
            if "pip" in text.lower():
                print("\n✓ Found 'pip' in the text!")
                
                # Show context around first occurrence
                index = text.lower().find("pip")
                start = max(0, index-50)
                end = min(len(text), index+53)
                print(f"Context: ...{text[start:end]}...")
            else:
                print("\n✗ 'pip' not found in the text.")
                
except Exception as e:
    print(f"Error processing {pdf_path}: {str(e)}")
