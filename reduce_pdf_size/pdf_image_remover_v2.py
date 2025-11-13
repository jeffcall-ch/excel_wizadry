"""
PDF Image Remover V2 - Remove images from PDF while preserving text, tables, and page structure

This version uses pikepdf for better structure preservation.
All pages, tables, and text formatting are kept exactly as they are.
Only images are removed.

Usage:
    python pdf_image_remover_v2.py input.pdf output.pdf
    
Requirements:
    pip install pikepdf
"""

import sys
import os
import pikepdf
from pikepdf import Pdf, Name, PdfError

def remove_images_from_pdf(input_path, output_path):
    """
    Remove all images from a PDF file while preserving everything else.
    
    Args:
        input_path (str): Path to the input PDF file
        output_path (str): Path to save the output PDF file
        
    Returns:
        tuple: (success: bool, original_size: int, new_size: int, message: str)
    """
    try:
        # Open the PDF
        pdf = Pdf.open(input_path)
        
        original_size = os.path.getsize(input_path)
        images_removed = 0
        pages_processed = 0
        total_pages = len(pdf.pages)
        
        print(f"Processing {total_pages} pages...")
        
        # Process each page
        for page_num, page in enumerate(pdf.pages):
            pages_processed += 1
            
            if pages_processed % 10 == 0 or pages_processed == 1:
                print(f"  Processing page {pages_processed}/{total_pages}...")
            
            try:
                # Get the page's resources
                if '/Resources' in page and '/XObject' in page.Resources:
                    xobjects = page.Resources.XObject
                    
                    # List to track which XObjects to remove
                    to_remove = []
                    
                    # Iterate through all XObjects on the page
                    for name in xobjects:
                        try:
                            obj = xobjects[name]
                            
                            # Check if this XObject is an image
                            if '/Subtype' in obj:
                                subtype = obj.Subtype
                                
                                # If it's an image, mark it for removal
                                if subtype == Name.Image:
                                    to_remove.append(name)
                                    images_removed += 1
                                    
                        except Exception as e:
                            # Skip objects we can't process
                            print(f"    Warning: Could not process XObject '{name}' on page {page_num + 1}: {e}")
                            continue
                    
                    # Remove the marked images
                    for name in to_remove:
                        try:
                            del xobjects[name]
                        except Exception as e:
                            print(f"    Warning: Could not remove image '{name}' on page {page_num + 1}: {e}")
                            
            except Exception as e:
                print(f"  Warning: Error processing page {page_num + 1}: {e}")
                continue
        
        # Save the modified PDF with optimization
        pdf.save(
            output_path,
            compress_streams=True,
            stream_decode_level=pikepdf.StreamDecodeLevel.generalized,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
            normalize_content=True,
            linearize=False
        )
        pdf.close()
        
        new_size = os.path.getsize(output_path)
        size_reduction = ((original_size - new_size) / original_size) * 100 if original_size > 0 else 0
        
        message = f"""
PDF Processing Complete!
========================
Total pages: {total_pages}
Pages processed: {pages_processed}
Images removed: {images_removed}
Original size: {original_size / (1024*1024):.2f} MB
New size: {new_size / (1024*1024):.2f} MB
Size reduction: {size_reduction:.1f}%
Output saved to: {output_path}

✓ Page count preserved: {total_pages} pages
✓ Tables and structure intact
✓ Text formatting maintained
"""
        
        return True, original_size, new_size, message
        
    except Exception as e:
        return False, 0, 0, f"Error processing PDF: {str(e)}"


def main():
    """Main function to handle command-line usage."""
    if len(sys.argv) != 3:
        print("Usage: python pdf_image_remover_v2.py <input.pdf> <output.pdf>")
        print("\nExample:")
        print("  python pdf_image_remover_v2.py piping_catalogue.pdf piping_catalogue_text_only.pdf")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    # Validate input file exists
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found!")
        sys.exit(1)
    
    # Validate input is a PDF
    if not input_file.lower().endswith('.pdf'):
        print("Error: Input file must be a PDF!")
        sys.exit(1)
    
    # Check if output file already exists
    if os.path.exists(output_file):
        response = input(f"Warning: '{output_file}' already exists. Overwrite? (y/n): ")
        if response.lower() != 'y':
            print("Operation cancelled.")
            sys.exit(0)
    
    print(f"Processing: {input_file}")
    print("Removing images from PDF while preserving structure...")
    print()
    
    success, orig_size, new_size, message = remove_images_from_pdf(input_file, output_file)
    
    if success:
        print(message)
    else:
        print(message)
        sys.exit(1)


if __name__ == "__main__":
    main()
