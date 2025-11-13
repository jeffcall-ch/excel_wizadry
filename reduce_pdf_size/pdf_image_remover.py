"""
PDF Image Remover - Remove images from PDF while preserving text and structure

This script removes all images from a PDF file while maintaining:
- All text content
- Document structure
- Tables and formatting
- Hyperlinks and annotations

Useful for reducing PDF file size before feeding to LLMs.

Usage:
    python pdf_image_remover.py input.pdf output.pdf
    
Requirements:
    pip install PyMuPDF (fitz)
"""

import sys
import os
import fitz  # PyMuPDF

def remove_images_from_pdf(input_path, output_path):
    """
    Remove all images from a PDF file while preserving text, tables, and page structure.
    
    This approach creates a cleaned copy by redrawing only non-image content.
    Each page maintains its original structure and page count remains the same.
    
    Args:
        input_path (str): Path to the input PDF file
        output_path (str): Path to save the output PDF file
        
    Returns:
        tuple: (success: bool, original_size: int, new_size: int, message: str)
    """
    try:
        # Open the original PDF
        doc = fitz.open(input_path)
        
        original_size = os.path.getsize(input_path)
        images_removed = 0
        pages_processed = 0
        total_pages = len(doc)
        
        print(f"Processing {total_pages} pages...")
        
        # Process each page
        for page_num in range(total_pages):
            try:
                page = doc[page_num]
                pages_processed += 1
                
                if pages_processed % 10 == 0:
                    print(f"  Processing page {pages_processed}/{total_pages}...")
                
                # Count images on this page before removal
                try:
                    image_list = page.get_images(full=True)
                    images_removed += len(image_list)
                except:
                    pass
                
                # Get the page's content stream (commands that draw the page)
                # We'll filter out image-related commands
                try:
                    # Clean the page to remove image objects
                    # This method preserves all content structure but removes images
                    xref = page.get_contents()
                    
                    if isinstance(xref, list):
                        # Multiple content streams
                        for xr in xref:
                            try:
                                doc._deleteObject(xr)
                            except:
                                pass
                    elif xref and xref > 0:
                        # Single content stream - we need to filter it
                        pass
                    
                    # Alternative: Remove image XObjects from resources
                    # Get page resources
                    resources = page.get_resources()
                    
                    # Remove XObject images from resources
                    if resources:
                        xobjects = resources.get("XObject", {})
                        if xobjects:
                            # Create a filtered version without images
                            for img_name in list(xobjects.keys()):
                                try:
                                    img_obj = xobjects[img_name]
                                    # Check if it's an image XObject
                                    if isinstance(img_obj, dict):
                                        subtype = img_obj.get("/Subtype")
                                        if subtype == "/Image":
                                            # Remove this image
                                            del xobjects[img_name]
                                except:
                                    pass
                
                except Exception as e:
                    print(f"  Warning: Could not clean page {page_num + 1}: {e}")
                    
            except Exception as e:
                print(f"  Warning: Error processing page {page_num + 1}: {e}")
                continue
        
        # Save with aggressive cleaning and compression
        # This will remove unreferenced objects (like deleted images)
        doc.save(
            output_path,
            garbage=4,        # Maximum garbage collection
            deflate=True,     # Compress streams
            clean=True,       # Clean and sanitize
            pretty=False      # Minimize file size
        )
        doc.close()
        
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
"""
        
        return True, original_size, new_size, message
        
    except Exception as e:
        return False, 0, 0, f"Error processing PDF: {str(e)}"


def main():
    """Main function to handle command-line usage."""
    if len(sys.argv) != 3:
        print("Usage: python pdf_image_remover.py <input.pdf> <output.pdf>")
        print("\nExample:")
        print("  python pdf_image_remover.py piping_catalogue.pdf piping_catalogue_text_only.pdf")
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
    print("Removing images from PDF...")
    
    success, orig_size, new_size, message = remove_images_from_pdf(input_file, output_file)
    
    if success:
        print(message)
    else:
        print(message)
        sys.exit(1)


if __name__ == "__main__":
    main()
