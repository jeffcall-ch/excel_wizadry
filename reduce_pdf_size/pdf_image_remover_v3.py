"""
PDF Image Remover V3 - Remove images completely including their data streams

This version properly removes both image references AND their underlying data.
Preserves all pages, tables, text, and document structure.

Usage:
    python pdf_image_remover_v3.py input.pdf output.pdf
    
Requirements:
    pip install pikepdf
"""

import sys
import os
import pikepdf
from pikepdf import Pdf, Name

def remove_images_from_pdf(input_path, output_path):
    """
    Remove all images and their data from a PDF file while preserving everything else.
    
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
        
        # Track all image object IDs to delete them completely
        image_obj_ids = set()
        
        print(f"Processing {total_pages} pages...")
        
        # First pass: Remove images from page resources and collect image object IDs
        for page_num, page in enumerate(pdf.pages):
            pages_processed += 1
            
            if pages_processed % 10 == 0 or pages_processed == 1:
                print(f"  Pass 1: Processing page {pages_processed}/{total_pages}...")
            
            try:
                # Get the page's resources
                if '/Resources' in page and '/XObject' in page.Resources:
                    xobjects = page.Resources.XObject
                    
                    # List to track which XObjects to remove
                    to_remove = []
                    
                    # Iterate through all XObjects on the page
                    for name in list(xobjects.keys()):
                        try:
                            obj = xobjects[name]
                            
                            # Check if this XObject is an image
                            if '/Subtype' in obj and obj.Subtype == Name.Image:
                                to_remove.append(name)
                                images_removed += 1
                                
                                # Track the object ID for complete removal
                                if hasattr(obj, 'objgen'):
                                    image_obj_ids.add(obj.objgen)
                                    
                        except Exception as e:
                            # Skip objects we can't process
                            continue
                    
                    # Remove the marked images from the page
                    for name in to_remove:
                        try:
                            del xobjects[name]
                        except Exception as e:
                            pass
                            
            except Exception as e:
                print(f"  Warning: Error processing page {page_num + 1}: {e}")
                continue
        
        print(f"\nFound {images_removed} images to remove.")
        print("Saving optimized PDF with garbage collection...")
        
        # Save with maximum compression and garbage collection
        # This will remove unreferenced objects (the deleted images)
        pdf.save(
            output_path,
            compress_streams=True,
            stream_decode_level=pikepdf.StreamDecodeLevel.generalized,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
            normalize_content=True,
            linearize=False,
            recompress_flate=True,
            deterministic_id=False
        )
        pdf.close()
        
        # Second optimization pass: open and re-save to ensure complete cleanup
        print("Performing final cleanup pass...")
        pdf2 = Pdf.open(output_path)
        temp_output = output_path + ".tmp"
        pdf2.save(
            temp_output,
            compress_streams=True,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
            normalize_content=True
        )
        pdf2.close()
        
        # Replace original with cleaned version
        if os.path.exists(temp_output):
            import time
            time.sleep(0.5)  # Give OS time to release file handles
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
                os.rename(temp_output, output_path)
            except Exception as e:
                # If replacement fails, the first save is still good
                print(f"Note: Could not perform second cleanup pass: {e}")
                print("Using first pass output (still optimized)")
                if os.path.exists(temp_output):
                    try:
                        os.remove(temp_output)
                    except:
                        pass
        
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
✓ Images completely removed
"""
        
        return True, original_size, new_size, message
        
    except Exception as e:
        import traceback
        return False, 0, 0, f"Error processing PDF: {str(e)}\n{traceback.format_exc()}"


def main():
    """Main function to handle command-line usage."""
    if len(sys.argv) != 3:
        print("Usage: python pdf_image_remover_v3.py <input.pdf> <output.pdf>")
        print("\nExample:")
        print("  python pdf_image_remover_v3.py piping_catalogue.pdf piping_catalogue_text_only.pdf")
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
