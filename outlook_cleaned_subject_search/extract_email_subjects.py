"""
Extract email subjects from .msg files in a folder and save to CSV.
This script reads Outlook .msg files and extracts their subjects for analysis.
"""

import os
import csv
from pathlib import Path

try:
    import win32com.client
except ImportError:
    print("Please install pywin32: pip install pywin32")
    exit(1)


def extract_subjects_from_folder(folder_path, output_csv="email_subjects.csv"):
    """
    Extract subjects from all .msg files in the specified folder.
    
    Args:
        folder_path: Path to folder containing .msg files
        output_csv: Output CSV file name
    """
    folder = Path(folder_path)
    
    if not folder.exists():
        print(f"Error: Folder not found: {folder_path}")
        return
    
    # Find all .msg files
    msg_files = list(folder.glob("*.msg"))
    
    if not msg_files:
        print(f"No .msg files found in {folder_path}")
        return
    
    print(f"Found {len(msg_files)} .msg files")
    
    # Initialize Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    subjects = []
    errors = []
    
    for msg_file in msg_files:
        try:
            # Open the .msg file
            msg = outlook.Session.OpenSharedItem(str(msg_file.absolute()))
            
            subjects.append({
                'filename': msg_file.name,
                'subject': msg.Subject,
                'sender': msg.SenderName,
                'date': str(msg.ReceivedTime) if hasattr(msg, 'ReceivedTime') else str(msg.SentOn)
            })
            
            print(f"✓ {msg_file.name}")
            
        except Exception as e:
            error_msg = f"✗ {msg_file.name}: {str(e)}"
            print(error_msg)
            errors.append(error_msg)
    
    # Write to CSV
    if subjects:
        output_path = folder / output_csv
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
            fieldnames = ['filename', 'subject', 'sender', 'date']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            writer.writerows(subjects)
        
        print(f"\n✓ Extracted {len(subjects)} subjects to: {output_path}")
        print(f"✗ {len(errors)} errors")
    else:
        print("\nNo subjects extracted.")
    
    if errors:
        print("\nErrors:")
        for error in errors:
            print(f"  {error}")


if __name__ == "__main__":
    folder_path = r"C:\temp\Outlook test"
    extract_subjects_from_folder(folder_path)
