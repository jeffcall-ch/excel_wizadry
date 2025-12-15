"""
Outlook Email Search Application - Phase 1
Main entry point

This application searches Outlook emails with various criteria.
WARNING: This is a READ-ONLY application. It will never modify any Outlook data.
"""

import tkinter as tk
from tkinter import messagebox
import sys
from search_gui import SearchGUI


def main():
    """Main entry point for the application."""
    try:
        # Create main window
        root = tk.Tk()
        
        # Set window icon (optional - uses default if not available)
        try:
            root.iconbitmap('outlook_search.ico')
        except:
            pass  # Icon file not found, use default
        
        # Create and run the application
        app = SearchGUI(root)
        
        # Center window on screen
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        # Start the GUI event loop
        root.mainloop()
        
    except Exception as e:
        messagebox.showerror("Application Error", 
                           f"An error occurred starting the application:\n{str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
