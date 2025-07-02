# windowframe.py

import win32gui
import win32api
import win32con
import time
import tkinter as tk

def get_window_class_name(hwnd):
    """Get the class name of a window by its handle."""
    try:
        class_name = win32gui.GetClassName(hwnd)
        return class_name
    except:
        return ""

def get_window_title(hwnd):
    """Get the title of a window by its handle."""
    try:
        return win32gui.GetWindowText(hwnd)
    except:
        return ""

def is_window_foreground(hwnd):
    """Check if the window is the foreground window."""
    try:
        foreground_hwnd = win32gui.GetForegroundWindow()
        return hwnd == foreground_hwnd
    except:
        return False
        
def is_window_moving(hwnd, prev_rect):
    """Check if a window is being moved by comparing its current position with previous position."""
    try:
        current_rect = win32gui.GetWindowRect(hwnd)
        if prev_rect is None:
            return False
        return current_rect != prev_rect
    except:
        return False

def get_main_window_handle(hwnd):
    """Try to get the main window handle from a child window handle."""
    try:
        # Walk up the window hierarchy to find the main window
        while True:
            parent = win32gui.GetParent(hwnd)
            if not parent or parent == hwnd:
                break
            hwnd = parent
        return hwnd
    except:
        return hwnd

def enum_window_callback(hwnd, target_windows):
    """Callback for EnumWindows. Finds all visible top-level windows."""
    if win32gui.IsWindowVisible(hwnd):
        target_windows.append(hwnd)
    return True

def get_all_top_level_windows():
    """Returns all visible top-level windows."""
    windows = []
    win32gui.EnumWindows(enum_window_callback, windows)
    return windows

def find_outermost_window(hwnd):
    """
    Uses screen coordinates to find the outermost window that contains
    the provided window handle. This is more reliable than just using
    parent relationships, as some applications use complex window nesting.
    """
    # Get all top-level windows
    all_top_level = get_all_top_level_windows()
    
    if hwnd in all_top_level:
        # This is already a top-level window
        return hwnd
    
    # Get the screen coordinates of our target window
    try:
        target_rect = win32gui.GetWindowRect(hwnd)
        target_x, target_y = target_rect[0], target_rect[1]
        
        # Try the traditional parent approach first
        main_hwnd = get_main_window_handle(hwnd)
        
        # If the parent approach gave us a different window, and it's top-level, use that
        if main_hwnd != hwnd and main_hwnd in all_top_level:
            print(f"Found parent window with handle: {main_hwnd}")
            return main_hwnd
            
        # If parent approach didn't work, find a top-level window that contains these coordinates
        for window in all_top_level:
            try:
                rect = win32gui.GetWindowRect(window)
                # Check if this window contains our target point
                if (rect[0] <= target_x <= rect[2] and 
                    rect[1] <= target_y <= rect[3]):
                    # Check if window has a title (most application main windows do)
                    if win32gui.GetWindowText(window).strip():
                        print(f"Found containing window: {window} with title: {win32gui.GetWindowText(window)}")
                        return window
            except:
                continue
    except:
        pass
        
    # If all else fails, return the original handle
    return hwnd

def get_target_window_handle():
    print("Click inside the window you want to frame...")
    # Wait for the left mouse button to be pressed
    while True:
        if win32api.GetKeyState(0x01) < 0:  # 0x01 is the virtual key code for the left mouse button
            pos = win32gui.GetCursorPos()
            initial_handle = win32gui.WindowFromPoint(pos)
            
            # Get the class name and window title for debugging
            class_name = get_window_class_name(initial_handle)
            window_title = get_window_title(initial_handle)
            
            # Print debug information
            print(f"Initial window handle: {initial_handle}")
            print(f"Class name: {class_name}")
            print(f"Window title: {window_title}")
            
            # Always try to find the outermost containing window
            final_handle = find_outermost_window(initial_handle)
            
            if final_handle != initial_handle:
                final_class = get_window_class_name(final_handle)
                final_title = get_window_title(final_handle)
                print(f"Found outermost window with handle: {final_handle}")
                print(f"Outermost class: {final_class}")
                print(f"Outermost title: {final_title}")
            
            return final_handle
        time.sleep(0.1)  # Prevents high CPU usage

def create_frame_window(color, border_width=2):
    """
    Creates a transparent window with a colored frame border.
    
    Args:
        color: The color of the frame border
        border_width: The width of the frame border in pixels
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window initially
    
    # Create the main top-level window (this will be our frame)
    frame = tk.Toplevel(root)
    
    # Critical step: Make it a tool window which always stays on top
    frame.wm_attributes("-toolwindow", True)
    frame.overrideredirect(True)  # Remove title bar and borders
    
    # Set up the transparent color
    transparent_color = 'gray15'
    frame.attributes("-transparentcolor", transparent_color)
    
    # Force the window to stay on top of ALL other windows
    frame.attributes("-topmost", True)
    
    # Add Windows-specific extended window styles
    # After the window is created, we'll use win32 APIs to further adjust its behavior
    
    # Configure the frame with the chosen color
    frame.configure(bg=color)
    
    # Create an inner frame with the transparent color
    # This creates the "hollow" center
    inner_frame = tk.Frame(frame, bg=transparent_color)
    
    # Position the inner frame to create a border of exactly the specified width
    # Use place instead of pack for precise positioning
    inner_frame.place(x=border_width, y=border_width, 
                     relwidth=1, relheight=1, 
                     width=-2*border_width, height=-2*border_width)
                     
    # Update to ensure window is created
    frame.update_idletasks()

    return root, frame

def main():
    print("WindowFrame script started.")
    
    print("Please select a frame color:")
    colors = {
        "1": "Red", "2": "Green", "3": "Blue",
        "4": "Yellow", "5": "Purple", "6": "Cyan", "7": "White"
    }
    for key, value in colors.items():
        print(f"{key} - {value}")

    choice = input("Enter the number for your chosen color: ")
    selected_color = colors.get(choice)

    if not selected_color:
        print("Invalid choice. Exiting.")
        return
        
    print(f"You chose: {selected_color}")
    
    print("\nInstructions:")
    print("1. Move your mouse over the window you want to frame")
    print("2. Click once inside the window")
    print("3. Press Ctrl+C in this terminal window to exit when done\n")
    
    # Get the window handle from the user's click
    target_handle = get_target_window_handle()
    print(f"Window selected with handle: {target_handle}")
    
    # Create the frame window with the selected color
    border_width = 5  # Width of the colored border
    root, frame_window = create_frame_window(selected_color.lower(), border_width)
    
    try:
        # Force the window to the foreground initially to ensure it's created properly
        frame_window.lift()
        frame_window.attributes('-topmost', True)
        frame_window.update()
        
        # Wait a bit to make sure the window is created
        time.sleep(0.5)
        
        # Get the frame window's handle using more reliable methods
        frame_hwnd = None
        
        # Try multiple methods to get the correct window handle
        frame_hwnd = win32gui.FindWindow(None, frame_window.winfo_name())
        
        # If the above didn't work, try getting it from the window ID
        if not frame_hwnd:
            try:
                frame_id = frame_window.winfo_id()
                frame_hwnd = frame_id
            except:
                pass
                
        # Last resort fallback - sometimes works with tkinter
        if not frame_hwnd:
            frame_hwnd = win32gui.GetForegroundWindow()
            
        print(f"Frame window handle: {frame_hwnd}")
        
        # Apply additional window styles with Win32 API to ensure it stays on top
        if frame_hwnd:
            # Set extended window style to stay on top
            GWL_EXSTYLE = -20
            WS_EX_TOPMOST = 0x00000008  # Forces a top-most window
            WS_EX_TRANSPARENT = 0x00000020  # Makes the window transparent to mouse events
            WS_EX_LAYERED = 0x00080000  # Required for transparency effects
            WS_EX_NOACTIVATE = 0x08000000  # Prevents the window from becoming active
            
            # Combine the styles for our overlay frame
            ex_style = WS_EX_TOPMOST | WS_EX_NOACTIVATE
            
            # Apply the extended style to our window
            current_style = win32gui.GetWindowLong(frame_hwnd, GWL_EXSTYLE)
            win32gui.SetWindowLong(frame_hwnd, GWL_EXSTYLE, current_style | ex_style)
            
            # Force it to the top-most position initially
            win32gui.SetWindowPos(
                frame_hwnd, 
                win32con.HWND_TOPMOST, 
                0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
            )
        
        prev_rect = None
        
        # Main frame tracking loop
        print("Frame is now tracking the target window. Press Ctrl+C to exit.")
        while True:
            try:
                # Get the position and size of the target window
                current_rect = win32gui.GetWindowRect(target_handle)
                
                # Calculate dimensions
                x, y, width, height = current_rect[0], current_rect[1], current_rect[2] - current_rect[0], current_rect[3] - current_rect[1]
                
                # Position our frame window at the exact coordinates of the target window
                frame_window.geometry(f"{width}x{height}+{x}+{y}")
                
                # Keep the window on top with both tkinter and win32 methods for maximum reliability
                frame_window.attributes('-topmost', True)
                frame_window.lift()
                
                if frame_hwnd:
                    # Use Win32 API to ensure it stays topmost
                    win32gui.SetWindowPos(
                        frame_hwnd, 
                        win32con.HWND_TOPMOST, 
                        0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
                    )
                
                # Update the window to apply changes
                frame_window.update()
                prev_rect = current_rect
                
            except Exception as e:
                print(f"Error updating frame position: {e}")
                
            time.sleep(0.05)  # Slightly faster refresh for better movement tracking
    except KeyboardInterrupt:
        print("Script stopped by user.")
    except Exception as e:
        # This handles the case where the target window is closed
        print(f"Error: {e}")
        print("Target window closed or an error occurred. Exiting.")
    finally:
        root.destroy()

if __name__ == "__main__":
    main()
