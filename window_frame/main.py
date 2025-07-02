import tkinter as tk
from tkinter import ttk
import ctypes
import win32gui
import win32con
import win32api
import threading
import time

# --- Configuration ---
FRAME_THICKNESS = 4
COLORS = {
    "Red": "#FF0000",
    "Blue": "#0000FF",
    "Green": "#00FF00",
    "Yellow": "#FFFF00",
    "Cyan": "#00FFFF",
    "Magenta": "#FF00FF",
    "Black": "#000000",
    "White": "#FFFFFF",
}

# --- Global State ---
target_hwnd = None
frame_windows = []
tracking_active = False
selected_color = tk.StringVar(value="#FF0000")
root = None


def get_toplevel_window(hwnd):
    """Gets the top-level parent of a window handle."""
    parent = win32gui.GetParent(hwnd)
    while parent:
        hwnd = parent
        parent = win32gui.GetParent(hwnd)
    return hwnd


def select_window():
    """Waits for a mouse click and captures the window handle."""
    global target_hwnd, tracking_active

    # Hide the main window to get it out of the way
    if root:
        root.withdraw()
    time.sleep(0.5)  # Give it time to disappear

    print("Click on the window you want to frame...")

    # Wait for a left mouse button click
    while True:
        if win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000:
            break
        time.sleep(0.01)

    # Get window handle from mouse position
    pos = win32gui.GetCursorPos()
    hwnd = win32gui.WindowFromPoint(pos)
    target_hwnd = get_toplevel_window(hwnd)

    # Don't allow framing the script's own window or the desktop
    if target_hwnd == root.winfo_id() or target_hwnd == win32gui.GetDesktopWindow():
        print("Cannot frame this window. Please select another.")
        target_hwnd = None
        if root:
            root.deiconify()
        return

    window_text = win32gui.GetWindowText(target_hwnd)
    print(f"Selected window: '{window_text}' (HWND: {target_hwnd})")

    # Stop any previous tracking and start new tracking
    if tracking_active:
        stop_tracking()

    tracking_active = True
    threading.Thread(target=track_window, daemon=True).start()

    # Bring the control window back
    if root:
        root.deiconify()


def create_frame_windows():
    """Creates the four top-level windows that form the frame."""
    global frame_windows
    # Destroy old frames if they exist
    for fw in frame_windows:
        win32gui.DestroyWindow(fw)
    frame_windows = []

    # Create 4 borderless, transparent, top-most windows
    for i in range(4):
        wc = win32gui.WNDCLASS()
        wc.lpszClassName = f"FrameWindow{i}"
        wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wc.hbrBackground = win32api.GetStockObject(win32con.BLACK_BRUSH)
        wc.lpfnWndProc = lambda hwnd, msg, wparam, lparam: win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
        class_atom = win32gui.RegisterClass(wc)

        hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_TOOLWINDOW | win32con.WS_EX_LAYERED,
            class_atom,
            None,  # No window title
            win32con.WS_POPUP | win32con.WS_VISIBLE,
            0, 0, 1, 1,  # Initial position and size
            None,
            None,
            win32api.GetModuleHandle(None),
            None,
        )
        # Set transparency (255 is fully opaque)
        win32gui.SetLayeredWindowAttributes(hwnd, 0, 255, win32con.LWA_ALPHA)
        frame_windows.append(hwnd)


def track_window():
    """Main loop to update the frame's position and size."""
    global tracking_active

    create_frame_windows()
    last_color = None

    while tracking_active:
        if not win32gui.IsWindow(target_hwnd):
            print
