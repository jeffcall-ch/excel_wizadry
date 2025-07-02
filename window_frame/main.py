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
selected_color = None
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
        if win32gui.IsWindow(fw):
            win32gui.DestroyWindow(fw)
    frame_windows = []

    # Create 4 borderless, transparent, top-most windows
    for i in range(4):
        wc = win32gui.WNDCLASS()
        wc.lpszClassName = f"FrameWindow{i}"
        wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        # Create a solid brush for the frame color
        color_ref = int(selected_color.get()[1:], 16)
        # BGR format for CreateSolidBrush
        bgr_color = ((color_ref & 0xFF) << 16) | (color_ref & 0xFF00) | ((color_ref & 0xFF0000) >> 16)
        wc.hbrBackground = win32gui.CreateSolidBrush(bgr_color)
        wc.lpfnWndProc = lambda hwnd, msg, wparam, lparam: win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
        
        try:
            class_atom = win32gui.RegisterClass(wc)
        except win32gui.error as e:
            if e.winerror == 1410: # Class already exists
                pass
            else:
                raise e

        hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_TOOLWINDOW | win32con.WS_EX_LAYERED | win32con.WS_EX_TOPMOST,
            wc.lpszClassName,
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

    while tracking_active:
        if not win32gui.IsWindow(target_hwnd):
            print("Target window closed. Exiting.")
            stop_tracking()
            root.quit()
            return

        # Get the window rect of the target
        try:
            rect = win32gui.GetWindowRect(target_hwnd)
        except win32gui.error:
            print("Target window not found. Exiting.")
            stop_tracking()
            root.quit()
            return

        x, y, w, h = rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]

        # --- Position the 4 frame windows ---
        # Top
        win32gui.SetWindowPos(frame_windows[0], win32con.HWND_TOPMOST, x, y, w, FRAME_THICKNESS, win32con.SWP_NOACTIVATE)
        # Bottom
        win32gui.SetWindowPos(frame_windows[1], win32con.HWND_TOPMOST, x, y + h - FRAME_THICKNESS, w, FRAME_THICKNESS, win32con.SWP_NOACTIVATE)
        # Left
        win32gui.SetWindowPos(frame_windows[2], win32con.HWND_TOPMOST, x, y, FRAME_THICKNESS, h, win32con.SWP_NOACTIVATE)
        # Right
        win32gui.SetWindowPos(frame_windows[3], win32con.HWND_TOPMOST, x + w - FRAME_THICKNESS, y, FRAME_THICKNESS, h, win32con.SWP_NOACTIVATE)

        # Set the owner of the frame windows to the target window
        for fw in frame_windows:
            win32gui.SetWindowLong(fw, win32con.GWL_HWNDPARENT, target_hwnd)


        time.sleep(0.01)


def stop_tracking():
    """Stops the tracking loop and destroys the frame windows."""
    global tracking_active
    tracking_active = False
    for fw in frame_windows:
        if win32gui.IsWindow(fw):
            win32gui.DestroyWindow(fw)
    frame_windows.clear()


def create_gui():
    """Creates the main control GUI."""
    global root, selected_color
    root = tk.Tk()
    root.title("Window Framer")
    root.geometry("300x150")

    selected_color = tk.StringVar(value="#FF0000")

    main_frame = ttk.Frame(root, padding="10")
    main_frame.pack(fill="both", expand=True)

    # --- Widgets ---
    select_button = ttk.Button(main_frame, text="Select Window to Frame", command=select_window)
    select_button.pack(pady=10)

    color_label = ttk.Label(main_frame, text="Frame Color:")
    color_label.pack()

    color_menu = ttk.Combobox(main_frame, textvariable=selected_color, values=list(COLORS.keys()))
    color_menu.pack()
    # Set the default value
    color_menu.set(list(COLORS.keys())[0])


    def on_color_select(event):
        color_name = color_menu.get()
        selected_color.set(COLORS[color_name])
        if tracking_active:
            create_frame_windows()

    color_menu.bind("<<ComboboxSelected>>", on_color_select)


    stop_button = ttk.Button(main_frame, text="Stop Framing", command=stop_tracking)
    stop_button.pack(pady=10)


    root.protocol("WM_DELETE_WINDOW", root.quit)
    root.mainloop()

if __name__ == "__main__":
    create_gui()