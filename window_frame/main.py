import tkinter as tk
from tkinter import ttk
import ctypes
import win32gui
import win32con
import win32api
import threading
import time

class WindowFramer:
    """
    An application to create a colored, resizable frame around a selected window.
    It uses tkinter for the GUI and pywin32 for Windows API interaction.
    """
    FRAME_THICKNESS = 4
    COLORS = {
        "Red": "#FF0000", "Blue": "#0000FF", "Green": "#00FF00",
        "Yellow": "#FFFF00", "Cyan": "#00FFFF", "Magenta": "#FF00FF",
        "Black": "#000000", "White": "#FFFFFF",
    }
    FRAME_CLASS_NAME = "HighlightFrameWindow"

    def __init__(self):
        # --- Core App State ---
        self.root = tk.Tk()
        self.target_hwnd = None
        self.frame_windows = []

        # --- Threading and Synchronization ---
        self.tracking_thread = None
        self.stop_event = threading.Event()
        self.frame_lock = threading.Lock()

        # --- GUI Variables ---
        self.selected_color_name = tk.StringVar(value=list(self.COLORS.keys())[0])
        
        self._configure_root()
        self._create_gui()

    def _configure_root(self):
        """Configures the main tkinter window."""
        self.root.title("Window Framer")
        self.root.geometry("300x150")
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _create_gui(self):
        """Creates and lays out the widgets for the control panel."""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)

        select_button = ttk.Button(main_frame, text="Select Window to Frame", command=self._select_window_thread)
        select_button.pack(pady=10)

        color_label = ttk.Label(main_frame, text="Frame Color:")
        color_label.pack()

        color_menu = ttk.Combobox(main_frame, textvariable=self.selected_color_name, values=list(self.COLORS.keys()))
        color_menu.pack()
        color_menu.bind("<<ComboboxSelected>>", self._on_color_select)

        stop_button = ttk.Button(main_frame, text="Stop Framing", command=self.stop_tracking)
        stop_button.pack(pady=10)

    def run(self):
        """Registers resources and starts the main application loop."""
        self._register_frame_class()
        self.root.mainloop()

    def _get_toplevel_window(self, hwnd):
        """Gets the top-level parent of a window handle."""
        parent = win32gui.GetParent(hwnd)
        while parent:
            hwnd = parent
            parent = win32gui.GetParent(hwnd)
        return hwnd

    def _select_window_thread(self):
        """
        Starts a non-blocking thread to handle window selection to avoid freezing the GUI.
        """
        threading.Thread(target=self._select_window, daemon=True).start()

    def _select_window(self):
        """Waits for a mouse click and captures the window handle."""
        self.root.withdraw()
        time.sleep(0.5)  # Allow time for the window to disappear

        print("Click on the window you want to frame...")
        while True:
            if win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000:
                break
            time.sleep(0.01)

        pos = win32gui.GetCursorPos()
        hwnd = win32gui.WindowFromPoint(pos)
        target_hwnd = self._get_toplevel_window(hwnd)

        self.root.deiconify() # Bring the control window back immediately

        if target_hwnd == self.root.winfo_id() or target_hwnd == win32gui.GetDesktopWindow():
            print("Cannot frame this window. Please select another.")
            return

        self.target_hwnd = target_hwnd
        window_text = win32gui.GetWindowText(self.target_hwnd)
        print(f"Selected window: '{window_text}' (HWND: {self.target_hwnd})")

        self.start_tracking()

    def _create_frame_windows(self):
        """Creates the four top-level windows that form the frame."""
        with self.frame_lock:
            # Clean up any existing frame windows
            for fw in self.frame_windows:
                if win32gui.IsWindow(fw):
                    win32gui.DestroyWindow(fw)
            self.frame_windows = []

            # Create 4 new borderless, top-most windows
            for _ in range(4):
                hwnd = win32gui.CreateWindowEx(
                    win32con.WS_EX_TOOLWINDOW | win32con.WS_EX_LAYERED,
                    self.FRAME_CLASS_NAME,
                    None,
                    win32con.WS_POPUP | win32con.WS_VISIBLE,
                    0, 0, 1, 1,
                    None, None, win32api.GetModuleHandle(None), None
                )
                win32gui.SetLayeredWindowAttributes(hwnd, 0, 255, win32con.LWA_ALPHA)
                self.frame_windows.append(hwnd)
        self._update_frame_color()

    def _update_frame_color(self):
        """Updates the background color of the frame windows."""
        color_hex = self.COLORS[self.selected_color_name.get()]
        color_ref = int(color_hex[1:], 16)
        bgr_color = ((color_ref & 0xFF) << 16) | (color_ref & 0xFF00) | ((color_ref & 0xFF0000) >> 16)
        
        try:
            brush = win32gui.CreateSolidBrush(bgr_color)
            with self.frame_lock:
                for fw in self.frame_windows:
                    if win32gui.IsWindow(fw):
                        win32gui.SetClassLong(fw, win32con.GCL_HBRBACKGROUND, brush)
                        win32gui.InvalidateRect(fw, None, True) # Force redraw
        except win32gui.error as e:
            print(f"Error updating frame color: {e}")


    def _track_window_loop(self):
        """Main loop to update the frame's position and size."""
        self._create_frame_windows()

        while not self.stop_event.is_set():
            if not self.target_hwnd or not win32gui.IsWindow(self.target_hwnd):
                print("Target window closed or lost. Stopping.")
                self.root.after(0, self.stop_tracking)
                break

            try:
                rect = win32gui.GetWindowRect(self.target_hwnd)
                x, y, w, h = rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]

                with self.frame_lock:
                    if not all(win32gui.IsWindow(fw) for fw in self.frame_windows):
                        print("Frame windows were destroyed. Stopping.")
                        break
                    
                    # Set Z-order to be just above the target window
                    target_z = win32gui.GetWindow(self.target_hwnd, win32con.GW_HWNDPREV)

                    # Top, Bottom, Left, Right
                    win32gui.SetWindowPos(self.frame_windows[0], target_z, x, y, w, self.FRAME_THICKNESS, win32con.SWP_NOACTIVATE)
                    win32gui.SetWindowPos(self.frame_windows[1], target_z, x, y + h - self.FRAME_THICKNESS, w, self.FRAME_THICKNESS, win32con.SWP_NOACTIVATE)
                    win32gui.SetWindowPos(self.frame_windows[2], target_z, x, y, self.FRAME_THICKNESS, h, win32con.SWP_NOACTIVATE)
                    win32gui.SetWindowPos(self.frame_windows[3], target_z, x + w - self.FRAME_THICKNESS, y, self.FRAME_THICKNESS, h, win32con.SWP_NOACTIVATE)

            except win32gui.error:
                print("Error getting window rect. Target likely closed.")
                self.root.after(0, self.stop_tracking)
                break
            
            time.sleep(0.016) # ~60 FPS

    def start_tracking(self):
        """Starts the window tracking thread."""
        if self.tracking_thread and self.tracking_thread.is_alive():
            self.stop_event.set()
            self.tracking_thread.join()

        self.stop_event.clear()
        self.tracking_thread = threading.Thread(target=self._track_window_loop, daemon=True)
        self.tracking_thread.start()

    def stop_tracking(self):
        """Stops the tracking loop and destroys the frame windows."""
        print("Stopping tracking...")
        self.stop_event.set()
        self.target_hwnd = None
        with self.frame_lock:
            for fw in self.frame_windows:
                if win32gui.IsWindow(fw):
                    win32gui.DestroyWindow(fw)
            self.frame_windows.clear()

    def _on_color_select(self, event=None):
        """Callback for when a new color is selected from the dropdown."""
        print(f"Color changed to {self.selected_color_name.get()}")
        if self.tracking_thread and self.tracking_thread.is_alive():
            self._update_frame_color()

    def _register_frame_class(self):
        """Registers the window class for the frames."""
        wc = win32gui.WNDCLASS()
        wc.lpszClassName = self.FRAME_CLASS_NAME
        wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wc.hbrBackground = win32api.GetStockObject(win32con.BLACK_BRUSH)
        wc.lpfnWndProc = lambda hwnd, msg, wparam, lparam: win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
        try:
            win32gui.RegisterClass(wc)
        except win32gui.error as e:
            if e.winerror != 1410: # 1410: Class already exists
                raise

    def _unregister_frame_class(self):
        """Unregisters the window class."""
        try:
            win32gui.UnregisterClass(self.FRAME_CLASS_NAME, None)
        except win32gui.error:
            pass # Ignore if not registered

    def _on_closing(self):
        """Handles cleanup when the main window is closed."""
        print("Application closing...")
        self.stop_tracking()
        self._unregister_frame_class()
        self.root.destroy()

if __name__ == "__main__":
    app = WindowFramer()
    app.run()
