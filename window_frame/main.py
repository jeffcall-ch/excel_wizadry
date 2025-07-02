import tkinter as tk
from tkinter import ttk
import win32gui
import win32con
import win32api
import threading
import time

class WindowFramer:
    """
    An application to create a colored, resizable frame around a selected window.
    This version uses a robust recreation strategy and careful thread management.
    """
    FRAME_THICKNESS = 4
    COLORS = {
        "Red": "#FF0000", "Blue": "#0000FF", "Green": "#00FF00",
        "Yellow": "#FFFF00", "Cyan": "#00FFFF", "Magenta": "#FF00FF",
        "Black": "#000000", "White": "#FFFFFF",
    }
    FRAME_CLASS_NAME = "WindowFramerHighlightFrame"

    def __init__(self):
        # --- Core App State ---
        self.root = tk.Tk()
        self.target_hwnd = None
        self.frame_windows = []
        self.brush = None

        # --- Threading and Synchronization ---
        self.tracking_thread = None
        self.stop_event = threading.Event()
        self.recreate_event = threading.Event()
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

        select_button = ttk.Button(main_frame, text="Select Window to Frame", command=self._start_window_selection)
        select_button.pack(pady=10)

        color_label = ttk.Label(main_frame, text="Frame Color:")
        color_label.pack()

        color_menu = ttk.Combobox(main_frame, textvariable=self.selected_color_name, values=list(self.COLORS.keys()))
        color_menu.pack()
        color_menu.bind("<<ComboboxSelected>>", self._on_color_select)

        stop_button = ttk.Button(main_frame, text="Stop Framing", command=self.stop_tracking)
        stop_button.pack(pady=10)

    def run(self):
        """Starts the main application loop."""
        self.root.mainloop()

    def _get_toplevel_window(self, hwnd):
        """Gets the top-level parent of a window handle."""
        parent = win32gui.GetParent(hwnd)
        while parent:
            hwnd = parent
            parent = win32gui.GetParent(hwnd)
        return hwnd

    def _start_window_selection(self):
        """Starts a non-blocking thread to handle window selection."""
        threading.Thread(target=self._select_window_and_track, daemon=True).start()

    def _select_window_and_track(self):
        """Waits for a mouse click, captures the window, and starts tracking."""
        self.root.withdraw()
        time.sleep(0.5)

        print("Click on the window you want to frame...")
        while True:
            if win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000:
                break
            time.sleep(0.01)

        pos = win32gui.GetCursorPos()
        hwnd = win32gui.WindowFromPoint(pos)
        target_hwnd = self._get_toplevel_window(hwnd)

        self.root.deiconify()

        if not hwnd or target_hwnd == self.root.winfo_id() or target_hwnd == win32gui.GetDesktopWindow():
            print("Cannot frame this window. Please select another.")
            return

        self.target_hwnd = target_hwnd
        window_text = win32gui.GetWindowText(self.target_hwnd)
        print(f"Selected window: '{window_text}' (HWND: {self.target_hwnd})")

        self._start_tracking_thread()

    def _recreate_frame_windows(self):
        """Atomically destroys, unregisters, registers, and creates frame windows."""
        with self.frame_lock:
            for fw in self.frame_windows:
                if win32gui.IsWindow(fw):
                    win32gui.DestroyWindow(fw)
            self.frame_windows.clear()

            try:
                win32gui.UnregisterClass(self.FRAME_CLASS_NAME, None)
            except win32gui.error:
                pass

            if self.brush:
                win32gui.DeleteObject(self.brush)

            color_hex = self.COLORS[self.selected_color_name.get()]
            color_ref = int(color_hex[1:], 16)
            bgr_color = ((color_ref & 0xFF) << 16) | (color_ref & 0xFF00) | ((color_ref & 0xFF0000) >> 16)
            self.brush = win32gui.CreateSolidBrush(bgr_color)

            wc = win32gui.WNDCLASS()
            wc.lpszClassName = self.FRAME_CLASS_NAME
            wc.hbrBackground = self.brush
            wc.lpfnWndProc = lambda h, m, w, l: win32gui.DefWindowProc(h, m, w, l)
            win32gui.RegisterClass(wc)

            for _ in range(4):
                hwnd = win32gui.CreateWindowEx(
                    win32con.WS_EX_TOOLWINDOW | win32con.WS_EX_LAYERED | win32con.WS_EX_TOPMOST,
                    self.FRAME_CLASS_NAME, None,
                    win32con.WS_POPUP | win32con.WS_VISIBLE,
                    0, 0, 1, 1, None, None, win32api.GetModuleHandle(None), None
                )
                win32gui.SetLayeredWindowAttributes(hwnd, 0, 255, win32con.LWA_ALPHA)
                self.frame_windows.append(hwnd)

    def _track_window_loop(self):
        """Main loop to update the frame's position and size."""
        self.recreate_event.set() # Initial creation
        
        while not self.stop_event.is_set():
            try:
                if self.recreate_event.is_set():
                    self._recreate_frame_windows()
                    self.recreate_event.clear()

                if not self.target_hwnd or not win32gui.IsWindow(self.target_hwnd):
                    break

                rect = win32gui.GetWindowRect(self.target_hwnd)
                x, y, w, h = rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]

                with self.frame_lock:
                    if not all(win32gui.IsWindow(fw) for fw in self.frame_windows):
                        break
                    
                    z_order = win32con.HWND_TOPMOST
                    
                    win32gui.SetWindowPos(self.frame_windows[0], z_order, x, y, w, self.FRAME_THICKNESS, win32con.SWP_NOACTIVATE)
                    win32gui.SetWindowPos(self.frame_windows[1], z_order, x, y + h - self.FRAME_THICKNESS, w, self.FRAME_THICKNESS, win32con.SWP_NOACTIVATE)
                    win32gui.SetWindowPos(self.frame_windows[2], z_order, x, y, self.FRAME_THICKNESS, h, win32con.SWP_NOACTIVATE)
                    win32gui.SetWindowPos(self.frame_windows[3], z_order, x + w - self.FRAME_THICKNESS, y, self.FRAME_THICKNESS, h, win32con.SWP_NOACTIVATE)
            except win32gui.error:
                break
            
            time.sleep(0.016)
        
        # --- Thread-Safe Cleanup ---
        # This code runs inside the tracking thread when the loop exits.
        self.root.after(0, self.stop_tracking)


    def _start_tracking_thread(self):
        """Starts or restarts the window tracking thread."""
        if self.tracking_thread and self.tracking_thread.is_alive():
            self.stop_event.set()
            self.tracking_thread.join()

        self.stop_event.clear()
        self.tracking_thread = threading.Thread(target=self._track_window_loop, daemon=True)
        self.tracking_thread.start()

    def stop_tracking(self):
        """Signals the tracking thread to stop and cleans up resources."""
        if self.stop_event.is_set():
            return
        print("Stopping tracking...")
        self.stop_event.set()
        if self.tracking_thread and self.tracking_thread.is_alive():
            self.tracking_thread.join() # Wait for thread to finish cleanup

        with self.frame_lock:
            for fw in self.frame_windows:
                if win32gui.IsWindow(fw):
                    win32gui.DestroyWindow(fw)
            self.frame_windows.clear()
        
        self.target_hwnd = None

    def _on_color_select(self, event=None):
        """Signals the tracking thread to recreate the frames with a new color."""
        if self.tracking_thread and self.tracking_thread.is_alive():
            self.recreate_event.set()

    def _on_closing(self):
        """Handles cleanup when the main window is closed."""
        print("Application closing...")
        self.stop_tracking()
        try:
            win32gui.UnregisterClass(self.FRAME_CLASS_NAME, None)
            if self.brush:
                win32gui.DeleteObject(self.brush)
        except win32gui.error:
            pass
        self.root.destroy()

if __name__ == "__main__":
    app = WindowFramer()
    app.run()
