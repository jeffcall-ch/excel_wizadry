# window_frame

## Purpose
This folder contains a Tkinter-based Windows utility for creating a colored, resizable frame around a selected window.
- `main.py`: Launches a GUI that lets you select a window and apply a colored frame for visual highlighting.

**Input:**
- User selection of target window and frame color.

**Output:**
- Visual frame drawn around the selected window (on Windows OS).

## Usage
1. Install dependencies:
   ```powershell
   pip install pywin32
   ```
2. Run the script from the command line:
   ```powershell
   python main.py
   ```

## Examples
```powershell
python main.py
```

## Known Limitations
- Windows OS only (uses win32 APIs).
- Requires admin rights for some window operations.
- Only works with windows that are not minimized or hidden.
- Frame color and thickness are configurable in the GUI.
