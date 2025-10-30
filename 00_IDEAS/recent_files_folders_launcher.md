“Recent Files / Folders Launcher (WinForms, single file)”

> Build a tiny **Windows desktop launcher** in **C# WinForms (.NET 8)** named **Recent Launcher** that opens frequently used **files or folders** with one click.
> **Constraints:** single self-contained file `Program.cs` (no designer, no external libs). Must compile/run with `dotnet run`.
>
> **Functional requirements**
>
> 1. **List UI**: Main window shows a ListBox of bookmarks. Each item stores: `DisplayName`, `FullPath`, `Type` (File/Folder).
> 2. **Buttons**:
>
>    * **Add…** opens a dialog: user can choose **either** a folder (FolderBrowserDialog) **or** a file (OpenFileDialog via dropdown choice).
>    * **Open** launches the selected item:
>
>      * Folder → open in File Explorer.
>      * File → open with default associated app.
>    * **Remove** deletes the selected bookmark from the list.
>    * **Open location** (optional) opens the parent folder and selects the file if it’s a file.
> 3. **Double-click** on a list item = same as **Open**.
> 4. **Keyboard shortcuts**:
>
>    * Enter = Open, Delete = Remove, Ctrl+N = Add, Ctrl+O = Open, Ctrl+L = Open location, Ctrl+S = Save.
> 5. **Persistence**: Bookmarks are saved to a small JSON file at `%APPDATA%\RecentLauncher\bookmarks.json`.
>
>    * Load on startup (if missing, start empty).
>    * Save on every add/remove and on form close.
> 6. **Validation & UX**:
>
>    * If a saved path no longer exists, show status “(missing)” next to DisplayName in the list, and disable **Open** for that item.
>    * Prevent duplicates (same FullPath); if user tries to add one, show a friendly message.
>    * Support **drag-and-drop** from Explorer onto the list to add items (file or folder).
>    * Allow **Rename Display Name** via F2 or context menu (“Rename…” → simple input box).
> 7. **Layout**:
>
>    * Top: Title label “Recent Launcher”.
>    * Center: ListBox (fills window).
>    * Bottom: Buttons **Add**, **Open**, **Open location**, **Remove**, and a right-aligned label with count (`N items`).
>    * Window size ~ 700×450, remember last window size/position across runs (store in same JSON as `windowBounds`).
> 8. **Error handling**: Wrap Process.Start calls; show MessageBox on errors (e.g., missing permissions).
> 9. **Code quality**:
>
>    * All in one file (`Program.cs`).
>    * Small `Bookmark` record/class `{ string DisplayName; string FullPath; string Type; }`.
>    * Helper methods: `LoadBookmarks()`, `SaveBookmarks()`, `RefreshList()`, `OpenSelected()`, `AddPath(string)`, `RemoveSelected()`, `RenameSelected()`.
>    * Comments explaining key sections so I can narrate live.
>
> **Deliverable**
>
> * Output only the full contents of `Program.cs`.
> * No placeholders; fully runnable.
> * Use `System.Text.Json` for persistence and `Process.Start(new ProcessStartInfo { UseShellExecute = true, FileName = … })` for launching.
> * Ensure drag-and-drop is enabled and handled (files and folders).
> * Add a simple context menu on the ListBox with **Open**, **Open location**, **Rename**, **Remove**.


