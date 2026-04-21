# FileExplorer — Product Requirements Document

**Version**: 1.0  
**Date**: April 2026  
**Audience**: Designers, product owners, anyone evaluating or rebuilding the application.  
This document describes what the user sees and how they work with the application. It contains no implementation details.

---

## 1. Overview

FileExplorer is a **Windows desktop file manager** designed for engineers working with large collections of Office documents, PDFs, and project files. Its key differentiator over Windows Explorer is that it **reads inside files** and surfaces document metadata — project name and revision — directly as sortable, scannable columns in the file list, without the user having to open each file.

The interface follows the visual language of Windows 11 (Mica blur, Fluent icons, native window chrome) and feels native to the OS.

---

## 2. Who Uses It and Why

**Primary user**: A project engineer or document controller who navigates between many project folders daily, needs to quickly identify which revision of a drawing or specification is current, and moves or copies files between folders routinely.

**Core pain points solved**:

| Pain point | How the app addresses it |
|---|---|
| "Which revision is this file?" requires opening it | Revision column shows inline — no opening required |
| "Which project does this belong to?" | Project column shows the extracted project name |
| Slow to switch between several project folders | Browser-style tabs — each folder is its own tab |
| Windows Explorer forgets sort and column state per folder | Per-folder sort preferences are remembered |
| Searching inside a deep folder tree is slow | Fast full-text filename search (sub-second) across entire subtree |

---

## 3. Application Shell

### 3.1 Window

The application opens as a standard resizable Windows 11 window with a native title bar. The title bar merges into the tab strip — the tabs run along the very top of the window, and the drag region for moving the window sits around the tabs.

The window title shows the current folder name followed by `- FileExplorer`.

### 3.2 Zones (top to bottom)

```
┌──────────────────────────────────────────────────────────────────┐
│  [Tab 1]  [Tab 2]  [+ New Tab]         [— □ ×]  ← title bar     │
├──────────────────────────────────────────────────────────────────┤
│  ◄ ► ▲ ⟳  │  C: › Projects › Current Job ›  │  🔍 Search...     │
├──────────────────────────────────────────────────────────────────┤
│                                         │                         │
│         File list (main area)           │  Inspector sidebar      │
│                                         │  (320 px)               │
├──────────────────────────────────────────────────────────────────┤
│  12 items · 3 selected         5.4 GB free on C:                 │
└──────────────────────────────────────────────────────────────────┘
```

The **inspector sidebar** is always visible on the right. There is no left-side tree panel visible by default.

---

## 4. Tab Bar

Each open folder gets its own tab. Tabs look and behave like a browser's tab strip.

### 4.1 What a tab shows

- **Tab label**: the folder's name (e.g., `Current Job`).
- **Tooltip on hover**: the full path.
- **Close button** (×): appears on each tab unless the tab is pinned.
- **Pinned tab**: narrower; no close button; pin icon shown instead of folder name.

### 4.2 Tab actions (right-click context menu on a tab)

- Close tab
- Close other tabs
- Close tabs to the right
- Duplicate tab (opens a new tab at the same folder)
- Pin / Unpin tab

### 4.3 Keyboard shortcuts

| Action | Keys |
|---|---|
| New tab | Ctrl+T |
| Close current tab | Ctrl+W |
| Next tab | Ctrl+Tab |
| Previous tab | Ctrl+Shift+Tab |
| Jump to tab 1–8 | Ctrl+1 through Ctrl+8 |
| Jump to last tab | Ctrl+9 |

### 4.4 Minimum tab behaviour

- The last remaining tab cannot be closed.
- Pinned tabs cannot be closed from the × button or "close" menu items.

---

## 5. Navigation Bar

The navigation bar sits below the tab strip. It has three parts:

### 5.1 Navigation buttons (left group)

| Button | Action |
|---|---|
| ◄ Back | Go back to the previous folder in this tab's history |
| ► Forward | Go forward to the next folder in this tab's history |
| ▲ Up | Go to the parent folder |
| ⟳ Refresh | Reload the current folder |

Buttons are greyed out when the action is unavailable (e.g., Back when at the start of history).

### 5.2 Address bar / breadcrumb (centre, fills remaining width)

**Normal state — breadcrumb display**:

The current path is shown as clickable segments separated by `›` chevrons:

```
  C:  ›  Projects  ›  Current Job  ›  Site Drawings  ›
```

- Clicking any **segment** navigates directly to that folder.
- Clicking a **`›` chevron** opens a dropdown list of sibling folders at that level — the user can jump sideways in the tree without going up and back down.
- There is always a trailing `›` after the last segment.
- If the path is long the bar scrolls horizontally; the rightmost (current) segment is always visible.

**Edit state — typed path**:

Clicking on the blank area of the address bar (between or around the segments) switches to an editable text field showing the full path, with all text selected ready to overtype.

- As the user types, folder names are suggested in a dropdown (autocomplete from real file system).
- Press **Enter** or select a suggestion to navigate.
- Press **Escape** to cancel and return to breadcrumb display.
- Clicking outside the bar (losing focus) also cancels and goes back to breadcrumb display.

### 5.3 Search box (right, fixed width)

A search field labelled `Search...`. Typing and pressing Enter searches for files anywhere inside the current folder and all subfolders. See §9 for full search behaviour.

---

## 6. File List

The file list occupies the main central area of the window. It displays the contents of the current folder as a details (columns) view.

### 6.1 Columns

| Column | Default width | What it shows |
|---|---|---|
| **Name** | wide | File or folder name |
| **Project** | medium | Project name extracted from the file's content |
| **Rev** | narrow | Revision identifier extracted from the file's content |
| **Date modified** | medium | Last-modified date and time |
| **Size** | narrow | File size (blank for folders) |

The Project and Rev columns are blank for unsupported file types and for folders.

### 6.2 Row appearance

- Rows are compact (short height) — more files fit on screen at once.
- File icons (16 px) appear at the left of the Name column.
- **File-name colour coding** (dark mode): file names are coloured by type to allow rapid scanning without reading extensions:
  - PDF → red
  - Excel files → green
  - Word files → blue
  - Archives (.zip, .7z) → yellow
  - Text / Markdown → teal
  - Everything else → default text colour
  - In light mode, only PDF files are coloured; all others use default text colour.
- Hidden files and system files are shown or hidden based on the current setting.
- File extensions in the Name column can be shown or hidden.

### 6.3 Date group dividers

When sorted by Date modified, files are grouped under dividers such as **Today**, **Yesterday**, **This week**, **Last month**, etc. The divider rows are styled differently (smaller text, subtle colour) and are not selectable.

### 6.4 Sorting

Clicking a **column header** sorts by that column. Clicking again reverses the sort direction. A small chevron in the header shows the active sort column and direction.

Sort preferences are saved per folder — returning to a folder restores the sort you used there last time.

### 6.5 Selection

- Click to select one item.
- Ctrl+Click to add/remove individual items.
- Shift+Click to select a range.
- Ctrl+A to select all.
- Clicking on blank space deselects.

### 6.6 Opening items

- **Double-click a folder** — navigates into that folder in the current tab.
- **Double-click a file** — opens the file with its default application (same as Windows Explorer).
- **Shortcut (.lnk) files** — automatically resolved; navigates into the target folder or opens the target file.

### 6.7 Empty folder state

When a folder is empty, a large icon and a short message are shown centred in the list area.

### 6.8 Loading state

While a folder is loading, a circular progress indicator is shown centred in the list area. The previous folder's contents remain visible until the new contents are ready, preventing a blank-screen flash.

### 6.9 Processing banner

When the app is reading file contents to populate the Project/Rev columns, a small banner appears at the bottom-left of the file list showing a spinner and a short status message (e.g., *"Reading document metadata…"*). It disappears automatically when done.

---

## 7. Inspector Sidebar

The sidebar is a fixed 320 px panel on the right side of the window. It is always visible. It does not preview file contents — it provides navigation shortcuts.

### 7.1 Navigation hub (top section)

A list of pre-configured destination folders or quick-access locations. Each item shows a folder name and a right-arrow icon indicating it is a navigation target. Clicking an item navigates the current tab to that location.

### 7.2 Favourites (bottom section)

A list of folders the user has pinned. Behaviour is the same as navigation hub items. A bold section header labelled **Favourites** separates the two lists.

A 1 px horizontal divider sits between the hub and the favourites section with small top/bottom spacing.

---

## 8. Status Bar

The status bar runs along the bottom of the window (24 px tall).

- **Left side**: item count and selection summary, e.g. *"24 items"* or *"24 items · 3 selected (1.2 MB)"*.
- **Right side**: free space on the current drive, e.g. *"5.4 GB free on C:"*.

Updates instantly when selection changes or after any file operation.

---

## 9. Search

### 9.1 Scope

Search covers **filenames only** (not file contents). It searches recursively through all subfolders of the current directory.

### 9.2 Behaviour

1. The user types a query in the search box and presses Enter.
2. The file list switches to **search results mode**: the current folder listing is replaced by a flat list of all matching files from anywhere in the subtree, with their full paths visible.
3. A loading indicator is shown while the search runs.
4. Results appear as they are found (near-instant for fast queries).
5. Search is case-insensitive and supports glob patterns (e.g., `*.pdf`, `report*`).

### 9.3 Returning to folder view

- Clearing the search box and pressing Enter, or pressing Escape, returns the file list to the current directory listing.
- Navigating to a folder (address bar, breadcrumb, sidebar) also clears the search.

---

## 10. File Operations

### 10.1 Copy / Cut / Paste

| Action | Keyboard / Menu |
|---|---|
| Copy | Ctrl+C |
| Cut | Ctrl+X |
| Paste | Ctrl+V |

- After cutting/copying, pasting into another folder (another tab, or after navigating) places the files there.
- After a cut+paste, the source files disappear from the source folder immediately (no waiting for the operation to complete on disk).
- After a paste, the pasted files appear in the destination folder immediately, sorted into the correct position in the list.

### 10.2 Delete

| Action | Key |
|---|---|
| Delete (to Recycle Bin) | Delete |
| Permanent delete | Shift+Delete |

### 10.3 Rename

- Press **F2** with one item selected, or use the right-click context menu.
- The file name becomes an inline editable text field directly in the Name column.
- Press **Enter** to confirm, **Escape** to cancel.
- The item name updates visually as soon as Enter is pressed; the row moves to its new sorted position immediately.
- If the rename fails (e.g., name conflict), the name reverts and an error message is shown.

### 10.4 Undo

- Ctrl+Z undoes the last file operation (move, copy, rename, delete).
- Undo is available for one level.

### 10.5 Properties

- Alt+Enter or right-click → Properties opens the standard Windows file properties dialog.

### 10.6 Open in Terminal / other shell actions

Right-click context menu includes shell-integrated actions (exact set depends on any registered shell extensions).

---

## 11. Document Metadata Columns (Project & Rev)

### 11.1 What is extracted and from which files

The app reads the **first page** of supported files to find project name and revision information. Supported file types:

- Microsoft Word: `.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.dotm`
- Microsoft Excel: `.xls`, `.xlsx`, `.xlsm`, `.xlt`, `.xltx`, `.xltm`, `.xlsb`
- PDF: `.pdf`

For all other file types the Project and Rev columns are blank.

### 11.2 What the user sees

- On first visiting a folder containing supported files, the Project and Rev columns start blank and fill in progressively as the app reads each file in the background.
- The processing banner (§6.9) indicates this is happening.
- On returning to the same folder, all values are shown immediately with no loading delay, because results are cached locally.
- The cache is invalidated automatically when a file is modified (the app detects the changed date/size and re-reads it).

### 11.3 Revision fallback

If no explicit revision label is found inside the file, the app looks for a version number pattern (e.g., `2.1`, `4.0`) in the **filename itself** and uses that as the revision.

---

## 12. OneDrive / Cloud Status

When the current folder is inside a OneDrive sync folder, an additional visual indicator may appear on file icons showing their sync state:

| State | Meaning |
|---|---|
| Cloud-only | File exists in the cloud only; opening it will trigger a download |
| Locally available | File is synced and available offline |
| Always available | File is pinned to always be available offline |
| Syncing | File is in the process of being synced |

This column is hidden when the current folder is not under any detected OneDrive sync root.

---

## 13. Settings

Settings toggled from the menu or keyboard shortcuts:

| Setting | Default | Effect |
|---|---|---|
| Show hidden files | Off | Shows/hides files marked Hidden in their attributes |
| Show file extensions | On | Shows/hides the extension in the Name column (e.g., `.xlsx`) |
| Show preview pane | Off | Toggles the preview pane (note: current sidebar is the inspector, not the preview) |

Pane widths (sidebar, navigation panel) are remembered between sessions.

---

## 14. Session Restore

When the application is closed and reopened, it remembers which folder each tab was showing and restores the session. Tabs reopen in the same location as when the app was closed.

---

## 15. Single-instance Behaviour

If a second instance of the application is launched (e.g., by opening a folder from Windows Explorer's "Open with"), the existing window is brought to focus and the requested folder is opened in a new tab (or an existing tab showing that folder is activated). A second window is never opened.

---

## 16. Keyboard Shortcuts — Full Reference

| Action | Keys |
|---|---|
| New tab | Ctrl+T |
| Close tab | Ctrl+W |
| Next tab | Ctrl+Tab |
| Previous tab | Ctrl+Shift+Tab |
| Tab 1–8 | Ctrl+1 – Ctrl+8 |
| Last tab | Ctrl+9 |
| Back | Alt+← |
| Forward | Alt+→ |
| Up to parent | Alt+↑ |
| Refresh | F5 |
| Focus address bar | Alt+D or Ctrl+L |
| Rename | F2 |
| Copy | Ctrl+C |
| Cut | Ctrl+X |
| Paste | Ctrl+V |
| Delete | Delete |
| Permanent delete | Shift+Delete |
| Undo | Ctrl+Z |
| Select all | Ctrl+A |
| Properties | Alt+Enter |
| Search | Ctrl+F (or type in the search box) |
| Show/hide hidden files | Value not specified |
| Show/hide extensions | Value not specified |

---

## 17. Out of Scope (v1)

The following are **not** part of the current product:

- File preview panel (images, PDFs, text rendered inline)
- Editing file metadata (writing back project/revision to the file)
- Network drive browsing via UNC paths (may work incidentally, not designed for)
- FTP / cloud storage as a first-class location
- Bulk rename tooling
- File tagging or star/rating systems
- Plugin or extension system
