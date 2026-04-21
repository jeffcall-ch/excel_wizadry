# FileExplorer — Technical Audit Blueprint

**Scope**: Engineering-grade, code-sourced documentation for a future C++ rewrite.  
**Constraint**: All values are verbatim from the source code. Where a value is not explicitly defined in the provided source, this document states so rather than guessing.

---

## Table of Contents

1. [Visual / Aesthetic Audit](#1-visual--aesthetic-audit)  
   1.1 [Colour Palette](#11-colour-palette)  
   1.2 [Typography](#12-typography)  
   1.3 [Spacing & Layout Dimensions](#13-spacing--layout-dimensions)  
   1.4 [File-Extension Colour Map](#14-file-extension-colour-map)  
   1.5 [Window & Chrome Structure](#15-window--chrome-structure)  
2. [UI Component Architecture](#2-ui-component-architecture)  
   2.1 [Application Entry Point & DI](#21-application-entry-point--di)  
   2.2 [Main Window Layout](#22-main-window-layout)  
   2.3 [Tab System](#23-tab-system)  
   2.4 [Address Bar & Breadcrumb](#24-address-bar--breadcrumb)  
   2.5 [File List View](#25-file-list-view)  
   2.6 [Inspector Sidebar View](#26-inspector-sidebar-view)  
   2.7 [Tree View](#27-tree-view)  
   2.8 [Status Bar](#28-status-bar)  
   2.9 [Search Box](#29-search-box)  
3. [Integration Layer](#3-integration-layer)  
   3.1 [File System Service](#31-file-system-service)  
   3.2 [RQ Search Engine Bridge](#32-rq-search-engine-bridge)  
   3.3 [Identity Extraction Service](#33-identity-extraction-service)  
   3.4 [SQLite Identity Cache](#34-sqlite-identity-cache)  
   3.5 [Cloud Status Service (OneDrive)](#35-cloud-status-service-onedrive)  
   3.6 [Settings Service](#36-settings-service)  
   3.7 [Tab Path State Service](#37-tab-path-state-service)  
   3.8 [Directory Snapshot Cache](#38-directory-snapshot-cache)  
4. [Data Models](#4-data-models)  
5. [How It Works — Functional Manual](#5-how-it-works--functional-manual)  
   5.1 [Application Startup Sequence](#51-application-startup-sequence)  
   5.2 [Navigating to a Directory](#52-navigating-to-a-directory)  
   5.3 [Address Bar Edit Mode](#53-address-bar-edit-mode)  
   5.4 [Tab Operations](#54-tab-operations)  
   5.5 [Search Flow](#55-search-flow)  
   5.6 [File Operations (Copy / Cut / Paste / Rename / Delete)](#56-file-operations-copy--cut--paste--rename--delete)  
   5.7 [Identity Extraction & Cache Pipeline](#57-identity-extraction--cache-pipeline)  
   5.8 [OneDrive Cloud Status Resolution](#58-onedrive-cloud-status-resolution)  
   5.9 [FileSystemWatcher Incremental Refresh](#59-filesystemwatcher-incremental-refresh)  
   5.10 [Session Persistence](#510-session-persistence)

---

## 1. Visual / Aesthetic Audit

### 1.1 Colour Palette

The application defines **no custom colour resources** in `App.xaml`. Every background, foreground, border, and hover fill that is not listed in §1.4 below is resolved entirely through WinUI 3 `ThemeResource` tokens at runtime.

Colours explicitly hard-coded in XAML or code:

| Usage | Value | Source file |
|---|---|---|
| File-list item hover background | `#146FA8DC` (≈10 % accent blue) | `FileListView.xaml` |
| Breadcrumb segment button background | `Microsoft.UI.Colors.Transparent` | `MainWindow.xaml.cs` |
| Breadcrumb separator button background | `Microsoft.UI.Colors.Transparent` | `MainWindow.xaml.cs` |
| Extension colour — PDF (always-on) | `#E57373` (ARGB `FF, 229, 115, 115`) | `ExtensionToBrushConverter.cs` |
| Extension colour — Excel (dark mode) | `#81C784` (ARGB `FF, 129, 199, 132`) | `ExtensionToBrushConverter.cs` |
| Extension colour — Word (dark mode) | `#64B5F6` (ARGB `FF, 100, 181, 246`) | `ExtensionToBrushConverter.cs` |
| Extension colour — Archive (dark mode) | `#FFD54F` (ARGB `FF, 255, 213, 79`) | `ExtensionToBrushConverter.cs` |
| Extension colour — Text/Markdown (dark mode) | `#4DB6AC` (ARGB `FF, 77, 182, 172`) | `ExtensionToBrushConverter.cs` |

All other colours (`SubtleFillColorTransparentBrush`, `DividerStrokeColorDefaultBrush`, `LayerOnMicaBaseAltFillColorDefaultBrush`, `TextFillColorPrimaryBrush`, etc.) are WinUI ThemeResource tokens — their resolved RGB values are theme- and OS-version–dependent.

### 1.2 Typography

| Element | FontSize | FontWeight | CharacterSpacing | Opacity |
|---|---|---|---|---|
| Column header labels | 13.2 px | SemiBold | — | — |
| File row text | 13.2 px | Normal | — | — |
| Date-divider group header | 12.1 px | SemiBold | 40 | 0.75 |
| Process-banner status text | 12.1 px | Normal | — | — |
| Sort icon (Segoe Fluent Icons) | 10 px | — | — | — |
| File icon | 16 px | — | — | — |
| Empty-state icon | 64 px | — | — | 0.3 |
| Empty-state text | 16 px | — | — | 0.5 |
| Breadcrumb segment buttons | 12 px | Normal | — | — |
| Breadcrumb separator icon (E76C) | 10 px | — | — | 0.6 |
| Sidebar item text | 13.2 px | Normal | — | — |
| Sidebar favorites header | 13.2 px | SemiBold | — | — |
| Sidebar right-arrow icon (E76C) | Value not explicitly defined in provided source | — | — | 0.65 |
| Sidebar file icon | Value not explicitly defined in provided source | — | — | 0.75 |

All text uses the default WinUI fallback font family (`FontFamily` not overridden in XAML).

### 1.3 Spacing & Layout Dimensions

**Root window Grid rows:**

| Row | Height | Content |
|---|---|---|
| 0 | `Auto` (driven by TabView title bar height) | `TabView` (tabs + title bar drag region) |
| 1 | `40` px | Navigation bar (Back/Fwd/Up/Refresh + breadcrumb + search) |
| 2 | `0` px (removed / unused) | — |
| 3 | `*` (fills remaining height) | Tree column + file list + splitter + sidebar |
| 4 | `24` px | Status bar |

**Root window Grid columns (Row 3):**

| Column | Width | Content |
|---|---|---|
| 0 | `0` (collapsed) | TreeView navigation pane |
| 1 | `*` | File list |
| 2 | `8` px | GridSplitter |
| 3 | `320` px | Inspector sidebar |

**Navigation bar (Row 1) Grid columns:**

| Column | Width | Content |
|---|---|---|
| 0 | `Auto` | Back / Forward / Up / Refresh buttons |
| 1 | `*` | Address bar area (breadcrumb + AutoSuggestBox overlay) |
| 2 | `320` px | Search box |

**Address bar area:**

- `CornerRadius="4"`
- `BorderThickness="1"`
- Border colour: `CardStrokeColorDefaultBrush` (ThemeResource)

**File list item:**

- `Height="22"`, `MinHeight="22"`, `Padding="0"`, `Margin="0"`

**Column header:**

- Background: `LayerOnMicaBaseAltFillColorDefaultBrush`
- `BorderThickness="0,0,0,1"` (bottom divider only)

**Column gripper:**

- Width: `6` px
- Inner `Border` width: `1` px, colour: `DividerStrokeColorDefaultBrush`

**Loading indicator:**

- `ProgressRing` 40 × 40 px, centred in the file list

**Processing banner** (bottom-left overlay):

- `CornerRadius="4"`, `Padding="8,4"`
- `ProgressRing` 12 × 12 px + 12.1 px status text

**Sidebar divider:**

- 1 px horizontal `Border`, `Margin="0,6"`

**Sidebar item row:**

- `MinHeight="22"`, `Padding="8,0"`

### 1.4 File-Extension Colour Map

The `ExtensionToBrushConverter` applies foreground colour to the `Name` column text.  
**PDF colour is active in both light and dark themes.  
All other colours are suppressed in light theme** (`IsDarkThemeActive()` check via `ActualTheme == ElementTheme.Dark`).

| Extensions | Hex Colour | Category |
|---|---|---|
| `.pdf` | `#E57373` | PDF |
| `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.xlt`, `.xltx`, `.xltm` | `#81C784` | Excel |
| `.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.dotm` | `#64B5F6` | Word |
| `.zip`, `.7z` | `#FFD54F` | Archive |
| `.txt`, `.md` | `#4DB6AC` | Plain text |
| All others | `DependencyProperty.UnsetValue` (inherits theme default) | — |

Dark-mode detection precedence: `App.MainWindow.Content.ActualTheme` → `Application.Current.RequestedTheme`.

### 1.5 Window & Chrome Structure

- **Window type**: `WinUI 3 Window` (unpackaged).
- **Title bar**: Extended into the `TabView` title area; drag region covers the tab strip.
- **Window title format**: `{TabTitle} - FileExplorer`; falls back to `FileExplorer` if no active tab.
- **Crash reports**: written to `%LOCALAPPDATA%\FileExplorer\crash-reports\`.
- **Log files**: rolling daily Serilog files in `%LOCALAPPDATA%\FileExplorer\logs\log-.txt`, retained 14 days.
- **WER mini-dumps**: `%LOCALAPPDATA%\FileExplorer\crash-reports\wer-dumps\`, dump type 1 (mini), max 10.

---

## 2. UI Component Architecture

### 2.1 Application Entry Point & DI

**File**: `FileExplorer.App/App.xaml.cs`  
**Pattern**: `IServiceCollection` / `IServiceProvider` (Microsoft.Extensions.DependencyInjection).

Services registered:

| Lifetime | Type |
|---|---|
| Singleton | `MainViewModel` |
| Singleton | `TreeViewModel` |
| Singleton | `Func<TabState?, TabContentViewModel>` (tab factory delegate) |
| Transient | `TabContentViewModel` (indirect — created through factory only) |
| All infrastructure services | via `services.AddInfrastructureServices()` extension method |

Logging: Serilog with rolling file sink. `ILogger<T>` resolved through `Microsoft.Extensions.Logging.ILoggerFactory`.

**Tab factory** (verbatim constructor args order into `TabContentViewModel`):

1. `IFileSystemService`
2. `IShellIntegrationService`
3. `ICloudStatusService`
4. `IPreviewService`
5. `IIdentityExtractionService`
6. `IIdentityCacheService`
7. `ITabPathStateService`
8. `ISettingsService`
9. `IDirectorySnapshotCache`
10. `ILogger<TabContentViewModel>`
11. `TabState?` (existing session state, or `null` for a fresh tab)

### 2.2 Main Window Layout

**File**: `FileExplorer.App/MainWindow.xaml` + `MainWindow.xaml.cs`

The root element is a `Grid` with 5 rows (see §1.3 for dimensions). Key named elements:

| XAML Name | Type | Purpose |
|---|---|---|
| `TabViewRoot` | `TabView` | Tab strip + content host |
| `AddressBarArea` | `Border` | Container for breadcrumb/edit overlay |
| `BreadcrumbScroll` | `ScrollViewer` | Horizontal scroller containing breadcrumb |
| `BreadcrumbPanel` | `StackPanel` (Horizontal) | Dynamically-built path segments |
| `AddressBar` | `AutoSuggestBox` (Width=`*`) | Editable path field (overlay mode) |
| `SearchBox` | `AutoSuggestBox` (Width=320) | Search query input |
| `TreeSplit` | `TreeView` | Left navigation pane (collapsed, Width=0) |
| `FileList` | `FileListView` (UserControl) | Main file details pane |
| `PreviewSplitter` | `GridSplitter` | Drag handle between file list and sidebar |
| `InspectorSidebar` | `InspectorSidebarView` (UserControl) | Right panel (destinations, favourites) |
| `StatusBarPanel` | `Grid` | Bottom 24 px status strip |

**InfoBar overlays**: two `InfoBar` controls layered over the file list for "processing" and "error" notifications; both use `IsOpen` binding for visibility.

### 2.3 Tab System

**ViewModel**: `MainViewModel.cs` owns `ObservableCollection<TabContentViewModel> Tabs` and `TabContentViewModel? ActiveTab`.

**Tab lifecycle**:

1. `OpenTabAsync(path, state)` — calls factory, subscribes `tab.Navigated += OnTabNavigated`, appends to `Tabs`, sets `ActiveTab`, calls `tab.NavigateAsync(path)`.
2. `CloseTab(tab)` — refuses if `tab.IsPinned` or `Tabs.Count <= 1`; unsubscribes `Navigated`, calls `tab.Dispose()`, removes from `Tabs`; sets `ActiveTab` to neighbour.
3. `CloseOtherTabs`, `CloseTabsToRight` — analogous batch close (skip pinned).
4. `DuplicateTabAsync(sourceTab)` — factory creates new tab, inserted at `index + 1`, navigates to `sourceTab.CurrentPath`.
5. `TogglePinTab(tab)` — flips `tab.IsPinned` and `tab.TabState.IsPinned`.
6. `GoToTab(int index)` — Ctrl+1–8 = 0-based index; Ctrl+9 = last tab.
7. `NextTab()` / `PreviousTab()` — wraps around with modulo.

**Tab header update**: `RefreshTabHeader(tab)` sets `TabViewItem.Header = tab.TabTitle`, `IsClosable = !tab.IsPinned`, and `ToolTipService` to `tab.CurrentPath`.

**Default startup path** (hard-coded fallback in `MainViewModel`):  
`C:\Users\szil\OneDrive - Kanadevia Inova\Desktop\Projects`

Resolution order: command-line `--path` arg → `DefaultStartupPath` → `settings.NewTabDefaultPath` → `%USERPROFILE%`.

### 2.4 Address Bar & Breadcrumb

**Normal (breadcrumb) mode**:
- `BreadcrumbScroll` visible, `AddressBar` collapsed.
- `UpdateAddressBar(path)` rebuilds `BreadcrumbPanel.Children` on every navigation:
  - Walks path segments from root to leaf via `Path.GetDirectoryName` loop.
  - Creates one `Button` per segment (`Padding="6,2,6,2"`, `CornerRadius=4`, `FontSize=12`, transparent background, zero border).
  - Inserts a `CreateBreadcrumbSeparatorButton` chevron (glyph `\uE76C`, 10 px, Opacity 0.6) before each segment and appends a trailing chevron after the last segment.
  - Scrolls `BreadcrumbScroll` to its right edge (`ChangeView(ScrollableWidth, null, null)`).
- Clicking a **segment** → `BreadcrumbSegment_Click` → `NavigateToAddressAsync(fullPath)`.
- Clicking a **chevron** → `BreadcrumbSeparator_Click` → opens a `MenuFlyout` with sibling directories fetched via `GetSiblingDirectoriesAsync`.

**Edit mode**:
- Triggered by `AddressBarArea_Tapped` (unless a segment `Button` was the tap source — checked by walking `VisualTreeHelper`).
- `ShowAddressBarEditMode()`: hides `BreadcrumbScroll`, shows `AddressBar` with current path text, focuses it, then selects all text via `FindChildOfType<TextBox>`.
- Submit: `AddressBar_QuerySubmitted` → `NavigateToAddressAsync(text)` → collapses `AddressBar`, shows `BreadcrumbScroll`.
- Cancel: Escape key in `AddressBar_KeyDown` → collapses `AddressBar`, shows `BreadcrumbScroll`.
- Lost focus: `AddressBar_LostFocus` posts a low-priority closure to collapse (deferred to allow suggestion clicks to fire first).

**Autocomplete**: `AddressBar_TextChanged` fires on user input; when `text.Length > 2` and the parent directory exists, enumerates up to 10 subdirectories matching `partial + "*"`.

### 2.5 File List View

**File**: `FileExplorer.App/Views/FileListView.xaml` (UserControl)

**Column definitions** (fixed widths unless noted):

| # | Header | Width | Binding |
|---|---|---|---|
| 0 | Name | 532 px | `Name` + extension-colour converter |
| 1 | Project | 180 px | `Entry.ExtractedProject` |
| 2 | Rev | 80 px | `Entry.ExtractedRevision` |
| 3 | Date modified | 160 px | `Entry.DateModified` |
| 4 | Size | 90 px | `Entry.Size` |
| 5 | (trailing) | `*` | — |

**Date/group divider row**: rendered as a special item type with `FontSize=12.1`, `FontWeight=SemiBold`, `CharacterSpacing=40`, `Opacity=0.75`.

**Sort chevrons**: Unicode glyphs from Segoe Fluent Icons — ascending `\uE70E` (ChevronUp), descending `\uE70D` (ChevronDown), 10 px.

**File icon**: 16 × 16 px `Image` bound to `IconBitmap`; shell icon is resolved via `SHGetFileInfo` (index stored in `FileSystemEntry.IconIndex`).

**Hover**: `ListViewItemPresenter.PointerOverBackground` = `#146FA8DC`.

**Empty state**: centred `FontIcon` (64 px, Opacity 0.3) + `TextBlock` (16 px, Opacity 0.5).

**Loading state**: centred `ProgressRing` (40 × 40 px).

**Processing overlay** (bottom-left): `ProgressRing` (12 px) + `TextBlock` (12.1 px), `CornerRadius=4`, `Padding="8,4"`.

### 2.6 Inspector Sidebar View

**File**: `FileExplorer.App/Views/InspectorSidebarView.xaml` (UserControl)  
**Width**: 320 px (set in parent `Grid` column definition).

**Layout** (top to bottom):

1. **Navigation hub list** — `ListView` of quick-access destinations. Items: `MinHeight=22`, `Padding="8,0"`, `FontSize=13.2`. Right-side icon `\uE76C` at Opacity 0.65 indicates navigation action.
2. **Divider** — 1 px `Border`, `Margin="0,6"`.
3. **Favorites header** — `TextBlock`, `FontSize=13.2`, `FontWeight=SemiBold`, `Margin="8,0,8,4"`.
4. **Favorites list** — `ListView` of pinned paths. Items: `MinHeight=22`, `Padding="8,0"`, `FontSize=13.2`. File icon at Opacity 0.75.

### 2.7 Tree View

**Column width**: `0` (the tree pane is collapsed in the current default layout). The `GridSplitter` and drag-resize are present but the pane starts hidden.

Exact tree population logic and data templates: value not explicitly defined in provided source (would require `TreeViewModel.cs` which was not fully read for this audit).

### 2.8 Status Bar

**Height**: 24 px (Row 4 of root Grid).  
Content: `StatusText` and `FreeSpaceText` from `TabContentViewModel`, updated in `UpdateStatusText()` and `UpdateFreeSpace()` after each navigation or file-count change.

### 2.9 Search Box

**Control**: `AutoSuggestBox`, Width=320 px, in nav-bar Column 2.  
**Flow**: user types → `SearchBox_QuerySubmitted` → `MainViewModel.SearchAsync(query)`.  
On query completion or navigation, `_isSearchActive` flag is set/cleared and `FileList.SetSearchMode(bool)` toggled.

---

## 3. Integration Layer

### 3.1 File System Service

**Interface**: `IFileSystemService`  
**Implementation**: `FileExplorer.Infrastructure/Services/FileSystemService.cs`

Key capabilities (from callers observed in ViewModels):

- `EnumerateDirectoryAsync(path, ct)` — async enumerator over `FileSystemEntry` via Win32 `FindFirstFileEx`.
- `RenameAsync(oldPath, newName)` — fires `OperationError` event on failure with `FileOperationErrorEventArgs`.
- `DeleteAsync(paths)`, `CopyAsync(sources, destination)`, `MoveAsync(sources, destination)` — file operations with undo support.
- `ExecuteMoveAsync(...)` — uses `MoveFileExW` with flag `MOVEFILE_COPY_ALLOWED` only (`MOVEFILE_WRITE_THROUGH` was removed for performance).
- `CanUndo` property — indicates whether an undo record exists.
- `UndoLastOperationAsync()` — reverts last registered operation.
- `OperationError` event (`EventHandler<FileOperationErrorEventArgs>`).

### 3.2 RQ Search Engine Bridge

**Interface**: `ISearchService`  
**Implementation**: `FileExplorer.Infrastructure/Services/RqSearchService.cs`  
**Binary location**: `{AppContext.BaseDirectory}\Assets\rq.exe`

**Invocation** (no shell; safe via `ArgumentList`):

```
rq.exe <directory> <searchTerm> --glob --follow-symlinks
```

**Result collection**: `OutputDataReceived` event; each non-empty `e.Data` line is one absolute file path added to a `ConcurrentBag<string>`.

**Cancellation**: `process.Kill(entireProcessTree: true)` on `OperationCanceledException`; returns empty list.

**Error capture**: stderr lines accumulated in `StringBuilder`, logged at Warning level after process exit.

**Return type**: `IReadOnlyList<string>` (all result paths).

**Caller** (`MainViewModel.SearchAsync`):
- Cancels previous `CancellationTokenSource` before starting.
- Clears `ActiveTab.Items`.
- Creates `FileSystemEntry` objects from each result path (via `DirectoryInfo` or `FileInfo`).
- Populates `ActiveTab.Items` directly (no sort applied on search results).

### 3.3 Identity Extraction Service

**Interface**: `IIdentityExtractionService`  
**Implementation**: `FileExplorer.Infrastructure/Services/IdentityExtractionService.cs`  
**Dependencies**: `UglyToad.PdfPig` (NuGet), `System.IO.Compression` (ZipArchive), `System.Xml.Linq`.

**Supported extensions** (exact set):

```
.doc  .docx  .docm  .dot  .dotx  .dotm
.xls  .xlsx  .xlsm  .xlt  .xltx  .xltm  .xlsb
.pdf
```

**In-process memory cache**: `ConcurrentDictionary<string, CachedIdentity>` keyed by file path.  
Invalidation key: `LastWriteTimeUtc.Ticks` + `file.Length`. Hit = cache returned directly.

**Dispatch table** (`ExtractIdentityAsync`):

| Extension | Method |
|---|---|
| `.docx`, `.docm`, `.dotx`, `.dotm` | `ExtractFromWordOpenXmlAsync` |
| `.xlsx`, `.xlsm`, `.xltx`, `.xltm` | `ExtractFromExcelOpenXmlAsync` |
| `.pdf` | `ExtractFromPdfFirstPageAsync` |
| `.doc`, `.dot`, `.xls`, `.xlt`, `.xlsb` | `ExtractFromLegacyBinaryScanAsync` |

**Word OpenXML** (`ExtractFromWordOpenXmlAsync`):

1. Open ZIP, read `word/document.xml`.
2. Extract first-page body tokens + header/footer XML tokens.
3. Join all tokens, run `ParseIdentityFromText`.
4. `ExtractProjectFromWordTokens` (anchor-first), then fallback `ExtractProjectFromWordHeaderFooterRawXmlAsync`.
5. Final: `ProjectTitle = FirstNonEmpty(anchorProject, parsedProject)`.

**Excel OpenXML** (`ExtractFromExcelOpenXmlAsync`):

1. Parse `xl/workbook.xml` and relationship file.
2. Score and order sheets by name: cover-sheet names first (`IsCoverSheetName`).
3. Per sheet: resolve worksheet path, parse cells, resolve against shared-strings table.
4. `FindProjectNameCellBelowAnchor(cells)` — look for a cell with "project" label and read the cell below it.
5. Stop at the first sheet that yields a non-empty project value.

**PDF** (`ExtractFromPdfFirstPageAsync`):

1. PdfPig `PdfDocument.Open`, read page 1.
2. `GetOrderedPdfWords`: words sorted descending by Top-Y then Left-X (reading order top→bottom, left→right).
3. `BuildPdfSortedFlowText`: group words by Y-coordinate within a 5.0-unit tolerance; output as `\n`-joined lines.
4. Anchor-based extraction (`ExtractPdfProjectFromAnchors`, `ExtractPdfRevisionFromAnchors`, `ExtractPdfDocumentNameFromAnchors`).
5. Regex fallbacks on the sorted flow text.
6. Raw text fallback if any field is still empty.

**Legacy binary** (`ExtractFromLegacyBinaryScanAsync`):

1. Read at most **4 MB** from the file using `FileMode.Open, FileAccess.Read, FileShare.ReadWrite|Delete`.
2. Decode as Latin-1 (byte-preserving).
3. Apply `ParseIdentityFromText` to the raw string.

**Post-extraction fixups** (applied regardless of format):

- **Revision fallback from filename**: regex `(?<!\d)(?<rev>\d+\.\d+)(?!\d)` on the filename stem. For PDFs: only applied when `identity.Revision` is empty.
- **DocumentName fallback from filename**: if `identity.DocumentName` is empty, derive from filename stem.

**Regex anchors** (compiled, `IgnoreCase`):

| Field | Pattern |
|---|---|
| Project | `\bproject(?!\s*number)\s*(?:name)?\s*[:\-]\s*(?<v>[^\r\n\t\|;]+)` |
| Revision | `\b(revision\|rev\.?)\s*[:\-]\s*(?<v>[^\r\n\t\|;]+)` |
| PDF internal revision | `(?:revision\|rev)\s*[:\-]?\s*(?<v>[A-Z0-9.\-]+)` |
| Document name / title | `\b(doc(?:ument)?\s*name\|title)\s*[:\-]\s*(?<v>[^\r\n\t\|;]+)` |
| Revision from filename | `(?<!\d)(?<rev>\d+\.\d+)(?!\d)` |

**Batch extraction** (`ExtractIdentityBatchAsync`):

- Filters unsupported extensions and deduplicates paths before processing.
- `Parallel.ForEachAsync` with `MaxDegreeOfParallelism = Math.Max(2, Environment.ProcessorCount)`.
- Only entries where `identity.HasAnyValue == true` are included in the returned dictionary.

### 3.4 SQLite Identity Cache

**Interface**: `IIdentityCacheService`  
**Implementation**: `FileExplorer.Infrastructure/Services/SqliteIdentityCacheService.cs`  
**Database file**: `identity_cache.db`

**Path resolution order**:

1. `FILEEXPLORER_IDENTITY_CACHE_DB` environment variable (full path, `%VAR%` expanded).
2. Walk up from `AppContext.BaseDirectory` up to 12 levels looking for `FileExplorer.sln`; use `{solutionRoot}\Development\identity_cache.db`.
3. Fallback: `%LOCALAPPDATA%\FileExplorer\Development\identity_cache.db`.

**Initialisation** (lazy, double-checked lock via `SemaphoreSlim`):

```sql
PRAGMA journal_mode = WAL;
PRAGMA synchronous = NORMAL;
```

**Schema DDL**:

```sql
CREATE TABLE IF NOT EXISTS FileIdentityCache (
    FileKey      TEXT    PRIMARY KEY,
    ProjectTitle TEXT,
    Revision     TEXT,
    DocumentName TEXT,
    LastWriteTime INTEGER NOT NULL,
    FileSize      INTEGER NOT NULL,
    ParentPath    TEXT    NOT NULL
);

CREATE INDEX IF NOT EXISTS IX_FileIdentityCache_ParentPath   ON FileIdentityCache (ParentPath);
CREATE INDEX IF NOT EXISTS idx_parent_path                   ON FileIdentityCache (ParentPath);
CREATE INDEX IF NOT EXISTS idx_ParentPath                    ON FileIdentityCache (ParentPath);
CREATE INDEX IF NOT EXISTS IX_FileIdentityCache_LastWriteTime ON FileIdentityCache (LastWriteTime);
```

> Note: three redundant indexes exist on `ParentPath` — a known legacy artifact.

**Key construction** (`BuildFileKey`):

```
FileKey = "{NormalizePath(fullPath)}|{fileSize}|{lastWriteTime.ToUniversalTime().Ticks}"
```

**Path normalisation** (`NormalizePath`):

1. Trim whitespace and surrounding `"` quotes.
2. Strip `\\?\UNC\` → `\\`, or `\\?\` → (remove prefix).
3. `Path.GetFullPath(candidate)` (best-effort; silently ignored on failure).
4. Strip the long-path prefix again (in case `GetFullPath` re-added it).
5. Replace `AltDirectorySeparatorChar` with `DirectorySeparatorChar`.
6. TrimEnd path separator *only* when path is not the root itself.
7. Convert to lowercase.

**Queries**:

| Operation | SQL summary |
|---|---|
| Bulk read by parent | `WHERE ParentPath = $parentPath` → `Dictionary<string, FileIdentityCacheRecord>` |
| Single read by key | `WHERE FileKey = $fileKey` → `FileIdentityCacheRecord?` |
| Upsert | `INSERT … ON CONFLICT(FileKey) DO UPDATE SET …` in a transaction |
| Delete by keys | Per-key `DELETE WHERE FileKey = $fileKey` in a transaction |
| Delete by path | `DELETE WHERE FileKey LIKE $prefix` where prefix = `{normalizedPath}|%` |

**Connection settings**: `SqliteOpenMode.ReadWriteCreate`, `SqliteCacheMode.Shared`.

### 3.5 Cloud Status Service (OneDrive)

**Interface**: `ICloudStatusService`  
**Implementation**: `FileExplorer.Infrastructure/Services/CloudStatusService.cs`  
**Technique**: Windows Cloud Filter API via P/Invoke (`NativeMethods`). AOT-safe.

**Sync root detection** (`DetectSyncRootsAsync`):

Checks, in order:

1. `%OneDrive%`, `%OneDriveConsumer%`, `%OneDriveCommercial%` environment variables.
2. Hard-coded paths: `%USERPROFILE%\OneDrive`, `%USERPROFILE%\OneDrive - Personal`.
3. `Directory.EnumerateDirectories(%USERPROFILE%, "OneDrive*")` for corporate variants.

Each candidate is tested (exact method `TryAddSyncRoot`: value not explicitly defined in provided source; assumed to probe the Cloud Filter API).

**Status values** (`CloudFileStatus` enum):

| Value | Meaning |
|---|---|
| `NotApplicable` | Not under any OneDrive sync root |
| `CloudOnly` | `FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS` or `FILE_ATTRIBUTE_OFFLINE` |
| `AlwaysAvailable` | `FILE_ATTRIBUTE_PINNED` |
| `Syncing` | `CF_PLACEHOLDER_STATE_PARTIAL` |
| `LocallyAvailable` | `CF_PLACEHOLDER_STATE_IN_SYNC` or `PARTIALLY_IN_SYNC` |

**Batch resolution**: `GetCloudStatusBatchAsync` — runs individual `GetCloudStatusAsync` calls in parallel via `Task.WhenAll`, guarded by a `SemaphoreSlim(1,1)`.

A faster overload accepts `IReadOnlyList<FileSystemEntry>` and uses the already-read `entry.Attributes` + `entry.ReparseTag` without additional I/O.

### 3.6 Settings Service

**Interface**: `ISettingsService`  
Settings persisted (from observed property usage in `MainViewModel`):

| Property | Type | Default |
|---|---|---|
| `ShowHiddenFiles` | `bool` | `false` |
| `ShowFileExtensions` | `bool` | `true` |
| `ShowPreviewPane` | `bool` | — (not set to explicit default in provided source) |
| `NavigationPaneWidth` | `double` | `250` |
| `PreviewPaneWidth` | `double` | `300` |
| `NewTabDefaultPath` | `string` | — |

`SaveTabSessionAsync(IList<TabState>)` and `LoadAsync()` / `SaveAsync()` are the persistence methods. Storage format: value not explicitly defined in provided source.

### 3.7 Tab Path State Service

**Interface**: `ITabPathStateService`  
**Purpose**: persists the current directory path and display title for each tab to a backing store (SQLite or JSON — format not explicitly defined in provided source).

API observed:

- `UpsertCurrentPathAsync(tabId, path, displayTitle)` — called on every `CurrentPath` change, debounced by the fact that `PersistAndApplyTabPathStateAsync` is fire-and-forget from property setter.
- `GetByTabIdAsync(tabId)` — retrieves persisted `DisplayTitle`; used to apply a meaningful folder-based title after the path state is known.

### 3.8 Directory Snapshot Cache

**Interface**: `IDirectorySnapshotCache`  
**Purpose**: pre-scanned directory snapshots populated at startup (or on speculation) for instant first-paint.

API observed:

- `TryConsume(path, out IReadOnlyList<FileSystemEntry> entries)` — atomic consume; if a snapshot exists for `path`, removes it from the cache and returns `true`.

When a cache hit occurs, `Items.Clear()` is called synchronously on the calling thread before returning (to prevent a race with `ReconcileItemsWithSorted`).

---

## 4. Data Models

### `FileSystemEntry` (`FileExplorer.Core/Models/FileSystemEntry.cs`)

Mutable class; `required` properties must be set at construction.

| Property | Type | Notes |
|---|---|---|
| `FullPath` | `string` | Required |
| `Name` | `string` | Required; display name (may differ from filename in edge cases) |
| `IsDirectory` | `bool` | Required |
| `Size` | `long` | Bytes; 0 for directories |
| `DateModified` | `DateTimeOffset` | Last-write time from file system |
| `TypeDescription` | `string` | Shell type string or computed fast description |
| `Extension` | `string` | Leading dot included (e.g., `.xlsx`) |
| `Attributes` | `FileAttributes` | Raw Win32 file attributes |
| `ReparseTag` | `uint` | NTFS reparse tag; 0 if not a reparse point |
| `IconIndex` | `int` | Shell icon index from `SHGetFileInfo` |
| `CloudStatus` | `CloudFileStatus` | Default `NotApplicable` |
| `ExtractedProject` | `string` | Populated from identity pipeline |
| `ExtractedRevision` | `string` | Populated from identity pipeline |
| `ExtractedDocumentName` | `string` | Populated from identity pipeline |
| `IsHidden` | `bool` (computed) | `Attributes & FileAttributes.Hidden` |
| `IsSystem` | `bool` (computed) | `Attributes & FileAttributes.System` |
| `ParentPath` | `string?` (computed) | `Path.GetDirectoryName(FullPath)` |
| `FileKey` | `string` (computed) | `"{normalizedPath}\|{isDir}\|{size}\|{utcTicks}"` |

Equality and `GetHashCode` compare: `NormalizePathForKey(FullPath)` (OrdinalIgnoreCase) + `IsDirectory` + `Size` + `DateModified.UtcTicks`.

### `DocumentIdentity` (`FileExplorer.Core/Models/DocumentIdentity.cs`)

Immutable `record`.

| Property | Type |
|---|---|
| `ProjectTitle` | `string` |
| `Revision` | `string` |
| `DocumentName` | `string` |
| `HasAnyValue` | `bool` (computed) |

`DocumentIdentity.Empty` is a shared static instance with all fields empty.

### `FileIdentityCacheRecord` (`FileExplorer.Core/Models/FileIdentityCacheRecord.cs`)

Immutable `record`. Maps 1-to-1 to the `FileIdentityCache` SQLite row.

| Property | Type | Notes |
|---|---|---|
| `FileKey` | `string` | Required; composite key |
| `ParentPath` | `string` | Required; normalised lowercase |
| `ProjectTitle` | `string` | Default empty |
| `Revision` | `string` | Default empty |
| `DocumentName` | `string` | Default empty |
| `LastWriteTime` | `long` | UTC ticks |
| `FileSize` | `long` | Bytes |

Method `ToDocumentIdentity()` projects to `DocumentIdentity`.

### `TabState`

Owned by `TabContentViewModel`. Stores `Id` (GUID), `CurrentPath`, `Title`, `IsPinned`, and a `NavigationHistory` instance. Serialised during session save via `ISettingsService.SaveTabSessionAsync`.

---

## 5. How It Works — Functional Manual

### 5.1 Application Startup Sequence

1. `App()` constructor: builds the DI container, configures Serilog, registers WER dump settings in `HKCU\Software\Microsoft\Windows\Windows Error Reporting\LocalDumps`.
2. `OnLaunched` creates `MainWindow`.
3. `MainWindow` constructor resolves `MainViewModel` from DI.
4. `MainViewModel.InitializeAsync()`:
   - Loads settings (`ISettingsService.LoadAsync`).
   - Applies `ShowHiddenFiles`, `ShowFileExtensions`, `ShowPreviewPane`, pane widths.
   - `ICloudStatusService.DetectSyncRootsAsync()` — scans for OneDrive roots.
   - Resolves initial path (see §2.3 default path logic).
   - Calls `OpenTabAsync(initialPath, null)`.
5. `OpenTabAsync` creates a `TabContentViewModel` via the factory delegate and calls `NavigateAsync(path)`.

### 5.2 Navigating to a Directory

All navigation funnels through `TabContentViewModel.NavigateCoreAsync(path, skipHistory)`:

1. **Path normalisation**: `NormalizeNavigationPath(path)` (static, handles `shell:` prefixes via `IShellIntegrationService.ResolveShellPath`).
2. **Guard**: directory must exist (`Directory.Exists`), otherwise log warning and return.
3. `_loadCts.Cancel()` — cancels any in-flight navigation.
4. `IsLoading = true`.
5. `CurrentPath = path` — triggers `OnCurrentPathChanged`, which:
   - Updates `TabState.CurrentPath`.
   - Derives `TabTitle` from `BuildFriendlyTabTitle`.
   - Fires `PersistAndApplyTabPathStateAsync` (fire-and-forget) to write to `ITabPathStateService`.
6. `History.Navigate(...)` — pushes entry (unless `skipHistory: true`).
7. `_watcherHandle.Dispose()` — stops previous `FileSystemWatcher`.
8. **Two-stage load** (`StreamLoadDirectoryAsync`):
   - **Cache hit**: `IDirectorySnapshotCache.TryConsume` → `Items.Clear()` synchronously, return snapshot.
   - **Cache miss**: Win32 stream via `IFileSystemService.EnumerateDirectoryAsync`.
     - First 50 entries → posted to UI thread via `_uiContext.Post` (clearing `Items` and inserting the first batch for immediate display).
     - Scan continues until exhausted.
9. `ApplyFolderSortPreference(path)` — loads saved sort column/direction for this path.
10. `ApplySort(items)` — sorts the full list.
11. `ReconcileItemsWithSorted(sorted)` — reconciles `Items` (which may contain the streaming first-batch) into the final sorted list by reusing existing `FileItemViewModel` instances.
12. `IsLoading = false`, `IsEmpty` computed, `UpdateStatusText()`, `UpdateFreeSpace()`.
13. `StartWatcher(path)` — attaches `FileSystemWatcher`.
14. If OneDrive detected: `ResolveCloudStatusesAsync` (fire-and-forget).
15. `Navigated?.Invoke(this, path)` — consumed by `MainViewModel.OnTabNavigated` → `NotifyNavigationState()` → command `CanExecute` refresh.

### 5.3 Address Bar Edit Mode

See §2.4 for the full address-bar state machine. Summary:

- Breadcrumb (read-only) → tap any blank area → edit mode (`AutoSuggestBox` visible, text selected).
- Type partial path → autocomplete from `Directory.EnumerateDirectories`.
- Enter / suggestion chosen → navigate, return to breadcrumb mode.
- Escape / lost focus → return to breadcrumb mode without navigating.

### 5.4 Tab Operations

| Action | Keyboard / UX | Code path |
|---|---|---|
| New tab | Ctrl+T | `MainViewModel.NewTabCommand` → `NewTabAsync` |
| Close tab | Ctrl+W / × button | `MainViewModel.CloseTabCommand` |
| Duplicate tab | Context menu | `MainViewModel.DuplicateTabCommand` |
| Pin/unpin tab | Context menu | `MainViewModel.TogglePinTabCommand` |
| Next/previous tab | Ctrl+Tab / Ctrl+Shift+Tab | `MainViewModel.NextTabCommand / PreviousTabCommand` |
| Jump to tab n | Ctrl+1–9 | `MainViewModel.GoToTabCommand(n)` |
| Activate existing tab for path | (second-instance pipe) | `OpenFolderOrActivateExistingTabAsync` — exact-path match, then open new |

### 5.5 Search Flow

1. User types in `SearchBox` and presses Enter → `SearchBox_QuerySubmitted` → `MainViewModel.SearchAsync(query)`.
2. Previous `_searchCts` is cancelled; new `CancellationTokenSource` created.
3. `ActiveTab.IsLoading = true`, `ActiveTab.Items.Clear()`.
4. `ISearchService.SearchAsync(searchDir, query, ct)` → calls `rq.exe` (see §3.2).
5. For each returned path: construct `FileSystemEntry` from `DirectoryInfo`/`FileInfo`.
6. `ActiveTab.Items.Add(new FileItemViewModel(entry))` — no further sorting.
7. `ActiveTab.IsLoading = false`, `RefreshStatusText()`.
8. `FileList.SetSearchMode(true)` → alters column visibility or header state (exact behaviour: value not explicitly defined in provided source).
9. **Clear search**: `ClearSearchAsync()` cancels search CTS, navigates to `ActiveTab.CurrentPath` (full directory reload), `FileList.SetSearchMode(false)`.

### 5.6 File Operations (Copy / Cut / Paste / Rename / Delete)

#### Cut / Copy to clipboard

`CutCopyToClipboardAsync(isCut)` (async, in `MainWindow.xaml.cs`):

1. Collects `SelectedItems` paths from the active tab.
2. Stores paths in `_clipboardSourcePaths` and `_clipboardIsMove = isCut`.
3. Creates a `DataPackage` with a `StorageItemsAsync` getter using `StorageFile.GetFileFromPathAsync` / `StorageFolder.GetFolderFromPathAsync`.
4. Sets `DataPackage.RequestedOperation`.
5. `Clipboard.SetContent(dataPackage)`.
6. `ViewModel.HasClipboardContent = true`.

#### Paste

`PasteFromClipboardAsync()` (in `MainWindow.xaml.cs`):

1. **Source path cache first**: if `_clipboardSourcePaths` is set (our own cut/copy), skip `GetStorageItemsAsync` — use the cached paths directly.
2. Otherwise fall back to `Clipboard.GetContent().GetStorageItemsAsync()`.
3. Calls `IFileSystemService.MoveAsync` or `CopyAsync` depending on `_clipboardIsMove`.
4. **Optimistic destination insert**: calls `ActiveTab.InsertItemsSorted(movedEntries)` immediately.
5. **Optimistic source removal**: if `_clipboardIsMove` and the source tab is in `ViewModel.Tabs`, removes source items from that tab's `Items` collection directly.

#### Rename (inline)

1. F2 → `StartRename()` sets `FileItemViewModel.IsRenaming = true`, `EditingName = Name`.
2. Inline `TextBox` appears in the Name cell (bound to `EditingName`).
3. Enter / focus-loss → `CommitRenameAsync(item)`:
   - Optimistic: `item.Entry.Name = newName`, `item.NotifyNameChanged()`.
   - `IFileSystemService.RenameAsync(oldPath, newName)`.
   - On failure: revert `item.Entry.Name = oldName`, fire `RenameError` event with error message.
4. After success or failure: `MoveItemToSortedPosition(item)`, `RevealItemRequested` event.

#### Delete

Flow not fully traced from provided source; calls `IFileSystemService.DeleteAsync(paths)`.

### 5.7 Identity Extraction & Cache Pipeline

Triggered from `TabContentViewModel` after a directory load completes (exact trigger point: value not explicitly defined in provided source — `IIdentityExtractionService` and `IIdentityCacheService` are constructor dependencies).

**Per-navigation pipeline** (inferred from constructor injection and interface usage):

1. Build `FileKey` for each `FileSystemEntry` in `Items`.
2. Bulk-read `IIdentityCacheService.GetEntriesByParentPathAsync(currentPath)` — one SQLite query for the entire folder.
3. For each item: if `FileKey` matches a cached row **and** `LastWriteTime`/`FileSize` match, apply `ProjectTitle`, `Revision`, `DocumentName` directly from the cache. No file I/O.
4. Items with cache misses (new or modified files) are queued for `IIdentityExtractionService.ExtractIdentityBatchAsync`.
5. Extraction runs in parallel (CPU-bound via `Parallel.ForEachAsync`).
6. Results upserted to SQLite via `IIdentityCacheService.UpsertEntriesAsync`.
7. `FileSystemEntry.ExtractedProject`, `ExtractedRevision`, `ExtractedDocumentName` populated, triggering UI update.

**Processing banner** is shown while extraction is in progress.

### 5.8 OneDrive Cloud Status Resolution

Triggered from `NavigateCoreAsync` when `ICloudStatusService.IsSyncRootDetected == true`:

`ResolveCloudStatusesAsync(ct)`:

1. Calls `ICloudStatusService.GetCloudStatusBatchAsync(entries)` — uses already-read `entry.Attributes` + `entry.ReparseTag`; no disk I/O.
2. Sets `entry.CloudStatus` on each item.
3. UI column binding: `ShowOneDriveColumn` property on `MainViewModel` gates column visibility.

### 5.9 FileSystemWatcher Incremental Refresh

`StartWatcher(path)` in `TabContentViewModel` creates a `_watcherHandle` (via `IFileSystemService`). On change events:

- Small changes: `RefreshIncrementalAsync(path)` — re-enumerates and reconciles `Items` without a full navigation, preserving scroll position.
- `IsAutoRefreshPaused` flag: set before clipboard operations, cleared afterwards, to suppress spurious watcher events during in-progress file moves.

`_lastLoadCompletedUtc` is updated after every full load or incremental refresh; used to gate whether a watcher event is stale.

### 5.10 Session Persistence

**On window close** / explicit save:

1. `MainViewModel.SaveSessionAsync()`:
   - Writes `NavigationPaneWidth` and `PreviewPaneWidth` back to settings.
   - `ISettingsService.SaveTabSessionAsync(Tabs.Select(t => t.TabState).ToList())`.
   - `ISettingsService.SaveAsync()`.

**On restore** (`InitializeAsync`):

- `ISettingsService.LoadAsync()` reads saved settings including tab paths.
- Currently opens only the **first/initial tab** from `ResolveInitialStartupPath`; full multi-tab session restore: value not explicitly defined in provided source (would require reading `SettingsService` implementation).

---

*End of Technical Audit Blueprint*
