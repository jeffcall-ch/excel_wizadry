# UI and Functional Updates Plan

Date: 2026-04-05
Project: FileExplorer (WinUI 3, .NET 8)

## 0. Scope and constraints

This plan implements only the requested UI and function changes:

1. High-density two-pane workspace where the details list is primary.
2. Rebuilt interactive breadcrumb with sibling flyouts.
3. Inspector sidebar with identity and metadata sections.
4. Physical dump area mapped to C:\temp\x_file_dump.
5. Deep fly-out navigation hub behavior on hover and click.
6. High-speed extraction logic for docx/xlsx with parallel processing.

Out of scope in this version:

1. Quick Copy buttons in identity card.
2. Text content preview and image thumbnail preview.
3. Unrelated command bar, tab system, and undo redesign.

Decisions confirmed:

1. C# extraction is primary for identity fields; rq is used for recursive folder stats.
2. Supported identity extraction file families include Word and Excel types (doc, docx, xls, xlsx, xlsm, and related formats where feasible).
3. For unrelated file types, extracted fields remain empty.
4. Dump area drop behavior defaults to Move.
5. Breadcrumb separator sibling menu includes folders only.
6. Inspector sidebar is resizable by splitter.
7. The top flyout launcher button will be removed once the sidebar navigation hub is fully active.
8. Performance target prioritizes full identity population for all files in the loaded set.

## 1. Current baseline and impact map

Current implementation already provides:

1. Main shell with tree pane, file list, and preview pane in MainWindow.xaml and MainWindow.xaml.cs.
2. Details list with columns Name, OneDrive, Date modified, Type, Size in Views/FileListView.xaml.
3. Existing breadcrumb segment click navigation in MainWindow.xaml.cs (UpdateAddressBar and BreadcrumbSegment_Click).
4. Existing asynchronous deep fly-out infrastructure in MainWindow.xaml.cs for shortcut launcher cache.
5. Existing rq CLI integration for search in Infrastructure/Services/RqSearchService.cs.

Impact summary:

1. MainWindow layout will be refactored to remove tree and preview panes, and add a right inspector sidebar.
2. File list schema will be extended to include extracted fields and extension color styling.
3. Breadcrumb separators will become interactive sibling dropdown triggers.
4. New services and models are required for identity extraction and recursive folder stats.
5. Existing shortcut fly-out logic will be reused and generalized as the new navigation hub.

## 2. Target architecture

## 2.1 Visual structure

1. Top row: back, forward, up, refresh, interactive breadcrumb, search.
2. Main content row: left details list (primary, wide), right inspector sidebar (narrow, resizable by splitter).
3. No tree view pane.
4. No preview pane.

## 2.2 Main file list columns

Required columns in order:

1. Name
2. Extracted Project
3. Extracted Revision
4. Date Modified
5. Size

Column behavior:

1. Keep sorting on all visible columns.
2. Keep virtualization and keyboard navigation.
3. Keep drag and drop.

## 2.3 Data flow

1. Directory enumeration remains in TabContentViewModel via IFileSystemService.
2. Identity extraction and recursive stats run asynchronously in background tasks.
3. File rows initially render fast, then extracted fields hydrate progressively.
4. Inspector listens to selected row changes and requests metadata/identity payload.

## 3. Implementation plan by workstream

## 3.1 Workstream A: Shell layout refactor (remove tree, add inspector)

Tasks:

1. Update FileExplorer.App/MainWindow.xaml main content Grid to two columns: FileList and InspectorSidebar.
2. Remove TreeView, tree splitter, and preview pane from visible layout.
3. Keep existing top navigation controls and status bar.
4. Add new user control FileExplorer.App/Views/InspectorSidebarView.xaml with sections:
   - Identity Card
   - Metadata
   - Physical Dump Area
   - Navigation Hub
5. Wire selection context from active tab into InspectorSidebarView.

Likely files:

1. FileExplorer.App/MainWindow.xaml
2. FileExplorer.App/MainWindow.xaml.cs
3. FileExplorer.App/Views/InspectorSidebarView.xaml (new)
4. FileExplorer.App/Views/InspectorSidebarView.xaml.cs (new)

Acceptance criteria:

1. Left details list occupies majority width.
2. Right narrow inspector is always visible.
3. Tree pane and preview pane are not rendered.

## 3.2 Workstream B: Main file list columns and extension coloring

Tasks:

1. Replace current OneDrive and Type columns with Extracted Project and Extracted Revision.
2. Add extension-based text coloring for dark mode:
   - .xls, .xlsx, .xlsm -> #81C784
   - .doc, .docx -> #64B5F6
   - .zip, .7z -> #FFD54F
   - .txt, .md -> #4DB6AC
3. Implement converter for extension-to-brush and apply only in dark theme.
4. Preserve current row density tuning and selection visuals.
5. For non-Word/Excel families, extracted columns stay empty.

Likely files:

1. FileExplorer.App/Views/FileListView.xaml
2. FileExplorer.App/Views/FileListView.xaml.cs
3. FileExplorer.App/Converters/ExtensionToBrushConverter.cs (new)
4. FileExplorer.Core/ViewModels/TabContentViewModel.cs
5. FileExplorer.Core/Models/FileSystemEntry.cs

Acceptance criteria:

1. Required columns appear and bind correctly.
2. Extension coloring applies only to targeted extensions in dark mode.
3. No regressions in sort, rename, keyboard selection, or drag and drop.

## 3.3 Workstream C: Interactive breadcrumb with sibling fly-outs

Tasks:

1. Keep segment label click navigation to that segment path.
2. Change separator chevron to clickable element per segment boundary.
3. On chevron click, load sibling directories of the segment parent asynchronously.
4. Build flyout menu for sibling folders only; selecting one navigates immediately.
5. Add caching and cancellation to avoid UI stutter under rapid clicks.

Likely files:

1. FileExplorer.App/MainWindow.xaml
2. FileExplorer.App/MainWindow.xaml.cs
3. Optional new helper: FileExplorer.Core/Models/BreadcrumbSegment.cs (new)
4. Optional new service: FileExplorer.Infrastructure/Services/BreadcrumbFolderService.cs (new)

Acceptance criteria:

1. Breadcrumb displays Folder > Subfolder > Current > visual pattern.
2. Segment click navigates to that exact path.
3. Separator click opens sibling folder menu without UI freeze.

## 3.4 Workstream D: Inspector identity card and metadata

Tasks:

1. Add identity fields in inspector:
   - Project Title
   - Revision
   - Document Name
2. No quick copy buttons.
3. Metadata block:
   - Timestamps in DD.MM.YYYY HH:mm format.
   - For folders: recursive file count and total size.
4. Remove all content preview and thumbnails from inspector path.
5. Build fallback states for missing or inaccessible metadata.

Likely files:

1. FileExplorer.App/Views/InspectorSidebarView.xaml (new)
2. FileExplorer.App/Views/InspectorSidebarView.xaml.cs (new)
3. FileExplorer.Core/Models/FileSystemEntry.cs
4. FileExplorer.Core/ViewModels/TabContentViewModel.cs
5. FileExplorer.Core/ViewModels/MainViewModel.cs

Acceptance criteria:

1. Identity card updates on row selection with anchor-derived values.
2. Folder recursive stats appear asynchronously and do not block navigation.
3. No text/image content peek is shown.

## 3.5 Workstream E: rq utility integration and extraction pipeline

Tasks:

1. Extend rq integration beyond search:
   - recursive folder size and file count
2. Add new interface for utility operations to avoid overloading ISearchService semantics.
3. Add robust timeout and cancellation handling for external rq calls.
4. Log timing metrics for extraction and stats calls.

Likely files:

1. FileExplorer.Core/Interfaces/IRqUtilityService.cs (new)
2. FileExplorer.Infrastructure/Services/RqUtilityService.cs (new)
3. FileExplorer.Infrastructure/ServiceCollectionExtensions.cs
4. FileExplorer.Infrastructure/Services/RqSearchService.cs (small refactor for shared process invocation helper)

Acceptance criteria:

1. Folder recursive stats are produced through rq utility integration path.
2. Calls are non-blocking and cancellable.
3. Error handling degrades gracefully to N/A in UI.

## 3.6 Workstream F: C# sniper extraction for Word and Excel families

Tasks:

1. Office Open XML extractor path:
   - .docx: stream ZipArchive entry word/document.xml and apply anchor regex
   - .xlsx and .xlsm: parse workbook XML and worksheet XML for anchor-adjacent values
2. Legacy Office fallback path:
   - .doc and .xls families attempt rq-assisted extraction if available
   - if extraction is not reliable, return empty identity fields instead of blocking UI
3. Parallel processing strategy:
   - Task.Run for file batches
   - Parallel.ForEach for per-file extraction
   - bounded concurrency to avoid IO thrash
4. Cache extraction results by file path + LastWriteTime for instant revisits.

Likely files:

1. FileExplorer.Core/Interfaces/IIdentityExtractionService.cs (new)
2. FileExplorer.Infrastructure/Services/IdentityExtractionService.cs (new)
3. FileExplorer.Core/Models/DocumentIdentity.cs (new)
4. FileExplorer.Core/ViewModels/TabContentViewModel.cs

Acceptance criteria:

1. Anchor extraction works for Word and Excel file families with graceful fallback.
2. Folder with 150+ files remains responsive and metadata hydration starts immediately.
3. End-to-end extraction target is under 300 ms when cache is warm and IO conditions are typical.

## 3.7 Workstream G: Physical file dump area

Tasks:

1. Add inspector bottom panel bound to C:\temp\x_file_dump.
2. Ensure directory creation on startup if missing.
3. Accept drag and drop from file list into dump area.
4. Default drop operation is Move, with conflict handling and clear progress feedback.
5. Show current staged file count and total size in the dump panel.

Likely files:

1. FileExplorer.App/Views/InspectorSidebarView.xaml (new)
2. FileExplorer.App/Views/InspectorSidebarView.xaml.cs (new)
3. FileExplorer.Core/Interfaces/IFileSystemService.cs (reuse existing operations)
4. FileExplorer.Core/ViewModels/MainViewModel.cs

Acceptance criteria:

1. Dragging items into dump area physically stages files under C:\temp\x_file_dump.
2. Operation progress and errors are visible.
3. Dump area persists between navigations.

## 3.8 Workstream H: Navigation hub deep fly-outs

Tasks:

1. Move shortcut launcher concepts into inspector edge hub with top-level favorite icons.
2. Keep hover-to-expand cascade for folder levels.
3. Click behavior:
   - folder click navigates file list
   - file click opens file
4. Keep prefetch and background caching for near-zero latency expansion.
5. Decommission the top-bar shortcut launcher after the sidebar hub reaches feature parity.

Likely files:

1. FileExplorer.App/MainWindow.xaml.cs (extract reusable flyout cache logic)
2. FileExplorer.App/Views/InspectorSidebarView.xaml.cs
3. Optional helper class for menu cache context (new)

Acceptance criteria:

1. Hover expansion is immediate for cached levels.
2. No UI hangs during deeper level expansion.
3. Click actions map exactly to required navigation/open behavior.

## 4. Performance and reliability strategy

1. Never block UI thread on directory enumeration or XML/ZIP parsing.
2. Use cancellation tokens for every async metadata request tied to current selection or current folder load.
3. Use bounded parallelism for extraction tasks.
4. Use memory cache keyed by path + last write timestamp for identity extraction.
5. Add telemetry logs:
   - breadcrumb sibling load latency
   - identity extraction latency
   - folder recursive stats latency
   - fly-out expansion latency

## 5. Testing plan

## 5.1 Unit tests

1. Identity extraction parser tests for docx and xlsx anchor variants.
2. Breadcrumb sibling enumeration tests.
3. Recursive stats parser tests from rq output.
4. Extension-to-color converter tests.

## 5.2 Integration tests

1. Navigate deep path with interactive breadcrumb and sibling flyouts.
2. Drag files to dump area and verify physical destination.
3. Hover expand fly-outs to level 5 without freeze.

## 5.3 Regression tests

1. Existing file list navigation and selection behavior.
2. Existing copy, move, delete, rename operations.
3. Existing tab and history behavior.

Likely test files:

1. FileExplorer.Tests/TabContentViewModelTests.cs (extend)
2. FileExplorer.Tests/FileSystemServiceTests.cs (extend)
3. New tests:
   - FileExplorer.Tests/IdentityExtractionServiceTests.cs
   - FileExplorer.Tests/BreadcrumbNavigationTests.cs
   - FileExplorer.Tests/RqUtilityServiceTests.cs

## 6. Delivery sequencing

Phase 1:

1. Layout refactor and file list column schema.
2. Remove tree and preview visuals.
3. Basic inspector shell and data contract.

Phase 2:

1. Interactive breadcrumb chevrons with async siblings.
2. Extension coloring and dark-mode validation.

Phase 3:

1. Identity extraction service implementation.
2. rq utility integration for recursive stats.
3. Inspector metadata hydration.

Phase 4:

1. Physical dump area drag and stage workflow.
2. Navigation hub fly-out integration and prefetch.

Phase 5:

1. Performance tuning and test completion.
2. Packaging and deployment validation.

## 7. Risks and mitigations

1. Legacy Office extraction reliability can vary by document internals.
   - Mitigation: use non-blocking fallback and leave fields empty when confidence is low.
2. xlsx coversheet detection inconsistency across workbook layouts.
   - Mitigation: add heuristic priority list and configurable anchor map.
3. Large directory metadata scanning can exceed target latency.
   - Mitigation: staged hydration, cancellation, and aggressive caching.
4. Fly-out over-expansion can spike IO.
   - Mitigation: depth and item caps with background prefetch windows.

## 8. Open questions requiring confirmation

No open questions remain for the current scope.

## 9. Definition of done

1. UI layout matches required two-pane architecture with no tree pane and no content preview.
2. Main list columns and extension coloring are exactly as specified.
3. Breadcrumb labels and separators provide required click behaviors asynchronously.
4. Inspector identity and metadata sections function with required field formats.
5. Physical dump area stages files to C:\temp\x_file_dump.
6. Navigation hub supports hover cascade and click actions with no perceived lag.
7. Tests pass and portable packaging script succeeds.
