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

---

# Enhancement Series 2 Plan: UI Stability, SQLite Caching, and Deep Indexing

Date: 2026-04-06
Location: Wuerenlingen, Switzerland
Engineering Target: Deterministic zero-flicker UI with self-cleaning SQLite cache and sub-100 ms folder hydration on warm cache.

## S2-0. Scope

This series implements:

1. UI anti-flicker hardening for large folders.
2. SQLite metadata cache layer and sync lifecycle.
3. Recursive deep indexing for favorites using rq_search integration.
4. Defensive handling for locks, long paths, and noise files.
5. Developer reset and documentation utilities.

Out of scope for this series:

1. OCR extraction.
2. Search ranking redesign.
3. New visual theme redesign.

## S2-1. Architecture delta

New components:

1. IIdentityCacheService and SqliteIdentityCacheService for read/write cache operations.
2. FolderIdentitySyncCoordinator for Sync-on-Open reconciliation logic.
3. DeepIndexBackgroundService for recursive favorites indexing.
4. UiHydrationBuffer in tab/view model layer for debounced batch apply.
5. FavoritesEngineResolver for rq_search binary path validation.

Data flow after this series:

1. Folder open -> disk snapshot and SQLite snapshot -> reconcile keys.
2. Existing keys hydrate instantly from SQLite.
3. Missing keys go to sniper extraction queue.
4. Buffered metadata flush updates UI in controlled batches.
5. FileSystemWatcher updates both UI and SQLite incrementally.

## S2-2. UI Stability and anti-flicker plan

### S2-2.1 Virtualization

Tasks:

1. Ensure Details View list uses virtualization-friendly host and item container settings.
2. Confirm recycling behavior remains active under sorting and incremental updates.
3. Validate container reuse during rapid scrolling with 10k+ rows.

Acceptance:

1. No full list re-layout on metadata flush.
2. Scroll FPS remains stable while hydration runs.

### S2-2.2 Sorting and selection lock

Tasks:

1. Introduce a hydration window lock that suppresses re-sort triggers caused by Project/Revision updates.
2. Preserve current selected item and anchored top visible row during hydration.
3. Re-enable sort only after flush cycle completion or on explicit header click.

Acceptance:

1. No jump in scroll position while metadata arrives.
2. Multi-select set does not mutate during background updates.

### S2-2.3 Batch UI updates (250 ms debounce, 50+ batch)

Tasks:

1. Buffer extraction results in memory and debounce apply by 250 ms.
2. Flush in chunks of at least 50 items when available.
3. Coalesce repeated updates for same file key before each flush.

Acceptance:

1. UI updates occur in discrete batches, not per-item flicker.
2. CPU and Dispatcher utilization drop versus per-item apply.

### S2-2.4 Multi-selection sidebar behavior

Tasks:

1. Detect selection count greater than 1 in inspector binding flow.
2. Hide single-file Identity Card fields in multi-select mode.
3. Show aggregate stats block:
   - selected item count
   - total size
   - count by file extension/type

Acceptance:

1. Sidebar no longer cycles identity values during multi-select.
2. Aggregate numbers remain stable while hydration continues.

## S2-3. SQLite metadata cache layer

### S2-3.1 Database location

Target path:

1. Repository/Development/identity_cache.db

Tasks:

1. Add configurable root for cache location with sane default above.
2. Auto-create directory and DB on first run.

### S2-3.2 Schema

Table: FileIdentityCache

Columns:

1. FileKey (TEXT PRIMARY KEY)
2. ProjectTitle (TEXT)
3. Revision (TEXT)
4. DocumentName (TEXT)
5. LastWriteTime (INTEGER or TEXT UTC ticks/iso)
6. FileSize (INTEGER)
7. ParentPath (TEXT)

Indexes:

1. IX_FileIdentityCache_ParentPath on ParentPath
2. Optional IX_FileIdentityCache_LastWriteTime for maintenance jobs

### S2-3.3 Composite key

Rule:

1. FileKey = normalizedFullPath|fileSize|lastWriteTime
2. normalizedFullPath = full path lowercased and normalized separators

Tasks:

1. Centralize key generation in one helper shared by sync and watcher code.
2. Add tests for case-insensitive path variants and rename scenarios.

### S2-3.4 SQLite performance pragmas

On initialization:

1. PRAGMA journal_mode = WAL;
2. PRAGMA synchronous = NORMAL;

Acceptance:

1. Concurrent read/write stable under hydration and watcher events.
2. No UI thread blocking on DB operations.

## S2-4. Cache lifecycle and watcher sync

### S2-4.1 Sync-on-Open (folder entry)

Workflow per folder open:

1. Query SQLite keys where ParentPath equals current folder.
2. Scan disk to produce live key set.
3. Delete orphans in batched delete statements.
4. Queue new keys for sniper extraction.
5. Upsert extraction results when done.

Acceptance:

1. Re-entering folder hydrates from cache in under 100 ms on warm cache target hardware.
2. Orphans are removed deterministically.

### S2-4.2 FileSystemWatcher integration

Tasks:

1. Watch current folder for create, rename, delete, and changed events.
2. Reflect watcher events in UI list and SQLite cache without full refresh.
3. Debounce duplicate watcher events and handle rename as atomic key move.

Acceptance:

1. External file edits appear in UI and cache automatically.
2. No folder re-entry required for consistency.

## S2-5. Deep indexing with rq_search (C++ integration)

### S2-5.1 Discovery

Tasks:

1. Enumerate all top-level favorites from Navigation Hub.
2. Use rq_search recursively to discover all files under each favorite root.
3. Feed discovered paths into same key reconciliation pipeline.

### S2-5.2 Engine config and validation

Tasks:

1. Resolve rq_search path from Config.json first, then environment variable fallback.
2. Validate executable existence and version compatibility at startup.
3. If missing, show non-blocking warning in inspector sidebar.

### S2-5.3 Recursive background sync

Tasks:

1. Run deep sync in background with BelowNormal CPU priority.
2. Throttle IO and DB batch sizes to avoid degrading foreground navigation.
3. Expose progress summary in logs and optional sidebar status line.

Acceptance:

1. Favorite tree is indexed progressively without UI stall.
2. Foreground folder navigation remains responsive during deep sync.

## S2-6. Defensive engineering and edge cases

### S2-6.1 File locking

Tasks:

1. Continue opening files with FileShare.ReadWrite where supported.
2. On lock or parse failure, mark retryable and skip without crashing worker.
3. Retry on next refresh cycle or watcher change.

### S2-6.2 Long paths

Tasks:

1. Normalize all file IO and cache key generation for long-path-safe handling.
2. Validate SQLite key generation and watcher behavior for paths beyond 260 chars.

### S2-6.3 Noise filtering

Tasks:

1. Exclude hidden/system files from indexing scope.
2. Exclude temporary Office artifacts with prefix ~$.
3. Exclude known transient lock/cache suffix patterns where appropriate.

Acceptance:

1. Cache contains only relevant files.
2. Temp or system noise does not cause churn.

## S2-7. Developer utilities

### S2-7.1 ResetCache script

Deliverable:

1. scripts/ResetCache.ps1

Behavior:

1. Stop app process if needed.
2. Delete identity_cache.db, identity_cache.db-shm, identity_cache.db-wal.
3. Print clear before/after status.

### S2-7.2 SQLite README

Deliverable:

1. Development/SQLite_README.md

Must document:

1. Schema and indexes.
2. Composite key formula.
3. Sync-on-Open reconciliation flow.
4. Watcher update flow.
5. Troubleshooting and reset steps.

## S2-8. Delivery sequencing

Phase A: UI stability foundation

1. Virtualization verification and sorting lock.
2. Debounced batch flush pipeline.
3. Multi-selection aggregate sidebar mode.

Phase B: SQLite cache core

1. DB bootstrap, pragmas, schema, and repository location.
2. Cache read path for folder open.
3. Key builder and upsert pipeline.

Phase C: Sync lifecycle

1. Sync-on-Open reconcile algorithm.
2. Orphan delete and new-key extraction queue.
3. Watcher-driven incremental cache updates.

Phase D: Deep indexing

1. rq_search engine resolver and warnings.
2. Recursive favorites scan and background sync scheduler.
3. BelowNormal priority and throttling tuning.

Phase E: Defensive hardening and utilities

1. Lock handling and retry policy.
2. Long-path and noise-filter tests.
3. Reset script and SQLite README finalization.

## S2-9. Test and validation plan

Unit tests:

1. FileKey generation normalization and collisions.
2. SQLite upsert, delete orphan batch, and ParentPath query performance.
3. Debounce buffer behavior and batch size guarantees.
4. Noise filtering rules and long-path handling.

Integration tests:

1. Open folder warm cache hydration latency target.
2. External add/rename/delete reflected in UI and DB.
3. Multi-select inspector aggregate mode stability.
4. Deep indexing run while interacting with UI.

Performance gates:

1. Warm-cache folder hydration median below 100 ms for representative folders.
2. No perceptible list jump during metadata hydration.
3. CPU budget stable during deep indexing background run.

## S2-10. Definition of done

1. Details list remains stable with no hydration-induced flicker or scroll jump.
2. SQLite cache is authoritative, self-cleaning, and consistent with disk state.
3. Sync-on-Open and FileSystemWatcher keep UI and cache aligned automatically.
4. Deep indexing covers all favorites using rq_search with robust engine validation.
5. Lock, long-path, and noise scenarios are handled without crashes.
6. Reset script and SQLite README are available for developers.
