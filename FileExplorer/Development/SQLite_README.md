# SQLite Identity Cache

This document describes the SQLite metadata cache used by FileExplorer identity hydration.

## Database location

Default database path:

- Development/identity_cache.db

The resolved path can also be overridden with environment variable FILEEXPLORER_IDENTITY_CACHE_DB.

## Schema

Table: FileIdentityCache

Columns:

- FileKey TEXT PRIMARY KEY
- ProjectTitle TEXT
- Revision TEXT
- DocumentName TEXT
- LastWriteTime INTEGER
- FileSize INTEGER
- ParentPath TEXT

Indexes:

- IX_FileIdentityCache_ParentPath on ParentPath
- IX_FileIdentityCache_LastWriteTime on LastWriteTime

SQLite pragmas applied at startup:

- PRAGMA journal_mode = WAL;
- PRAGMA synchronous = NORMAL;

## Composite key logic

FileKey format:

normalizedFullPath|fileSize|lastWriteTime

Normalization rules:

- Strip long-path prefixes (\\?\ and \\?\UNC\)
- Convert to full path when possible
- Convert separators to standard directory separator
- Lowercase path

## Sync workflow

### Sync-on-open

On folder navigation:

1. Query SQLite rows where ParentPath equals current folder.
2. Build live keys from current file list.
3. Delete orphan keys that exist in SQLite but not in live set.
4. Apply cached identities immediately.
5. Extract identities for missing keys and upsert new cache rows.
6. Flush UI updates in debounced batches.

### Live watcher updates

When files are created/changed/renamed/deleted:

1. Update file list item.
2. Re-extract identity for created/changed/renamed files when supported.
3. Upsert cache row for extracted identity.
4. Delete cache rows by normalized path when files are removed.

## Noise filtering

Files are excluded from indexing when:

- Hidden or system attributes are set
- Name starts with ~$ (temporary Office files)
- Entry is a directory

## Troubleshooting

1. If cache looks stale, run scripts/ResetCache.ps1.
2. Re-open folder to trigger cache rebuild for visible entries.
3. Verify logs for deep-index and cache initialization messages.
