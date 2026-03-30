# Price Sheet Workflow Summary

## Goal
Build a repeatable workflow that reads the source BoM Excel file, stores clean data in SQLite, prepares an aggregated table for price sheet logic, and exports the aggregated result back to Excel.

## Input
- Source file: `CA100-KVI-50296373_0.0 - BoM General Hangers and support material.xlsm`
- Source sheet to process: `Total Qty`

## Import Rules (Excel -> SQLite)
1. Find the header row dynamically in `Total Qty` by matching required header titles:
   - `Item Number`
   - `Qty`
   - `Description` (or `Decscription` typo if present)
2. Skip the empty row that usually appears directly below the header.
3. Read all rows after that until the last useful (non-empty) row.
4. Normalize values during import:
   - Remove units from numeric text like `kg`, `m`, `mm`.
   - Keep only numeric value when unit is present.
   - Preserve numeric type in SQLite (not text when numeric).
   - For decimal numbers, keep two decimal precision.
5. Save imported rows into table `RAW_DATA_REV0`.

## Database Tables
- `RAW_DATA_REV0`: raw imported and cleaned rows from `Total Qty`.
- `AGGREGATED_DATA_REV0`: structure reserved for aggregation/calculation output for price sheet export.

## Revision Concept
- `REV0` is the first revision tag.
- Future runs may use `REV1`, `REV2`, etc., creating corresponding revision-specific tables.

## Export Rules (SQLite -> Excel)
1. Read from `AGGREGATED_DATA_REV0`.
2. Export to an Excel file in the same folder.
3. Keep decimal values formatted with two decimal places in the output workbook.

## Current Scope
- Import and base table creation are defined now.
- Aggregation definitions (how `AGGREGATED_DATA_REV0` is filled) will be defined later.