# System Priority Batching Method - Full AI Assistant README

## 1. Purpose

This document is the complete operational guide for the system-based batching method implemented by the script:

- split_shipping_batches_by_system.py

This method creates shipping batches from support structures ending with /SU, using:

1. Resolved system mapping from su_system_sequence_resolved.csv.
2. Aggregated support weight from the ORIGINAL workbook sheet Per_Support_Structure.

The method is designed to prevent installation sequence conflicts by assigning and batching in strict system priority order.

---

## 2. Scope and Outputs

The process produces one workbook:

- su_system_batched.xlsx

Workbook sheets:

1. Detail
2. Summary

Batch count target:

- 5 batches (configurable)

Weight limit rule:

- Target batch weight is 4000 kg.
- Maximum allowed batch weight is target plus 25 percent, which is 5000 kg.
- This method enforces the maximum limit.

Important:

- This method does not enforce a minimum batch fill threshold.
- A batch below 4000 kg is allowed if required by sequence and max-limit constraints.

---

## 3. Input Files

## 3.1 Required CSV input

File:

- su_system_sequence_resolved.csv

Required columns:

1. support_structure_su
2. resolved_system_first_in_priority

Optional columns used when available:

1. original_owner_system
2. all_systems_mapped_priority_order
3. system_changed_for_schedule_risk

The script keeps one row per support_structure_su (first occurrence if duplicates exist).

## 3.2 Required workbook input

File:

- CA100-KVI-50296373_2.0 - BoM General Hangers and support material ORIGINAL DATA .xlsm

Sheet:

- Per_Support_Structure

Required columns in that sheet:

1. Support Structure
2. Total Weight [kg]

The script auto-detects the header row by scanning the top rows and finding both required column names.

## 3.3 Configuration input

File:

- system_batching.ini

Sections:

1. paths
2. batching
3. systems
4. overrides

---

## 4. Final Priority List

Current final sequence is:

1. Instrument air
2. Service air
3. Compressed air
4. Nitrogen
5. Process Water
6. Towns Water
7. Bottom ash
8. Drain water
9. Blow down
10. Exhaust Air
11. Condensate drains
12. HP water
13. Fuel oil
14. Ammonia water

This exact order controls batching flow.

---

## 5. Critical Business Rules

## 5.1 Weight source rule

Support weight is never taken from the resolved CSV.

Weight is always calculated from Per_Support_Structure by:

1. Filtering to supports ending with /SU.
2. Grouping by Support Structure.
3. Summing all Total Weight [kg] rows per support.

So if one support has 5 to 6 rows in Per_Support_Structure, all rows are aggregated into one support total.

## 5.2 Priority flow rule

Batching is strict top-down by priority list:

1. Start from first system in priority list.
2. Fill current batch until adding more would exceed max 5000 kg.
3. Continue in next batch.
4. Move down priority list in order.

This means batch 1 should contain earliest-priority systems first, not a weight-rebalanced mix from late systems.

## 5.3 System split rule

A system should not be split if possible.

Implementation behavior:

1. If a full system fits as a whole into the current batch (or next batch after rollover), keep it whole.
2. Split a system only when unavoidable due to max batch weight.

## 5.4 Alphabetical order rule

Inside each system, supports are sorted alphabetically by KKS code.

## 5.5 Override rule

Support-level overrides from INI are applied after reading CSV system assignment.

Example currently configured:

1. 1HJF10BQ100/SU = Fuel oil
2. 1HJF20BQ100/SU = Fuel oil

This guarantees those supports are treated as Fuel oil in batching.

## 5.6 KKS normalization rule

Support code normalization for joins and comparisons:

1. Trim whitespace.
2. Uppercase.
3. Remove leading slash if present.

This avoids slash-format mismatch between sources.

---

## 6. Algorithm Walkthrough

## Step 1: Load config

Reads paths, batch settings, priority list, and optional overrides.

## Step 2: Build weight lookup

From ORIGINAL workbook Per_Support_Structure:

1. Detect header row dynamically.
2. Read Support Structure and Total Weight [kg].
3. Keep only /SU supports.
4. Aggregate total weight per support.

## Step 3: Load resolved scope

From su_system_sequence_resolved.csv:

1. Validate required columns.
2. Normalize support codes.
3. Keep first row per support.
4. Join aggregated weight.
5. Normalize system text.
6. Replace blank system with UNKNOWN.
7. Apply support-level overrides.
8. Compute system rank by priority.

## Step 4: Sequential batch assignment

For each system in priority order:

1. Sort supports alphabetically.
2. Compute full system total.
3. If system total fits current batch max, place whole system.
4. Otherwise place supports one by one until max reached, then continue in next batch.
5. If any single support exceeds max limit, fail with clear error.

## Step 5: Build output tables

Detail sheet columns:

1. KKS code
2. System
3. KKS Total Weight [kg]
4. Batch number
5. Original owner system
6. All systems
7. System changed for schedule risk

Summary sheet top table:

1. Batch number
2. Total_Weight_kg
3. Systems
4. System_Count
5. KKS_Count
6. Target_Weight_kg
7. Min_Allowed_kg
8. Max_Allowed_kg
9. Delta_vs_Target_kg
10. Within_Tolerance

Summary sheet lower table:

1. Batch number
2. System
3. System_Total_Weight_kg

## Step 6: Excel formatting

Applies:

1. Detail filter and freeze pane.
2. Column widths.
3. Subtotal formula in the Detail sheet header area and bottom line.
4. Summary column widths.

---

## 7. Configuration Reference

File:

- system_batching.ini

## 7.1 paths section

1. resolved_csv
2. original_file
3. original_sheet
4. output_file

## 7.2 batching section

1. batch_count
2. target_weight_kg
3. tolerance

Max allowed is computed as:

- target_weight_kg times one plus tolerance.

Default with 4000 and 0.25 gives 5000.

## 7.3 systems section

1. priority

Comma-separated, strict order.

## 7.4 overrides section

Format:

- SUPPORT_KKS = System Name

Applied per support before assignment.

Use this for known data exceptions.

---

## 8. Run Instructions

From the batching folder, run:

python split_shipping_batches_by_system.py --config system_batching.ini

If Python environment is local venv, equivalent executable form can be used.

---

## 9. Current Known Behavior and Expectations

Based on current dataset and current config:

1. One early-priority system may occupy more than one batch if needed by max limit.
2. Lower-priority systems can be grouped into later batches after higher-priority systems are placed.
3. This is expected and aligns with strict sequence-first logic.

---

## 10. Validation Checklist After Every Run

Always validate the following:

1. Output file exists and has Detail and Summary sheets.
2. Row count in Detail equals number of unique /SU supports in source scope.
3. Every support has a batch number.
4. No batch exceeds max allowed weight.
5. System order progression across batches follows priority flow.
6. Override supports are assigned to expected system.
7. Systems split across batches are only those that cannot stay whole under max limit.

---

## 11. AI Assistant Response Template (Mandatory)

After each run, provide:

1. Run context
- Config file used
- Priority list used
- Override entries used

2. Artifacts
- Output workbook path

3. Numeric summary
- Supports batched count
- Batch totals
- Max limit
- Any batch over limit (yes or no)

4. Sequence summary
- First systems appearing in Batch 1, Batch 2, etc.
- Number of systems split across batches

5. Override verification
- List of override supports with resulting system and batch

6. Conclusion
- One line stating whether output is ready for use

---

## 12. Troubleshooting

## 12.1 Missing column errors

Symptom:

- Script says missing required columns.

Action:

1. Confirm column names in input files.
2. Confirm correct sheet name in INI.

## 12.2 Header row mismatch in workbook

Symptom:

- pandas cannot find required usecols.

Action:

1. Keep dynamic header detection in place.
2. Ensure required names Support Structure and Total Weight [kg] still exist in sheet.

## 12.3 Unassigned supports

Symptom:

- Error listing unresolved support codes.

Action:

1. Check join keys in CSV and workbook.
2. Check KKS formatting and slash normalization.

## 12.4 Single support heavier than max

Symptom:

- Script fails with heavy support list.

Action:

1. Increase max limit (target or tolerance), or
2. Handle that support by special business exception.

## 12.5 Unexpected UNKNOWN system

Symptom:

- UNKNOWN appears in Detail or Summary.

Action:

1. Correct upstream mapping in su_system_sequence_resolved.csv, or
2. Add support-specific override in INI.

---

## 13. Change Management Rules for Future Edits

When changing this method:

1. Keep weight source fixed to Per_Support_Structure aggregation.
2. Keep sequential top-down priority assignment unless explicitly requested otherwise.
3. Do not introduce balancing logic that reorders systems against priority.
4. Keep max limit enforcement strict.
5. Keep support-level override capability.
6. Re-run and report validation summary after every change.

---

## 14. Quick Summary

This method is intentionally sequence-first, not equal-weight-first.

It ensures:

1. Priority list drives batch progression.
2. Max batch weight is never exceeded.
3. Systems are kept whole where possible.
4. System splitting happens only when unavoidable.
5. Weight is calculated from authoritative structural rows in Per_Support_Structure.
