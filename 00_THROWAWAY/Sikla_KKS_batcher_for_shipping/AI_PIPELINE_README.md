# Sikla KKS Full Pipeline Guide for AI Assistants

This document is a full operational guide for rerunning and modifying the end-to-end batching workflow in this folder.

Scope covered by this guide:
- Base batching from BoM sheet
- Missing support detection against project status list
- C5 deduction
- Deleted and out-of-scope filtering
- Reintegration of in-scope missing supports
- UNKNOWN section normalization into N/A (configurable)
- N/A anchoring into a single batch per run
- Final arithmetic reconciliation
- Configurable building section sequence via INI

The workflow is now fully orchestrated by:
- run_full_batch_pipeline.py

Primary config file:
- batch_pipeline.ini

---

## 1. What the Pipeline Produces

Main output artifacts:
1. Step 1 batch file:
   - CA100-KVI-50296373_2.0 - BoM General Hangers and support material LS_erection_split_batched.xlsx
2. Step 2 missing report:
   - KKS_not_in_batch_report.xlsx
3. Step 3 integrated final batch file:
   - CA100-KVI-50296373_2.0 - BoM General Hangers and support material LS_erection_split_batched_with_missing_integrated.xlsx
4. Step 4 reconciliation report:
   - pipeline_reconciliation_report.xlsx

Reconciliation meaning:
- It compares project status /SU set against final batch /SU set after configured deductions.
- It confirms whether counts and exact sets match.

---

## 2. Scripts and Responsibilities

### 2.1 split_shipping_batches.py
Purpose:
- Builds initial 5 batches from the source BoM file.
- Keeps KKS and working area grouping rules.
- Applies section-priority sorting and batch ordering.

Key configurable inputs:
- input file and sheet
- number of batches
- tolerance
- section_order

### 2.2 reconcile_missing_kks.py
Purpose:
- Compares status list versus Step 1 batched file.
- Produces Missing_Final and diagnostic sheets.
- Deducts C5 in this stage and marks deleted.

### 2.3 integrate_missing_into_batches.py
Purpose:
- Takes Missing_Final and applies business exclusions.
- Integrates remaining supports into batches.
- Looks up support weight from ORIGINAL DATA file.
- Assigns outdoor rows (NO_AREA and other configured aliases) to N/A section.
- Can normalize UNKNOWN section rows into N/A.
- Keeps N/A section concentrated in one batch (single-batch anchor behavior).

Key configurable inputs:
- section_order
- na_working_areas
- na_section_label
- merge_unknown_to_na
- exclude_deleted
- exclude_bq_not_exist
- exclude_solid_not_gp

### 2.4 run_full_batch_pipeline.py
Purpose:
- One-command orchestrator for Step 1 to Step 4.
- Reads all settings from batch_pipeline.ini.
- Runs scripts in strict order.
- Writes final reconciliation report.

---

## 3. Step-by-Step Pipeline (0% to 100%)

### Step 0: Configuration
Edit batch_pipeline.ini:
- paths section: file paths and output names
- batching section: section order and balancing settings
- integration section: filtering policy for missing supports
- reconciliation section: deduction policy for final arithmetic check

### Step 1: Initial batching
run_full_batch_pipeline.py calls split_shipping_batches.py.

Output:
- step1_batched_output

### Step 2: Missing detection
run_full_batch_pipeline.py calls reconcile_missing_kks.py.

Output:
- step2_missing_report_output

### Step 3: Reintegration
run_full_batch_pipeline.py calls integrate_missing_into_batches.py.

Output:
- step3_final_output

### Step 4: Final compare
run_full_batch_pipeline.py performs set-based reconciliation and writes:
- step4_reconciliation_output

This is the final pass/fail for count and identity consistency.

---

## 4. Running the Full Flow

From this folder, run:

python run_full_batch_pipeline.py --config batch_pipeline.ini

Dry run mode (prints commands only):

python run_full_batch_pipeline.py --config batch_pipeline.ini --dry-run

---

## 5. How to Change Batch Sequence by Building Section

The most important control is:
- batching.section_order in batch_pipeline.ini

Example:
- section_order = UHA,UMA,N/A,UEA,UEB,UHQ

Behavior:
1. Batches containing first listed section are prioritized first.
2. Then next listed section, and so on.
3. Remaining sections not listed are placed afterward by fallback ordering.

Important:
- You do not need to edit Python code for ordering changes in normal use.
- Change only section_order in INI and rerun full pipeline.

---

## 6. Inclusion and Exclusion Policy Controls

Configured in integration section:
- exclude_deleted
- exclude_bq_not_exist
- exclude_solid_not_gp
- merge_unknown_to_na

Meaning:
- true: removed from integration scope
- false: kept and integrated

Configured outdoor aliases:
- na_working_areas
- na_section_label

If working area text matches any alias in na_working_areas, it is assigned to na_section_label during integration.

If merge_unknown_to_na is true, rows with unknown/unmatched section are also normalized to na_section_label.

N/A anchoring rule:
- N/A (including merged UNKNOWN when enabled) is assigned to one anchor batch in each run.
- This prevents N/A from being spread across multiple batches unless the code is intentionally changed.

Current default includes:
- NO_AREA
- BQ NOT EXIST

---

## 7. Final Arithmetic Reconciliation Controls

Configured in reconciliation section:
- deduct_deleted
- deduct_c5
- deduct_solid_not_gp
- deduct_bq_not_exist

This section controls only Step 4 arithmetic check, not integration logic.

Typical usage:
- If BQ not exist is included in batches, keep deduct_bq_not_exist = false.
- If BQ not exist is excluded from batches, set deduct_bq_not_exist = true.

---

## 8. Expected Quality Gates After Every Run

Check the Step 4 report first:
- pipeline_reconciliation_report.xlsx, sheet Summary

Required for clean run:
- Expected minus final = 0
- Final minus expected = 0
- Set match = True

Secondary checks:
- Step 3 Overview sheet has expected integration counts.
- Summary_Updated sheet weight shifts are reasonable.
- N/A distribution is as intended (single batch by default in current implementation).

---

## 9. AI Assistant Playbook for Future Change Requests

When user asks for sequence or scope changes, follow this order:
1. Identify whether change is configuration-only or code-level.
2. For configuration-only cases:
   - edit batch_pipeline.ini
   - rerun wrapper script
   - report Step 4 match status
3. For code-level cases:
   - update smallest relevant script only
   - rerun full wrapper pipeline
   - compare before/after overview metrics
4. Always provide:
   - updated output file names
   - final reconciliation summary
   - any changed config keys
   - N/A and UNKNOWN distribution summary

Examples of configuration-only requests:
- Change first two batch sections
- Include/exclude BQ not exist
- Include/exclude deleted
- Update outdoor alias list

Examples likely requiring code changes:
- New matching rules for KKS normalization
- New weight lookup source format
- New custom batching strategy beyond current fallback logic

---

## 10. Common Pitfalls and Fixes

### 10.1 KKS mismatch due to leading slash
Symptom:
- All status supports appear missing.

Fix:
- Ensure normalization keeps canonical leading slash for /SU identifiers.
- Already handled in scripts by normalize_kks.

### 10.2 N/A appears as NaN in pandas
Symptom:
- Reading output with pandas shows NaN where N/A was expected.

Fix:
- Use keep_default_na=False when reading for validation.
- Excel cell value remains literal N/A.

### 10.3 Noisy terminal command output
Symptom:
- Multi-line shell command parsing issues.

Fix:
- Prefer wrapper script execution instead of manual one-off long commands.

### 10.4 Reconciliation mismatch after policy change
Symptom:
- Step 4 Set match = False.

Fix:
1. Verify integration exclusion flags.
2. Verify reconciliation deduction flags.
3. Ensure both represent the same business policy.

---

## 11. File Inventory (Core)

- split_shipping_batches.py
- reconcile_missing_kks.py
- integrate_missing_into_batches.py
- run_full_batch_pipeline.py
- batch_pipeline.ini
- AI_PIPELINE_README.md

---

## 12. Suggested Operating Routine

For each new revision cycle:
1. Update input files in paths section if file names changed.
2. Update section_order and policy flags in INI.
3. Run wrapper script once.
4. Review Step 4 report first.
5. If mismatch exists, inspect Step 2 and Step 3 detail sheets.

This keeps the process reproducible and auditable.

---

## 13. Mandatory Assistant Run Summary (After Every Execution)

After each full run (or partial rerun), the assistant should always respond with this summary structure:

1. Run context:
- Config file used
- Key section_order value
- Integration policy switches used (deleted, BQ not exist, solid not GP)
- UNKNOWN to N/A setting

2. Output artifacts:
- Step 1 file path
- Step 2 file path
- Step 3 final supplier file path
- Step 4 reconciliation report path

3. Numeric validation snapshot:
- Status unique /SU
- Final unique /SU
- Deduction union
- Expected after deduction
- Expected minus final
- Final minus expected
- Set match (True/False)

4. Integration snapshot:
- Missing_Final rows
- Excluded total and breakdown
- Included for integration
- Integrated rows added
- Integrated with weight found / missing

5. Section distribution checks:
- UNKNOWN row count in final file
- N/A row count in final file
- N/A per-batch distribution

6. Conclusion:
- One-line statement whether result is communication-ready for supplier.
