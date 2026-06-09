# Sikla BoM Processor

Primary and secondary pipe support weight analysis for Sikla BoM data.

## What this script does

The script reads one Excel workbook and produces a formatted summary workbook with:

- Small Bore (SB) and Large Bore (LB) primary support totals
- SB and LB secondary support totals
- Combined totals
- Data statistics and anomaly notes
- Validation against the Total Qty sheet

The classification rule is:

- SB: Nominal Bore <= 40
- LB: Nominal Bore > 40

Unknown or non-parsable nominal bore values are excluded from the SB/LB split, but still included in the validation check.

## Input workbook and expected sheets

The script is currently configured to process this filename:

CA100-KVI-50296373_2.0 - BoM General Hangers and support material.xlsm

Expected sheets:

- Per_Pipe_Support
- Support_List
- Total Qty

Data assumptions:

- Header row is Excel row 7 (pandas header index 6)
- Blank separator rows may exist and are removed
- Required columns include:
	- Per_Pipe_Support: Pipe Support, Support Structure, Nominal Bore, Total Weight [kg]
	- Support_List: Support Structure, Nominal Bore, Total Weight [kg]
	- Total Qty: Total Weight [kg]

## Processing logic summary

1. Per_Pipe_Support (primary supports)
	 - Groups by Pipe Support (KKS code)
	 - Calculates max nominal bore per KKS and total primary weight per KKS
	 - Classifies each KKS as SB/LB/Unknown
	 - Aggregates primary weight per Support Structure

2. Support_List (total structure weight)
	 - Parses nominal bore values (single value or comma-separated list)
	 - Uses max nominal bore per Support Structure for SB/LB classification
	 - Derives secondary weight per structure:
		 secondary = total_wt_sl - primary_wt

3. Total Qty (validation)
	 - Sums Total Weight [kg] for workbook grand total comparison
	 - Compares calculated totals to Total Qty
	 - Status is OK when absolute difference is under 1.0 kg, otherwise CHECK

## Output

The script generates a timestamped Excel report in the same directory as the input workbook, with a name like:

Support_Weight_Analysis_<input_stem>_<YYYY-MM-DD_HHMMSS>.xlsx

The output workbook contains a Summary sheet with:

- Support weight summary table (SB/LB primary, secondary, and total)
- Data statistics (structure counts, excluded unknowns)
- Validation section (calculated vs Total Qty with status)

## Requirements

- Python 3.9+
- Python packages:
	- pandas
	- numpy
	- openpyxl

You can install dependencies from the repository root using:

pip install -r requirements.txt

## How to run

From this folder:

Sikla_BoM_processor_for_Erection_BoQ_doc

run:

python sikla_bom_processor.py

## Notes

- The script uses a hardcoded input filename in the FILE variable.
- If the workbook is renamed or moved, update the FILE value accordingly.
- Console output includes quick totals and warnings for data anomalies, such as structures missing between sheets or mixed nominal bore values inside one KKS group.