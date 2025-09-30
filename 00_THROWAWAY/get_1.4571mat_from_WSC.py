# -*- coding: utf-8 -*-
"""Extract 1.4571 material weights from BOQ sheets in Excel files.

Behaviour:
- Scans the script folder for `.xlsx` and `.xlsm` files.
- For each file, finds sheets with 'BOQ' in the sheet name (case-insensitive).
- Locates columns for `DN1`, `Material`, and `Total weight [kg]` (case-insensitive, substring match).
- Filters rows where `Material` indicates 1.4571 (accepts `1.4571` and `1,4571` with surrounding text).
- Parses `Total weight [kg]` robustly (handles `,` as decimal, thousand separators, spaces).
- Aggregates total weight per `DN1` for each file.
- Appends a summary row listing distinct DN1 values and the total weight across them.
- Writes a CSV named `get_1.4571mat_from_WSC_output.csv` with columns: `filename,DN1,total_weight_kg`.

Requirements: `pandas`, `openpyxl`.
"""

from __future__ import annotations

import csv
import os
import re
from pathlib import Path
from typing import Iterable, List, Optional

import pandas as pd


OUT_CSV = "get_1.4571mat_from_WSC_output.csv"
OUT_XLSX = "get_1.4571mat_from_WSC_output.xlsx"
TARGET_MAT_REGEX = re.compile(r"(?<!\d)(1[\.,]4571)(?!\d)", flags=re.IGNORECASE)


def find_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
	cols = list(df.columns)
	lc = [c.lower().strip() for c in cols]
	for target in candidates:
		t = target.lower()
		# exact match first
		for i, c in enumerate(lc):
			if c == t:
				return cols[i]
		# substring match
		for i, c in enumerate(lc):
			if t in c:
				return cols[i]
	return None


def parse_number(val) -> Optional[float]:
	if pd.isna(val):
		return None
	s = str(val).strip()
	if s == "":
		return None
	# Extract first number-like token
	m = re.search(r"-?[\d.,]+", s)
	if not m:
		return None
	token = m.group(0)
	# Remove spaces
	token = token.replace("\xa0", "").replace(" ", "")
	# If contains both '.' and ',': decide which is decimal by last occurrence
	if "." in token and "," in token:
		if token.rfind(",") > token.rfind("."):
			# comma looks like decimal separator: remove dots (thousands), replace comma
			token = token.replace(".", "").replace(",", ".")
		else:
			# dot is decimal, remove commas
			token = token.replace(",", "")
	else:
		# only commas -> treat as decimal separator
		if "," in token and "." not in token:
			token = token.replace(",", ".")
	try:
		return float(token)
	except Exception:
		return None


def material_is_14571(val) -> bool:
	if pd.isna(val):
		return False
	s = str(val).strip()
	if s == "":
		return False
	if TARGET_MAT_REGEX.search(s):
		return True
	# also try numeric equality
	num = parse_number(s)
	if num is None:
		return False
	return abs(num - 1.4571) < 1e-6


def process_file(path: Path) -> List[List]:
	"""Process a single workbook and return rows: [filename, DN1, total_weight]

	If no BOQ sheets found, returns a single row with sheetname '<no-BOQ-sheet>' and zeros.
	"""
	rows: List[List] = []
	try:
		xls = pd.ExcelFile(path, engine="openpyxl")
	except Exception as e:
		print(f"Failed to open {path.name}: {e}")
		return []

	boq_sheets = [s for s in xls.sheet_names if "boq" in s.lower()]
	if not boq_sheets:
		# signal no BOQ sheet
		rows.append([path.name, "<no-BOQ-sheet>", 0.0])
		return rows

	all_matches = []
	for sheet in boq_sheets:
		try:
			df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
		except Exception as e:
			print(f"Failed to read sheet {sheet} in {path.name}: {e}")
			continue

		# Find columns
		dn_col = find_column(df, ["dn1", "dn 1", "dn", "nominal diameter"]) or find_column(df, ["dn1"])
		mat_col = find_column(df, ["material", "mat.", "material code"]) or find_column(df, ["material"])
		weight_col = find_column(df, ["total weight [kg]", "total weight", "weight [kg]", "total weight kg"]) or find_column(df, ["total weight [kg]"])

		if not dn_col or not mat_col or not weight_col:
			print(f"Sheet {sheet} in {path.name} missing one of required columns (DN1/Material/Total weight [kg])")
			continue

		# Select and clean
		working = df[[dn_col, mat_col, weight_col]].copy()
		working.columns = ["DN1", "Material", "TotalWeight"]

		# Normalize DN1 as string
		working["DN1"] = working["DN1"].apply(lambda x: str(x).strip() if not pd.isna(x) else "")

		# Filter material
		working = working[working["Material"].apply(material_is_14571)]
		if working.empty:
			continue

		# Parse weights
		working["_w"] = working["TotalWeight"].apply(parse_number)
		working = working[working["_w"].notna()]
		if working.empty:
			continue

		# Aggregate per DN1
		grp = working.groupby("DN1", dropna=False) ["_w"].sum().reset_index()
		grp = grp.rename(columns={"_w": "total_weight_kg"})

		for _, r in grp.iterrows():
			rows.append([path.name, r["DN1"], float(r["total_weight_kg"])])

		# collect for file-level summary if needed
		all_matches.append(grp)

	# If we collected matches across multiple BOQ sheets, add a summary row
	if all_matches:
		combined = pd.concat(all_matches, ignore_index=True)
		# combined may have duplicate DN1 entries
		summary = combined.groupby("DN1")["total_weight_kg"].sum().reset_index()
		dn_list = list(summary["DN1"].astype(str).unique())
		total_sum = float(summary["total_weight_kg"].sum())
		rows.append([path.name, "SUMMARY: " + ",".join(dn_list), total_sum])

	return rows


def main():
	folder = Path(__file__).resolve().parent
	out_path = folder / OUT_CSV
	files = sorted(folder.iterdir())
	all_rows: List[List] = []
	for p in files:
		if not p.is_file():
			continue
		if p.suffix.lower() not in (".xlsx", ".xlsm"):
			continue
		print(f"Processing: {p.name}")
		rows = process_file(p)
		all_rows.extend(rows)

	# Split rows into data rows (non-summary) and summary rows
	data_rows = [r for r in all_rows if not (isinstance(r[1], str) and r[1].upper().startswith("SUMMARY:"))]
	summary_rows = [r for r in all_rows if (isinstance(r[1], str) and r[1].upper().startswith("SUMMARY:"))]

	# Write CSV for compatibility
	with out_path.open("w", newline="", encoding="utf-8") as f:
		writer = csv.writer(f)
		writer.writerow(["filename", "DN1", "total_weight_kg"])
		for r in data_rows:
			writer.writerow(r)

	# Also write XLSX with two sheets: Data and SUMMARY
	out_xlsx_path = folder / OUT_XLSX
	with pd.ExcelWriter(out_xlsx_path, engine="openpyxl") as xw:
		if data_rows:
			df_data = pd.DataFrame(data_rows, columns=["filename", "DN1", "total_weight_kg"])
		else:
			df_data = pd.DataFrame(columns=["filename", "DN1", "total_weight_kg"])
		df_data.to_excel(xw, sheet_name="Data", index=False)

		if summary_rows:
			df_summary = pd.DataFrame(summary_rows, columns=["filename", "DN1", "total_weight_kg"])
		else:
			df_summary = pd.DataFrame(columns=["filename", "DN1", "total_weight_kg"])
		df_summary.to_excel(xw, sheet_name="SUMMARY", index=False)

	print(f"Wrote outputs: {out_path} (rows: {len(data_rows)}) and {out_xlsx_path} (summary rows: {len(summary_rows)})")


if __name__ == "__main__":
	main()

