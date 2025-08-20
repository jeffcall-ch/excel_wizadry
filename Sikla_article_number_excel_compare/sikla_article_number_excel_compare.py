import pandas as pd
import os

# Step 1: Read both sheets from the Excel file
excel_path = os.path.join(os.path.dirname(__file__), 'price_list_rev1_vs_sikla_bom_rev2.xlsx')
all_sheets = pd.read_excel(excel_path, sheet_name=None)

# Step 2: Name DataFrames after sheet names (whitespace removed)
dfs = {}
for sheet_name, df in all_sheets.items():
	clean_name = sheet_name.replace(' ', '')
	# Strip whitespace from all cell entries
	df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
	dfs[clean_name] = df

# Now, dfs contains DataFrames keyed by cleaned sheet names, with all cell entries stripped of leading/trailing whitespace.

# Step 3: Compare based on "ARTICLE NUMBER" and "PO QUANTITY"
from datetime import datetime
import numpy as np

# Identify baseline and update DataFrames
baseline_df = list(dfs.values())[0]
update_df = list(dfs.values())[1]


key_col = "ARTICLE NUMBER"
qty_col = "PO QUANTITY"

# Ensure ARTICLE NUMBER is treated as text
baseline_df[key_col] = baseline_df[key_col].astype(str)
update_df[key_col] = update_df[key_col].astype(str)


# Aggregate PO QUANTITY and preserve DESCRIPTION (first non-null)
desc_col = "DESCRIPTION"
def aggregate_df(df):
	agg_dict = {qty_col: 'sum'}
	if desc_col in df.columns:
		agg_dict[desc_col] = lambda x: next((v for v in x if pd.notna(v)), None)
	return df.groupby(key_col, as_index=False).agg(agg_dict)

baseline_df = aggregate_df(baseline_df)
update_df = aggregate_df(update_df)

baseline_df = baseline_df.set_index(key_col)
update_df = update_df.set_index(key_col)

all_articles = set(baseline_df.index).union(set(update_df.index))

desc_col = "DESCRIPTION"
rows = []
for article in all_articles:
	base_qty = baseline_df[qty_col].get(article, np.nan)
	upd_qty = update_df[qty_col].get(article, np.nan)
	# Get description from update tab if available, else from baseline
	upd_desc = None
	base_desc = None
	if desc_col in update_df.columns:
		try:
			upd_desc = update_df.loc[article, desc_col]
		except Exception:
			upd_desc = None
	if desc_col in baseline_df.columns:
		try:
			base_desc = baseline_df.loc[article, desc_col]
		except Exception:
			base_desc = None
	# Ensure base_qty and upd_qty are scalar values
	base_missing = pd.isna(base_qty) if not isinstance(base_qty, (pd.Series, pd.DataFrame)) else base_qty.isna().all()
	upd_missing = pd.isna(upd_qty) if not isinstance(upd_qty, (pd.Series, pd.DataFrame)) else upd_qty.isna().all()
	# Choose description: if only in baseline, use baseline; else use update
	if base_missing:
		desc = upd_desc
		rows.append({
			key_col: article,
			desc_col: desc,
			f"{qty_col} (baseline)": np.nan,
			f"{qty_col} (update)": upd_qty,
			"DIFFERENCE": upd_qty,
			"REMOVED": False
		})
	elif upd_missing:
		desc = base_desc
		rows.append({
			key_col: article,
			desc_col: desc,
			f"{qty_col} (baseline)": base_qty,
			f"{qty_col} (update)": np.nan,
			"DIFFERENCE": -base_qty,
			"REMOVED": True
		})
	elif base_qty != upd_qty:
		desc = upd_desc if upd_desc is not None else base_desc
		rows.append({
			key_col: article,
			desc_col: desc,
			f"{qty_col} (baseline)": base_qty,
			f"{qty_col} (update)": upd_qty,
			"DIFFERENCE": upd_qty - base_qty,
			"REMOVED": False
		})

result_df = pd.DataFrame(rows)

# Step 4: Save output to Excel file, color removed rows red

input_filename = os.path.basename(excel_path)
input_stem = os.path.splitext(input_filename)[0]
now_str = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
output_name = f"{input_stem}_comparison_processed_{now_str}.xlsx"
output_path = os.path.join(os.path.dirname(__file__), output_name)

with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:

	result_df.drop(columns=["REMOVED"]).to_excel(writer, index=False, sheet_name="Differences")
	workbook  = writer.book
	worksheet = writer.sheets["Differences"]

	# 1. Add filter to top row
	worksheet.autofilter(0, 0, len(result_df), len(result_df.drop(columns=["REMOVED"]).columns) - 1)

	# Split and freeze at B2
	worksheet.freeze_panes(1, 1)

	# 2. Center align all text in all cells
	center_format = workbook.add_format({'align': 'center'})
	worksheet.set_column(0, len(result_df.drop(columns=["REMOVED"]).columns) - 1, None, center_format)

	# 3. Auto width all columns
	for i, col in enumerate(result_df.drop(columns=["REMOVED"]).columns):
		max_len = max(
			result_df[col].astype(str).map(len).max(),
			len(col)
		) + 2
		worksheet.set_column(i, i, max_len, center_format)

	# Color removed rows red
	red_format = workbook.add_format({'bg_color': '#FFC7CE'})
	for idx, removed in enumerate(result_df["REMOVED"]):
		if removed:
			worksheet.set_row(idx+1, None, red_format)

print(f"Output saved to {output_path}")
