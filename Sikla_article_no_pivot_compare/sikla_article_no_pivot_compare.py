import pandas as pd
import os

# Path to the Excel file
excel_file = os.path.join(os.path.dirname(__file__), "extracted_tables_20250829_122546_ARTICLE_COMPARE.xlsx")


# Read old, new, and descriptions sheets
old_df = pd.read_excel(excel_file, sheet_name=0, engine="openpyxl")
new_df = pd.read_excel(excel_file, sheet_name=1, engine="openpyxl")
desc_df = pd.read_excel(excel_file, sheet_name=2, engine="openpyxl")

# Ensure columns are named consistently
old_df = old_df.rename(columns={"Article No": "Article No", "Sum of QTY": "Sum of QTY"})
new_df = new_df.rename(columns={"Article No": "Article No", "Sum of QTY": "Sum of QTY"})




# Strip whitespace and set Article No as index, ensure all are strings
old_df["Article No"] = old_df["Article No"].astype(str).str.strip()
new_df["Article No"] = new_df["Article No"].astype(str).str.strip()
desc_df["Article No"] = desc_df["Article No"].astype(str).str.strip()
old_df = old_df.set_index("Article No")
new_df = new_df.set_index("Article No")

# Build description lookup (first found)
desc_lookup = desc_df.drop_duplicates(subset=["Article No"]).set_index("Article No")["Description"].to_dict()

# Combine all unique article numbers as strings
all_articles = sorted(set(old_df.index.astype(str)).union(set(new_df.index.astype(str))))


# Prepare result DataFrame
result = pd.DataFrame(index=all_articles)
result["Old QTY"] = old_df["Sum of QTY"].reindex(all_articles).fillna(0)
result["New QTY"] = new_df["Sum of QTY"].reindex(all_articles).fillna(0)
result["Change"] = result["New QTY"] - result["Old QTY"]

# Mark new/deleted articles
def status(row):
	if row["Old QTY"] == 0 and row["New QTY"] > 0:
		return "New"
	elif row["Old QTY"] > 0 and row["New QTY"] == 0:
		return "Deleted"
	else:
		return ""

result["Status"] = result.apply(status, axis=1)

# Add description column
def get_desc(article_no):
	return desc_lookup.get(article_no, "no description found")

result["Description"] = [get_desc(a) for a in result.index]

# Reset index to have Article No as a column
result = result.reset_index().rename(columns={"index": "Article No"})

# Sort by Article No
result = result.sort_values(by="Article No")

# Save to CSV
output_csv = os.path.join(os.path.dirname(__file__), "article_compare_result.csv")
result.to_csv(output_csv, index=False)

print(f"Comparison saved to {output_csv}")
