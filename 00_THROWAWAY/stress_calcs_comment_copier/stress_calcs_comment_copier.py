#!/usr/bin/env python3
"""Copy stress-calcs comments from a source workbook into a target workbook.

Logic:
1) FROM workbook: use column D (identifier) and column S (comment).
2) TO workbook: find matching identifier in column A and write comment into column T.
3) For matched rows, if TO column D is a string with length > 5, propagate that
   same comment to every TO row with the same column D value.

Output file is created next to the TO workbook and named:
processed_YYYYmmdd_HHMMSS.xlsx
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd

# Zero-based column indexes
COL_A = 0
COL_D = 3
COL_S = 18
COL_T = 19


def normalize(value: object) -> str | None:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    return text if text else None


def ensure_min_columns(df: pd.DataFrame, min_col_idx: int) -> pd.DataFrame:
    while df.shape[1] <= min_col_idx:
        df[df.shape[1]] = None
    return df


def main() -> int:
    parser = argparse.ArgumentParser(description="Copy comments between Excel files.")
    parser.add_argument(
        "--from-file",
        dest="from_file",
        default=r"C:\Users\szil\Repos\excel_wizadry\00_THROWAWAY\stress_calcs_comment_copier\extremesupportloads12.02.2026_LS_18.02.2026.xlsx",
        help="Source workbook path (FROM).",
    )
    parser.add_argument(
        "--to-file",
        dest="to_file",
        default=r"C:\Users\szil\Repos\excel_wizadry\00_THROWAWAY\stress_calcs_comment_copier\GP_BQ_11.02.2026_Stress_Calcs_Related.xlsm",
        help="Target workbook path (TO).",
    )
    parser.add_argument(
        "--from-sheet",
        default=0,
        help="FROM sheet name or index (default: 0).",
    )
    parser.add_argument(
        "--to-sheet",
        default=0,
        help="TO sheet name or index (default: 0).",
    )

    args = parser.parse_args()

    from_path = Path(args.from_file)
    to_path = Path(args.to_file)

    if not from_path.exists():
        raise FileNotFoundError(f"FROM file not found: {from_path}")
    if not to_path.exists():
        raise FileNotFoundError(f"TO file not found: {to_path}")

    from_df = pd.read_excel(from_path, sheet_name=args.from_sheet, header=None)
    to_df = pd.read_excel(to_path, sheet_name=args.to_sheet, header=None)

    if from_df.shape[1] <= COL_S:
        raise ValueError("FROM sheet does not contain required column S.")

    to_df = ensure_min_columns(to_df, COL_T)

    from_ids = from_df.iloc[:, COL_D].map(normalize)
    from_comments = from_df.iloc[:, COL_S]

    valid = from_ids.notna()
    id_to_comment: dict[str, object] = {}
    for key, comment in zip(from_ids[valid], from_comments[valid]):
        # Keep latest comment for duplicate identifiers.
        id_to_comment[key] = comment

    to_ids = to_df.iloc[:, COL_A].map(normalize)
    direct_comments = to_ids.map(id_to_comment)

    matched_mask = direct_comments.notna()
    to_df.loc[matched_mask, COL_T] = direct_comments[matched_mask].values

    # Propagation map based on TO column D values from matched rows.
    propagation_map: dict[str, object] = {}
    for idx in to_df.index[matched_mask]:
        d_val = to_df.iat[idx, COL_D]
        d_key = normalize(d_val)
        if d_key is not None and len(d_key) > 5:
            propagation_map[d_key] = to_df.iat[idx, COL_T]

    if propagation_map:
        to_d_norm = to_df.iloc[:, COL_D].map(normalize)
        propagated_comments = to_d_norm.map(propagation_map)
        prop_mask = propagated_comments.notna()
        to_df.loc[prop_mask, COL_T] = propagated_comments[prop_mask].values
    else:
        prop_mask = pd.Series(False, index=to_df.index)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = to_path.with_name(f"processed_{timestamp}.xlsx")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        sheet_name = args.to_sheet if isinstance(args.to_sheet, str) else "Sheet1"
        to_df.to_excel(writer, sheet_name=str(sheet_name), index=False, header=False)

    print(f"Output written: {out_path}")
    print(f"Unique IDs in FROM: {len(id_to_comment)}")
    print(f"Direct matches written to TO:T: {int(matched_mask.sum())}")
    print(f"Rows updated by D-propagation: {int(prop_mask.sum())}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
