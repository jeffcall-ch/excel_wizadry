from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd

DEFAULT_STATUS_FILE = (
    "AB010-KVI-50299274_1.0 - Pipe support status list - "
    "Lot General Piping Detail Engineering.xlsm"
)
DEFAULT_BATCHED_FILE = (
    "CA100-KVI-50296373_2.0 - BoM General Hangers and support material "
    "LS_erection_split_batched.xlsx"
)
DEFAULT_C5_FILE = "C5_KKS.xlsx"
DEFAULT_STATUS_SHEET = "Pipe support status list"
DEFAULT_BATCHED_SHEET = "Detail"
DEFAULT_OUTPUT_FILE = "KKS_not_in_batch_report.xlsx"


def normalize_label(value: object) -> str:
    text = str(value).replace("\n", " ").strip().lower()
    return " ".join(text.split())


def normalize_kks(value: object) -> str | None:
    if pd.isna(value):
        return None

    text = str(value).strip().upper().replace(" ", "")
    if not text or text == "NAN":
        return None

    # Status list uses values like 0CFX.../SU while other files use /0CFX.../SU.
    if "/SU" in text and not text.startswith("/"):
        text = f"/{text}"

    return text


def natural_sort_token(value: object) -> str:
    text = "" if pd.isna(value) else str(value).strip().upper()
    return re.sub(r"\d+", lambda m: f"{int(m.group()):06d}", text)


def join_unique(values: pd.Series) -> str:
    seen: set[str] = set()
    ordered: list[str] = []

    for value in values:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if not text or text.lower() == "nan":
            continue
        if text not in seen:
            seen.add(text)
            ordered.append(text)

    return " | ".join(ordered)


def find_column(columns: list[str], required_terms: list[str]) -> str | None:
    for col in columns:
        label = normalize_label(col)
        if all(term in label for term in required_terms):
            return col
    return None


def load_status(status_file: Path, sheet_name: str) -> tuple[pd.DataFrame, str]:
    status_df = pd.read_excel(status_file, sheet_name=sheet_name, engine="openpyxl")
    status_columns = list(status_df.columns)

    kks_col = find_column(status_columns, ["support", "drawing", "number"])
    if not kks_col:
        raise ValueError("Could not find status KKS column (Support Drawing Number).")

    status_col = find_column(status_columns, ["released", "deleted"])
    system_col = find_column(status_columns, ["system"])
    wa_col = find_column(status_columns, ["working", "area"])
    primary_col = find_column(status_columns, ["primary", "supports"])
    aic_col = find_column(status_columns, ["aic"])
    note_col = find_column(status_columns, ["note"])
    rev_cols = [
        col for col in status_columns if normalize_label(col).startswith("rev.")
    ]

    status_df = status_df.copy()
    status_df["KKS /SU"] = status_df[kks_col].map(normalize_kks)
    status_df = status_df[status_df["KKS /SU"].notna()].copy()
    status_df = status_df[status_df["KKS /SU"].str.contains("/SU", na=False)].copy()

    context_map: list[tuple[str, str]] = []
    if status_col:
        context_map.append(("Status", status_col))
    if note_col:
        context_map.append(("Note", note_col))
    if system_col:
        context_map.append(("System", system_col))
    if wa_col:
        context_map.append(("Working area", wa_col))
    if primary_col:
        context_map.append(("List of primary supports", primary_col))
    if aic_col:
        context_map.append(("AIC", aic_col))
    for col in sorted(rev_cols, key=normalize_label):
        context_map.append((str(col).strip(), col))

    agg_map = {src: join_unique for _, src in context_map}
    status_unique = status_df.groupby("KKS /SU", as_index=False).agg(agg_map)

    rename_map = {src: dst for dst, src in context_map}
    status_unique = status_unique.rename(columns=rename_map)

    if "Status" not in status_unique.columns:
        status_unique["Status"] = ""
    if "Note" not in status_unique.columns:
        status_unique["Note"] = ""
    if "Working area" not in status_unique.columns:
        status_unique["Working area"] = ""
    if "System" not in status_unique.columns:
        status_unique["System"] = ""

    status_deleted = status_unique["Status"].str.contains("deleted", case=False, na=False)
    note_deleted = status_unique["Note"].str.contains(
        "deleted|withdraw", case=False, na=False, regex=True
    )
    status_unique["Deleted in status"] = status_deleted | note_deleted

    status_unique = status_unique.sort_values("KKS /SU", key=lambda s: s.map(natural_sort_token))
    status_unique = status_unique.reset_index(drop=True)

    return status_unique, kks_col


def load_batched_set(batched_file: Path, sheet_name: str) -> set[str]:
    batched_df = pd.read_excel(batched_file, sheet_name=sheet_name, engine="openpyxl")

    if "KKS code" in batched_df.columns:
        kks_col = "KKS code"
    else:
        candidates = [
            col
            for col in batched_df.columns
            if "kks" in normalize_label(col) or "support" in normalize_label(col)
        ]
        if not candidates:
            raise ValueError("Could not find KKS column in batched file.")
        kks_col = candidates[0]

    kks = {
        k
        for k in (normalize_kks(value) for value in batched_df[kks_col])
        if k and "/SU" in k
    }
    return kks


def load_c5_set(c5_file: Path) -> set[str]:
    c5_df = pd.read_excel(c5_file, sheet_name=0, engine="openpyxl")
    c5_columns = list(c5_df.columns)

    candidates = [
        col
        for col in c5_columns
        if "kks" in normalize_label(col) or "support" in normalize_label(col)
    ]
    c5_col = candidates[0] if candidates else c5_columns[0]

    kks = {
        k for k in (normalize_kks(value) for value in c5_df[c5_col]) if k and "/SU" in k
    }
    return kks


def build_report_tables(
    status_unique: pd.DataFrame,
    batched_set: set[str],
    c5_set: set[str],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    status_set = set(status_unique["KKS /SU"])

    raw_missing_set = status_set - batched_set
    raw_missing = status_unique[status_unique["KKS /SU"].isin(raw_missing_set)].copy()
    raw_missing["In C5 list"] = raw_missing["KKS /SU"].isin(c5_set)

    deducted_c5 = raw_missing[raw_missing["In C5 list"]].copy()
    deducted_c5 = deducted_c5.sort_values(
        ["Working area", "KKS /SU"],
        key=lambda s: s.map(natural_sort_token),
    ).reset_index(drop=True)

    final_missing = raw_missing[~raw_missing["In C5 list"]].copy()
    final_missing["Final tag"] = final_missing["Deleted in status"].map(
        lambda x: "Deleted (excluded from active follow-up)" if x else "Active missing"
    )

    active_missing = final_missing[~final_missing["Deleted in status"]].copy()
    deleted_missing = final_missing[final_missing["Deleted in status"]].copy()

    active_missing = active_missing.sort_values(
        ["Working area", "KKS /SU"],
        key=lambda s: s.map(natural_sort_token),
    )
    deleted_missing = deleted_missing.sort_values(
        ["Working area", "KKS /SU"],
        key=lambda s: s.map(natural_sort_token),
    )

    final_with_deleted_bottom = pd.concat([active_missing, deleted_missing], ignore_index=True)

    summary_rows = [
        ("Unique /SU in status file", len(status_set)),
        ("Unique /SU in batched file", len(batched_set)),
        ("Raw missing /SU (status - batched)", len(raw_missing_set)),
        ("Raw missing that are C5", len(deducted_c5)),
        ("Final list after C5 deduction", len(final_with_deleted_bottom)),
        (
            "Final list ACTIVE (excluding deleted)",
            int((~final_with_deleted_bottom["Deleted in status"]).sum()),
        ),
        (
            "Final list DELETED (shown at bottom)",
            int(final_with_deleted_bottom["Deleted in status"].sum()),
        ),
    ]
    summary = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

    return summary, raw_missing, deducted_c5, final_with_deleted_bottom


def write_report(
    output_file: Path,
    summary: pd.DataFrame,
    raw_missing: pd.DataFrame,
    deducted_c5: pd.DataFrame,
    final_missing: pd.DataFrame,
) -> None:
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        final_missing.to_excel(writer, sheet_name="Missing_Final", index=False)
        raw_missing.to_excel(writer, sheet_name="Missing_Raw", index=False)
        deducted_c5.to_excel(writer, sheet_name="Deducted_C5", index=False)


def parse_args() -> argparse.Namespace:
    script_dir = Path(__file__).resolve().parent

    parser = argparse.ArgumentParser(
        description=(
            "Find /SU KKS codes that exist in status file but are missing in "
            "batched file, then deduct C5 list and place deleted at bottom."
        )
    )
    parser.add_argument(
        "--status-file",
        type=Path,
        default=script_dir / DEFAULT_STATUS_FILE,
        help="Status workbook path.",
    )
    parser.add_argument(
        "--status-sheet",
        default=DEFAULT_STATUS_SHEET,
        help="Status sheet name.",
    )
    parser.add_argument(
        "--batched-file",
        type=Path,
        default=script_dir / DEFAULT_BATCHED_FILE,
        help="Batched workbook path.",
    )
    parser.add_argument(
        "--batched-sheet",
        default=DEFAULT_BATCHED_SHEET,
        help="Batched sheet name.",
    )
    parser.add_argument(
        "--c5-file",
        type=Path,
        default=script_dir / DEFAULT_C5_FILE,
        help="C5 workbook path.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=script_dir / DEFAULT_OUTPUT_FILE,
        help="Output report xlsx path.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    status_file = args.status_file.resolve()
    batched_file = args.batched_file.resolve()
    c5_file = args.c5_file.resolve()
    output_file = args.output.resolve()

    for file_path in [status_file, batched_file, c5_file]:
        if not file_path.exists():
            raise FileNotFoundError(f"Input file not found: {file_path}")

    status_unique, status_kks_col = load_status(
        status_file=status_file,
        sheet_name=args.status_sheet,
    )
    batched_set = load_batched_set(
        batched_file=batched_file,
        sheet_name=args.batched_sheet,
    )
    c5_set = load_c5_set(c5_file=c5_file)

    summary, raw_missing, deducted_c5, final_missing = build_report_tables(
        status_unique=status_unique,
        batched_set=batched_set,
        c5_set=c5_set,
    )

    write_report(
        output_file=output_file,
        summary=summary,
        raw_missing=raw_missing,
        deducted_c5=deducted_c5,
        final_missing=final_missing,
    )

    print(f"Status file: {status_file}")
    print(f"Status KKS column used: {status_kks_col}")
    print(f"Batched file: {batched_file}")
    print(f"C5 file: {c5_file}")
    print(f"Output report: {output_file}")
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
