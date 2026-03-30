import argparse
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook


REQUIRED_HEADER_TITLES = {"item number", "qty", "description", "decscription"}
UNIT_PATTERN = re.compile(r"^\s*([+-]?\d+(?:[\.,]\d+)?)\s*(kg|mm|m)\s*$", re.IGNORECASE)
NUMBER_PATTERN = re.compile(r"^\s*([+-]?\d+(?:[\.,]\d+)?)\s*$")

MATERIAL_TYPE_RULES: list[dict] = [
    {
        "label": "04 Beam Sections",
        "prefixes": ["TP F", "FLS", "MS"],
        "keywords": ["channel", "rail", "profile", "beam section ms", "beam section tp"],
    },
    {
        "label": "02 Brackets Consoles",
        "prefixes": ["WD F", "WBD", "SA F"],
        "keywords": ["bracket", "console", "base", "pivot joint", "connector", "mounting angle", "beam clip", "beam connection", "holder", "eye", "plate", "end support", "adapter", "joining"],
    },
    {
        "label": "01 Primary Supports",
        "prefixes": ["Stabil", "Ratio"],
        "keywords": ["clamp", "hanger", "pipe clamp", "assembly set", "fixed point", "guided", "sliding", "guiding", "assembly", "axial", "threaded rod"],
    },
    {
        "label": "03 Bolts Screws Nuts",
        "prefixes": ["HZM", "SK", "Formlock"],
        "keywords": ["bolt", "screw", "nut", "washer", "resin anchor"],
    },
    {
        "label": "05 Installation Material",
        "prefixes": [],
        "keywords": ["glass fabric", "end cap"],
    },
]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def sanitize_column_name(name: str, fallback_index: int) -> str:
    cleaned = re.sub(r"\s+", "_", name.strip())
    cleaned = re.sub(r"[^0-9a-zA-Z_]", "", cleaned)
    if not cleaned:
        cleaned = f"COL_{fallback_index}"
    if cleaned[0].isdigit():
        cleaned = f"C_{cleaned}"
    return cleaned


def unique_column_names(raw_headers: list[Any]) -> list[str]:
    names: list[str] = []
    seen: dict[str, int] = {}
    for idx, header in enumerate(raw_headers, start=1):
        base = sanitize_column_name(str(header) if header is not None else "", idx)
        count = seen.get(base, 0)
        seen[base] = count + 1
        final_name = base if count == 0 else f"{base}_{count + 1}"
        names.append(final_name)
    return names


def parse_numeric_text(text: str) -> tuple[bool, Any]:
    m_unit = UNIT_PATTERN.match(text)
    if m_unit:
        num = m_unit.group(1).replace(",", ".")
        if "." in num:
            return True, round(float(num), 2)
        return True, int(num)

    m_num = NUMBER_PATTERN.match(text)
    if m_num:
        num = m_num.group(1).replace(",", ".")
        if "." in num:
            return True, round(float(num), 2)
        return True, int(num)

    return False, text.strip()


def normalize_cell_value(value: Any) -> Any:
    if value is None:
        return None

    if isinstance(value, bool):
        return int(value)

    if isinstance(value, int):
        return value

    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return round(value, 2)

    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        is_num, parsed = parse_numeric_text(stripped)
        if is_num:
            return parsed
        return stripped

    return str(value)


def sql_quote_identifier(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


def infer_sqlite_types(rows: list[list[Any]], col_count: int) -> list[str]:
    inferred: list[str] = []
    for col in range(col_count):
        non_null = [r[col] for r in rows if r[col] is not None]
        if not non_null:
            inferred.append("TEXT")
            continue

        all_int = all(isinstance(v, int) and not isinstance(v, bool) for v in non_null)
        all_num = all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in non_null)
        if all_int:
            inferred.append("INTEGER")
        elif all_num:
            inferred.append("REAL")
        else:
            inferred.append("TEXT")

    return inferred


def classify_material_type(description: str | None) -> str:
    if not description:
        return ""
    desc_lower = description.lower()
    for rule in MATERIAL_TYPE_RULES:
        for prefix in rule["prefixes"]:
            if desc_lower.startswith(prefix.lower()):
                return rule["label"]
    for rule in MATERIAL_TYPE_RULES:
        for keyword in rule["keywords"]:
            if keyword in desc_lower:
                return rule["label"]
    return ""


def _find_header_row_in_df(df: "pd.DataFrame") -> int:
    for i, row in df.iterrows():
        normalized = [normalize_text(v) for v in row]
        if (
            "item number" in normalized
            and "qty" in normalized
            and ("description" in normalized or "decscription" in normalized)
        ):
            return int(i)
    raise ValueError(
        "Could not find header row with Item Number, Qty and Description/Decscription in sheet 'Total Qty'."
    )


def import_total_qty_to_sqlite(
    excel_path: Path, revision: str = "REV0", db_path: Path | None = None
) -> Path:
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    raw_df = pd.read_excel(
        excel_path,
        sheet_name="Total Qty",
        header=None,
        engine="calamine",
    )

    header_row_idx = _find_header_row_in_df(raw_df)
    raw_headers: list[Any] = raw_df.iloc[header_row_idx].tolist()
    column_names = unique_column_names(raw_headers)

    # skip the header row itself and any immediately following blank row
    data_df = raw_df.iloc[header_row_idx + 1 :].reset_index(drop=True)
    # drop rows that are entirely NaN/None
    data_df = data_df.dropna(how="all")

    rows: list[list[Any]] = []
    for _, row in data_df.iterrows():
        cleaned = [normalize_cell_value(None if (v is None or (isinstance(v, float) and pd.isna(v))) else v) for v in row]
        if any(v is not None and v != "" for v in cleaned):
            rows.append(cleaned)

    first_data_row = header_row_idx + 2  # 1-based for display
    last_useful_row = first_data_row + len(rows) - 1

    if db_path is None:
        db_path = excel_path.with_suffix(".sqlite")
    raw_table = f"RAW_DATA_{revision}"
    aggregated_table = f"AGGREGATED_DATA_{revision}"

    types = infer_sqlite_types(rows, len(column_names))

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()

        cur.execute(f"DROP TABLE IF EXISTS {sql_quote_identifier(raw_table)}")

        column_defs = ", ".join(
            f"{sql_quote_identifier(col)} {col_type}"
            for col, col_type in zip(column_names, types, strict=True)
        )
        cur.execute(f"CREATE TABLE {sql_quote_identifier(raw_table)} ({column_defs})")

        placeholders = ", ".join(["?"] * len(column_names))
        insert_sql = f"INSERT INTO {sql_quote_identifier(raw_table)} VALUES ({placeholders})"
        cur.executemany(insert_sql, rows)

        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name = ?",
            (aggregated_table,),
        )
        aggregated_exists = cur.fetchone() is not None
        if not aggregated_exists:
            cur.execute(
                f"CREATE TABLE {sql_quote_identifier(aggregated_table)} AS "
                f"SELECT * FROM {sql_quote_identifier(raw_table)} WHERE 1=0"
            )

        conn.commit()
    finally:
        conn.close()

    print(f"Header row index: {header_row_idx}")
    print(f"Data rows: {first_data_row} to {last_useful_row}")
    print(f"Imported {len(rows)} rows from 'Total Qty' into {raw_table}")
    if aggregated_exists:
        print(f"Table {aggregated_table} already existed and was kept")
    else:
        print(f"Created empty table {aggregated_table}")
    print(f"Database created: {db_path}")
    return db_path


def populate_aggregated_table(db_path: Path, revision: str = "REV0") -> int:
    if not db_path.exists():
        raise FileNotFoundError(f"DB file not found: {db_path}")

    raw_table = f"RAW_DATA_{revision}"
    agg_table = f"AGGREGATED_DATA_{revision}"

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (raw_table,))
        if cur.fetchone() is None:
            raise ValueError(f"Table {raw_table} does not exist.")

        cur.execute(
            f"SELECT Item_Number, Qty, Description, Cut_Length_mm, Total_Weight_kg, Remarks, Coating "
            f"FROM {sql_quote_identifier(raw_table)}"
        )
        raw_rows = cur.fetchall()

        individual: list[dict] = []
        groups: dict[tuple, list[dict]] = {}
        bracket_groups: dict[tuple, list[dict]] = {}
        _bracket_x_re = re.compile(r" x \d{3,4}(.*)")

        for item_number, qty, description, cut_length_mm, total_weight_kg, remarks, coating in raw_rows:
            material_type = classify_material_type(description)
            row_dict = {
                "Article_Number": item_number,
                "Qty": qty,
                "Description": description,
                "Cut_Length_mm": cut_length_mm,
                "Total_Weight_kg": total_weight_kg,
                "Remarks": remarks,
                "Coating": coating,
                "Material_Type": material_type,
            }
            # 02 Brackets Consoles: special handling
            if material_type == "02 Brackets Consoles":
                # Step 1: strip " x NNN[suffix]" from description if present
                cleaned_desc = description or ""
                if description and " x " in description:
                    cleaned_desc = _bracket_x_re.sub(
                        lambda m: (" " + m.group(1).strip()) if m.group(1).strip() else "",
                        description,
                    ).strip()
                    row_dict["Description"] = cleaned_desc
                # Step 2: angled bracket → keep cut_length, stay as individual
                if description and "angled" in description.lower():
                    individual.append(row_dict)
                else:
                    # clear cut_length, aggregate on (article_number, cleaned_desc, coating)
                    row_dict["Cut_Length_mm"] = None
                    bracket_groups.setdefault((item_number, coating, cleaned_desc), []).append(row_dict)
            elif cut_length_mm is not None:
                individual.append(row_dict)
            else:
                groups.setdefault((item_number, coating), []).append(row_dict)

        aggregated: list[dict] = []
        for (item_number, coating), group in groups.items():
            total_qty = sum(r["Qty"] or 0 for r in group)
            non_null_w = [r["Total_Weight_kg"] for r in group if r["Total_Weight_kg"] is not None]
            total_weight = round(sum(non_null_w), 2) if non_null_w else None
            description = group[0]["Description"]
            seen: set[str] = set()
            parts: list[str] = []
            for r in group:
                rem = r["Remarks"]
                if rem and rem not in seen:
                    seen.add(rem)
                    parts.append(rem)
            agg_remarks = " | ".join(parts) if parts else None
            aggregated.append({
                "Article_Number": item_number,
                "Qty": total_qty,
                "Description": description,
                "Cut_Length_mm": None,
                "Total_Weight_kg": total_weight,
                "Remarks": agg_remarks,
                "Coating": coating,
                "Material_Type": group[0]["Material_Type"],
            })

        for (item_number, coating, cleaned_desc), group in bracket_groups.items():
            total_qty = sum(r["Qty"] or 0 for r in group)
            non_null_w = [r["Total_Weight_kg"] for r in group if r["Total_Weight_kg"] is not None]
            total_weight = round(sum(non_null_w), 2) if non_null_w else None
            aggregated.append({
                "Article_Number": item_number,
                "Qty": total_qty,
                "Description": cleaned_desc,
                "Cut_Length_mm": None,
                "Total_Weight_kg": total_weight,
                "Remarks": None,
                "Coating": coating,
                "Material_Type": "02 Brackets Consoles",
            })

        output_rows: list[tuple] = []
        for row_dict in aggregated + individual:
            material_type = row_dict.get("Material_Type") or classify_material_type(row_dict["Description"])
            output_rows.append((
                material_type,
                row_dict["Article_Number"],
                row_dict["Description"],
                row_dict["Qty"],
                row_dict["Cut_Length_mm"],
                row_dict["Total_Weight_kg"],
                row_dict["Remarks"],
                row_dict["Coating"],
            ))

        output_rows.sort(key=lambda r: (r[0] or "", (r[2] or "").lower()))

        cur.execute(f"DROP TABLE IF EXISTS {sql_quote_identifier(agg_table)}")
        cur.execute(
            f"CREATE TABLE {sql_quote_identifier(agg_table)} ("
            f'"Material_Type" TEXT, '
            f'"Article_Number" INTEGER, '
            f'"Description" TEXT, '
            f'"Qty" INTEGER, '
            f'"Cut_Length_mm" REAL, '
            f'"Total_Weight_kg" REAL, '
            f'"Remarks" TEXT, '
            f'"Coating" TEXT'
            f")"
        )
        cur.executemany(
            f"INSERT INTO {sql_quote_identifier(agg_table)} VALUES (?,?,?,?,?,?,?,?)",
            output_rows,
        )
        conn.commit()

        print(
            f"Populated {agg_table}: {len(output_rows)} rows "
            f"({len(aggregated)} aggregated [{len(bracket_groups)} bracket groups], {len(individual)} individual with cut length)"
        )
        return len(output_rows)
    finally:
        conn.close()


def export_aggregated_to_excel(db_path: Path, out_xlsx: Path | None = None, revision: str = "REV0") -> Path:
    if not db_path.exists():
        raise FileNotFoundError(f"DB file not found: {db_path}")

    table = f"AGGREGATED_DATA_{revision}"
    if out_xlsx is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_xlsx = db_path.with_name(f"{db_path.stem}_{table}_{timestamp}.xlsx")

    raw_table = f"RAW_DATA_{revision}"

    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name = ?",
            (table,),
        )
        if cur.fetchone() is None:
            raise ValueError(f"Table {table} does not exist in {db_path}")

        cur.execute(f"SELECT * FROM {sql_quote_identifier(table)}")
        agg_rows = cur.fetchall()
        cur.execute(f"PRAGMA table_info({sql_quote_identifier(table)})")
        agg_columns = [r[1] for r in cur.fetchall()]

        cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name = ?",
            (raw_table,),
        )
        has_raw = cur.fetchone() is not None
        if has_raw:
            cur.execute(f"SELECT * FROM {sql_quote_identifier(raw_table)}")
            raw_rows = cur.fetchall()
            cur.execute(f"PRAGMA table_info({sql_quote_identifier(raw_table)})")
            raw_columns = [r[1] for r in cur.fetchall()]
    finally:
        conn.close()

    def _write_sheet(ws, columns, rows):
        ws.append(columns)
        for row in rows:
            ws.append(list(row))
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if isinstance(cell.value, float):
                    cell.number_format = "0.00"
        # Freeze top row
        ws.freeze_panes = "A2"
        # AutoFilter on header row
        ws.auto_filter.ref = ws.dimensions
        # Auto-fit column widths
        for col_cells in ws.columns:
            max_len = max(
                (len(str(cell.value)) if cell.value is not None else 0) for cell in col_cells
            )
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)

    wb = Workbook()

    # Raw sheet first
    if has_raw:
        ws_raw = wb.active
        ws_raw.title = raw_table
        _write_sheet(ws_raw, raw_columns, raw_rows)
        ws_agg = wb.create_sheet(title=table)
    else:
        ws_agg = wb.active
        ws_agg.title = table

    _write_sheet(ws_agg, agg_columns, agg_rows)

    wb.save(out_xlsx)
    print(f"Exported {len(agg_rows)} aggregated + {len(raw_rows) if has_raw else 0} raw rows to: {out_xlsx}")
    return out_xlsx


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Import Total Qty sheet to SQLite and export aggregated table to Excel.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    import_parser = subparsers.add_parser("import", help="Import 'Total Qty' from Excel into SQLite.")
    import_parser.add_argument("--excel", required=True, help="Path to source Excel file (.xlsm/.xlsx)")
    import_parser.add_argument("--revision", default="REV0", help="Revision tag, e.g. REV0, REV1")

    export_parser = subparsers.add_parser("export", help="Export AGGREGATED table from SQLite to Excel.")
    export_parser.add_argument("--db", required=True, help="Path to SQLite database")
    export_parser.add_argument("--revision", default="REV0", help="Revision tag, e.g. REV0, REV1")
    export_parser.add_argument("--out", help="Output xlsx path (optional)")

    populate_parser = subparsers.add_parser("populate", help="Populate AGGREGATED table from RAW and export to Excel.")
    populate_parser.add_argument("--db", required=True, help="Path to SQLite database")
    populate_parser.add_argument("--revision", default="REV0", help="Revision tag, e.g. REV0, REV1")
    populate_parser.add_argument("--out", help="Output xlsx path (optional)")

    test_parser = subparsers.add_parser("test-run", help="Import + populate + export into _test_runs/ with timestamp.")
    test_parser.add_argument("--excel", required=True, help="Path to source Excel file (.xlsm/.xlsx)")
    test_parser.add_argument("--revision", default="REV0", help="Revision tag, e.g. REV0, REV1")

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.command == "import":
        import_total_qty_to_sqlite(Path(args.excel), revision=args.revision)
        return

    if args.command == "export":
        out = Path(args.out) if args.out else None
        export_aggregated_to_excel(Path(args.db), out_xlsx=out, revision=args.revision)
        return

    if args.command == "populate":
        populate_aggregated_table(Path(args.db), revision=args.revision)
        out = Path(args.out) if args.out else None
        export_aggregated_to_excel(Path(args.db), out_xlsx=out, revision=args.revision)
        return

    if args.command == "test-run":
        excel_path = Path(args.excel)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        test_dir = excel_path.parent / "_test_runs"
        test_dir.mkdir(exist_ok=True)
        db_path = test_dir / f"{timestamp}.sqlite"
        out_xlsx = test_dir / f"{timestamp}_AGGREGATED_DATA_{args.revision}.xlsx"
        print(f"--- Test run: {timestamp} ---")
        import_total_qty_to_sqlite(excel_path, revision=args.revision, db_path=db_path)
        populate_aggregated_table(db_path, revision=args.revision)
        export_aggregated_to_excel(db_path, out_xlsx=out_xlsx, revision=args.revision)
        return

    raise ValueError(f"Unsupported command: {args.command}")


if __name__ == "__main__":
    main()