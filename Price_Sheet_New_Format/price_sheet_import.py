from __future__ import annotations

import argparse
import collections
import decimal
import difflib
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import load_workbook

NA_VALUE = "N/A"

TARGET_MATERIAL_TYPES = ("ss", "cs", "mapress")
REQUIRED_FILE_CATEGORIES = {
    ("ss", False),
    ("ss", True),
    ("cs", False),
    ("cs", True),
    ("mapress", False),
    ("mapress", True),
}

NORMAL_REQUIRED_KEYS = (
    "system",
    "pipe",
    "type",
    "pipe_component",
    "dn",
    "material",
    "total_weight",
    "qty_pcs",
    "qty_m",
)

ERECTION_REQUIRED_KEYS = (
    "pipe_kks",
    "type",
    "pipe_component",
    "dn",
    "material",
    "qty_pcs_per_m",
)

PAINT_REQUIRED_KEYS = (
    "name_of_system",
    "name_of_pipe",
    "type",
    "description",
    "dn_1",
    "material",
    "quantity_pcs",
    "quantity_m",
    "external_surface_m2",
    "corrosion_class",
    "painting_colour",
    "insulated",
)

SQLITE_NUMERIC_COLUMNS = {
    "UID",
    "Size",
    "Qty pcs",
    "Qty m",
    "Weight kg",
    "Surface m2",
    "Order qty",
    "Weight w spare",
    "Unit price",
    "Total price",
    "Delta qty to previous revision",
}

AGGREGATE_SUM_COLUMNS = ("Qty pcs", "Qty m", "Weight kg", "Surface m2")


@dataclass
class ClassifiedFile:
    path: Path
    material_type: str
    is_erection: bool


def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def tokenize_filename(filename: str) -> List[str]:
    return [token for token in re.split(r"[^a-z0-9]+", filename.lower()) if token]


def classify_file(path: Path) -> Optional[ClassifiedFile]:
    tokens = tokenize_filename(path.name)
    if "material" not in tokens:
        return None

    material_type = None
    for candidate in TARGET_MATERIAL_TYPES:
        if candidate in tokens:
            material_type = candidate
            break

    if material_type is None:
        return None

    return ClassifiedFile(path=path, material_type=material_type, is_erection=("erection" in tokens))


def is_paint_file(path: Path) -> bool:
    tokens = tokenize_filename(path.name)
    return any(token.startswith("paint") for token in tokens)


def locate_required_files(
    input_dir: Path,
) -> Tuple[Dict[Tuple[str, bool], ClassifiedFile], List[str], List[Path], List[Path]]:
    found: Dict[Tuple[str, bool], ClassifiedFile] = {}
    warnings: List[str] = []
    paint_files: List[Path] = []
    ignored_files: List[Path] = []

    for path in sorted(input_dir.glob("*.xls*")):
        if path.name.startswith("~$"):
            continue

        classified = classify_file(path)
        if classified is None:
            if is_paint_file(path):
                paint_files.append(path)
            else:
                ignored_files.append(path)
            continue

        key = (classified.material_type, classified.is_erection)
        if key in found:
            warnings.append(
                f"Duplicate file category for {format_category(key)}: '{found[key].path.name}' and '{path.name}'."
            )
            continue
        found[key] = classified

    missing = sorted(REQUIRED_FILE_CATEGORIES - set(found.keys()))
    for category in missing:
        warnings.append(f"Missing required file for {format_category(category)}.")

    return found, warnings, paint_files, ignored_files


def format_category(category: Tuple[str, bool]) -> str:
    material_type, is_erection = category
    label = material_type.upper() if material_type != "mapress" else "MAPRESS"
    return f"{label} {'erection material' if is_erection else 'material'}"


def material_type_value(material_type: str, is_erection: bool) -> str:
    base = {"ss": "SS", "cs": "CS", "mapress": "Mapress"}[material_type]
    return f"{base} erection" if is_erection else f"{base} shop"


def choose_sheet_with_max_rows(workbook) -> str:
    return max(workbook.sheetnames, key=lambda name: workbook[name].max_row)


def load_sqlite_headers(headers_workbook: Path) -> List[str]:
    wb = load_workbook(headers_workbook, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    headers = [cell.value for cell in ws[1] if cell.value is not None and str(cell.value).strip()]
    return [str(h).strip() for h in headers]


def is_dn_header(normalized: str) -> bool:
    return normalized == "dn"


def aliases_for_key(key: str) -> List[str]:
    alias_map = {
        "system": ["system"],
        "pipe": ["pipe"],
        "pipe_kks": ["pipe kks", "pipe"],
        "type": ["type"],
        "pipe_component": ["pipe component", "pipecomponent"],
        "material": ["material"],
        "total_weight": ["total weight", "total weight kg", "total weight kg"],
        "qty_pcs": ["qty pcs", "qty pcs ", "qty pcs"],
        "qty_m": ["qty m"],
        "qty_pcs_per_m": ["qty pcs m", "qty pcs m", "qty pcs m"],
        "name_of_system": ["name of system"],
        "name_of_pipe": ["name of pipe"],
        "description": ["description"],
        "dn_1": ["dn 1", "dn1"],
        "quantity_pcs": ["quantity pcs", "quantity pcs"],
        "quantity_m": ["quantity m"],
        "external_surface_m2": ["external surface m2", "external surface"],
        "corrosion_class": ["corrosion class"],
        "insulated": ["insulated"],
    }
    return alias_map.get(key, [key])


def smart_find_paint_colour_column(
    normalized_map: Dict[str, int],
    used_indices: set[int],
) -> Optional[int]:
    strong_matches: List[int] = []
    weak_matches: List[int] = []

    for normalized, idx in normalized_map.items():
        if idx in used_indices:
            continue
        has_paint = "paint" in normalized
        has_color = ("colour" in normalized) or ("color" in normalized)
        if has_paint and has_color:
            strong_matches.append(idx)
        elif has_color:
            weak_matches.append(idx)

    if len(strong_matches) == 1:
        return strong_matches[0]
    if len(strong_matches) > 1:
        return None
    if len(weak_matches) == 1:
        return weak_matches[0]
    return None


def best_fuzzy_match(target: str, available: Dict[str, int], used_indices: set[int]) -> Optional[int]:
    candidates: List[Tuple[float, int]] = []
    for normalized, idx in available.items():
        if idx in used_indices:
            continue
        ratio = difflib.SequenceMatcher(None, target, normalized).ratio()
        if ratio >= 0.88:
            candidates.append((ratio, idx))

    if not candidates:
        return None

    candidates.sort(reverse=True)
    if len(candidates) > 1 and abs(candidates[0][0] - candidates[1][0]) < 0.05:
        return None

    return candidates[0][1]


def detect_header_row(ws, max_scan_rows: int = 40) -> int:
    best_row = 1
    best_score = -1

    for row_idx in range(1, min(ws.max_row, max_scan_rows) + 1):
        values = [ws.cell(row=row_idx, column=col).value for col in range(1, ws.max_column + 1)]
        normalized = [normalize_text(v) for v in values if normalize_text(v)]
        if not normalized:
            continue
        # Prefer rows that look like headers but also contain many non-empty cells.
        score = len(set(normalized)) * 10 + len(normalized)
        if score > best_score:
            best_score = score
            best_row = row_idx

    return best_row


def resolve_required_column_indices(
    ws,
    header_row_idx: int,
    is_erection: bool,
) -> Tuple[Dict[str, int], List[str]]:
    required_keys = ERECTION_REQUIRED_KEYS if is_erection else NORMAL_REQUIRED_KEYS
    warnings: List[str] = []

    headers = [ws.cell(row=header_row_idx, column=col).value for col in range(1, ws.max_column + 1)]
    normalized_map: Dict[str, int] = {}

    for idx, value in enumerate(headers, start=1):
        normalized = normalize_text(value)
        if normalized:
            normalized_map[normalized] = idx

    matched: Dict[str, int] = {}
    used_indices: set[int] = set()

    for key in required_keys:
        if key == "dn":
            dn_idx = None
            for normalized, idx in normalized_map.items():
                if is_dn_header(normalized):
                    dn_idx = idx
                    break
            if dn_idx is None:
                raise ValueError("Required column 'DN' not found with exact header 'DN'.")
            matched[key] = dn_idx
            used_indices.add(dn_idx)
            continue

        aliases = aliases_for_key(key)
        exact_idx = None
        for alias in aliases:
            alias_norm = normalize_text(alias)
            if alias_norm in normalized_map and normalized_map[alias_norm] not in used_indices:
                exact_idx = normalized_map[alias_norm]
                break

        if exact_idx is not None:
            matched[key] = exact_idx
            used_indices.add(exact_idx)
            continue

        fuzzy_idx = None
        for alias in aliases:
            alias_norm = normalize_text(alias)
            fuzzy_idx = best_fuzzy_match(alias_norm, normalized_map, used_indices)
            if fuzzy_idx is not None:
                warnings.append(
                    f"Fuzzy matched source column for '{key}' to header '{headers[fuzzy_idx - 1]}'."
                )
                break

        if fuzzy_idx is None:
            raise ValueError(f"Required column '{key}' was not found.")

        matched[key] = fuzzy_idx
        used_indices.add(fuzzy_idx)

    return matched, warnings


def resolve_paint_required_column_indices(
    ws,
    header_row_idx: int,
) -> Tuple[Dict[str, int], List[str]]:
    warnings: List[str] = []
    headers = [ws.cell(row=header_row_idx, column=col).value for col in range(1, ws.max_column + 1)]
    normalized_map: Dict[str, int] = {}

    for idx, value in enumerate(headers, start=1):
        normalized = normalize_text(value)
        if normalized:
            normalized_map[normalized] = idx

    matched: Dict[str, int] = {}
    used_indices: set[int] = set()

    for key in PAINT_REQUIRED_KEYS:
        if key == "dn_1":
            dn_idx = None
            for normalized, idx in normalized_map.items():
                if normalized == "dn 1":
                    dn_idx = idx
                    break
            if dn_idx is None:
                raise ValueError("Required paint column 'DN 1' not found.")
            matched[key] = dn_idx
            used_indices.add(dn_idx)
            continue

        if key == "painting_colour":
            color_idx = smart_find_paint_colour_column(normalized_map, used_indices)
            if color_idx is None:
                raise ValueError("Required paint column for 'Painting colour' not found (color/colour).")
            matched[key] = color_idx
            used_indices.add(color_idx)
            if normalize_text(headers[color_idx - 1]) not in {"painting colour", "painting color"}:
                warnings.append(
                    f"Smart matched paint color column to header '{headers[color_idx - 1]}'."
                )
            continue

        aliases = aliases_for_key(key)
        exact_idx = None
        for alias in aliases:
            alias_norm = normalize_text(alias)
            if alias_norm in normalized_map and normalized_map[alias_norm] not in used_indices:
                exact_idx = normalized_map[alias_norm]
                break

        if exact_idx is not None:
            matched[key] = exact_idx
            used_indices.add(exact_idx)
            continue

        fuzzy_idx = None
        for alias in aliases:
            alias_norm = normalize_text(alias)
            fuzzy_idx = best_fuzzy_match(alias_norm, normalized_map, used_indices)
            if fuzzy_idx is not None:
                warnings.append(
                    f"Fuzzy matched paint source column for '{key}' to header '{headers[fuzzy_idx - 1]}'."
                )
                break

        if fuzzy_idx is None:
            raise ValueError(f"Required paint column '{key}' was not found.")

        matched[key] = fuzzy_idx
        used_indices.add(fuzzy_idx)

    return matched, warnings


def get_cell_text(value: object) -> str:
    if value is None:
        return NA_VALUE
    text = str(value).strip()
    return text if text else NA_VALUE


def _decimal_places_from_excel_number_format(number_format: str) -> Optional[int]:
    if not number_format:
        return None

    # Use only the positive section in formats such as "0.00;[Red]-0.00".
    fmt = number_format.split(";", 1)[0]
    fmt = re.sub(r'"[^"]*"', "", fmt)
    fmt = re.sub(r"\\.", "", fmt)
    fmt = re.sub(r"\[[^\]]*\]", "", fmt)
    fmt = fmt.replace("*", "").replace("_", "")

    if "." not in fmt:
        # Numeric format without decimal separator means integer-like display.
        if re.search(r"[0#?]", fmt):
            return 0
        return None

    fractional = fmt.split(".", 1)[1]
    placeholders = re.findall(r"[0#?]", fractional)
    return len(placeholders)


def cell_to_sql_value(cell: Any) -> Any:
    value = cell.value
    if value is None:
        return NA_VALUE

    if isinstance(value, bool):
        return int(value)

    if isinstance(value, (int, float)):
        dec_value = decimal.Decimal(str(value))
        decimals = _decimal_places_from_excel_number_format(str(getattr(cell, "number_format", "")))

        if decimals is not None:
            quantizer = decimal.Decimal("1").scaleb(-decimals)
            dec_value = dec_value.quantize(quantizer, rounding=decimal.ROUND_HALF_UP)

        if decimals == 0:
            return int(dec_value)

        # Float keeps SQLite numeric storage class while honoring displayed precision.
        return float(dec_value)

    text = str(value).strip()
    return text if text else NA_VALUE


def parse_bom_revision(filename: str) -> str:
    match = re.search(r"_(\d+(?:\.\d+)?)\b", filename)
    if not match:
        return NA_VALUE
    return match.group(1)


def table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    cur = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name = ? LIMIT 1;",
        (table_name,),
    )
    return cur.fetchone() is not None


def db_has_any_rows(conn: sqlite3.Connection) -> bool:
    table_names = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table';")]
    for table_name in table_names:
        cur = conn.execute(f'SELECT 1 FROM "{table_name}" LIMIT 1;')
        if cur.fetchone() is not None:
            return True
    return False


def sanitize_table_name(raw_name: str) -> str:
    candidate = re.sub(r"[^a-zA-Z0-9_]", "_", raw_name)
    if not re.match(r"^[a-zA-Z_]", candidate):
        candidate = f"rev_{candidate}"
    return candidate


def create_table(conn: sqlite3.Connection, table_name: str, headers: Iterable[str]) -> None:
    col_defs = []
    for header in headers:
        sqlite_type = "NUMERIC" if header in SQLITE_NUMERIC_COLUMNS else "TEXT"
        col_defs.append(f'"{header}" {sqlite_type}')
    columns_sql = ", ".join(col_defs)
    conn.execute(f'CREATE TABLE "{table_name}" ({columns_sql});')


def build_rows_for_file(
    classified: ClassifiedFile,
    sqlite_headers: List[str],
    uid_start: int,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    wb = load_workbook(classified.path, data_only=True, read_only=True)
    sheet_name = choose_sheet_with_max_rows(wb)
    ws = wb[sheet_name]

    header_row_idx = detect_header_row(ws)
    source_columns, warnings = resolve_required_column_indices(ws, header_row_idx, classified.is_erection)

    source_to_sqlite = {
        "type": "Type",
        "pipe_component": "Description",
        "dn": "Size",
        "material": "Material",
    }

    if classified.is_erection:
        source_to_sqlite.update(
            {
                "pipe_kks": "KKS",
                "qty_pcs_per_m": "Qty pcs",
            }
        )
    else:
        source_to_sqlite.update(
            {
                "system": "System",
                "pipe": "KKS",
                "total_weight": "Weight kg",
                "qty_pcs": "Qty pcs",
                "qty_m": "Qty m",
            }
        )

    rows: List[Dict[str, Any]] = []
    uid = uid_start

    for row_cells in ws.iter_rows(
        min_row=header_row_idx + 1,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column,
        values_only=False,
    ):
        source_values: Dict[str, Any] = {}
        for source_key, col_idx in source_columns.items():
            source_values[source_key] = cell_to_sql_value(row_cells[col_idx - 1])

        if all(value == NA_VALUE for value in source_values.values()):
            continue

        record: Dict[str, Any] = {header: NA_VALUE for header in sqlite_headers}

        record["UID"] = uid
        uid += 1

        record["Material type"] = material_type_value(classified.material_type, classified.is_erection)
        record["File name - complete"] = classified.path.name
        record["BoM rev from file name"] = parse_bom_revision(classified.path.name)

        for source_key, sqlite_col in source_to_sqlite.items():
            if sqlite_col in record:
                record[sqlite_col] = source_values.get(source_key, NA_VALUE)

        rows.append(record)

    if not rows:
        raise ValueError(
            f"No data rows found in file '{classified.path.name}' after applying required column mapping."
        )

    return rows, warnings


def build_rows_for_paint_file(
    paint_file: Path,
    sqlite_headers: List[str],
    uid_start: int,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    wb = load_workbook(paint_file, data_only=True, read_only=True)
    sheet_name = choose_sheet_with_max_rows(wb)
    ws = wb[sheet_name]

    header_row_idx = detect_header_row(ws)
    source_columns, warnings = resolve_paint_required_column_indices(ws, header_row_idx)

    source_to_sqlite = {
        "name_of_system": "System",
        "name_of_pipe": "KKS",
        "type": "Type",
        "description": "Description",
        "dn_1": "Size",
        "material": "Material",
        "quantity_pcs": "Qty pcs",
        "quantity_m": "Qty m",
        "external_surface_m2": "Surface m2",
        "corrosion_class": "Corrosion category",
        "painting_colour": "Paint colour",
        "insulated": "Insulated",
    }

    rows: List[Dict[str, Any]] = []
    uid = uid_start

    for row_cells in ws.iter_rows(
        min_row=header_row_idx + 1,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column,
        values_only=False,
    ):
        source_values: Dict[str, Any] = {}
        for source_key, col_idx in source_columns.items():
            source_values[source_key] = cell_to_sql_value(row_cells[col_idx - 1])

        if all(value == NA_VALUE for value in source_values.values()):
            continue

        record: Dict[str, Any] = {header: NA_VALUE for header in sqlite_headers}
        record["UID"] = uid
        uid += 1

        record["Material type"] = "Paint shop"
        record["File name - complete"] = paint_file.name
        record["BoM rev from file name"] = parse_bom_revision(paint_file.name)

        for source_key, sqlite_col in source_to_sqlite.items():
            if sqlite_col in record:
                record[sqlite_col] = source_values.get(source_key, NA_VALUE)

        rows.append(record)

    if not rows:
        raise ValueError(
            f"No paint data rows found in file '{paint_file.name}' after applying required column mapping."
        )

    return rows, warnings


def insert_rows(conn: sqlite3.Connection, table_name: str, sqlite_headers: List[str], rows: List[Dict[str, Any]]) -> None:
    placeholders = ", ".join(["?" for _ in sqlite_headers])
    columns = ", ".join([f'"{h}"' for h in sqlite_headers])
    sql = f'INSERT INTO "{table_name}" ({columns}) VALUES ({placeholders});'

    values = [[row.get(header, NA_VALUE) for header in sqlite_headers] for row in rows]
    conn.executemany(sql, values)


def normalize_material_type_for_aggregation(material_type: str) -> str:
    text = material_type.strip().lower()
    if text.endswith("erection"):
        return "Erection"
    if text.startswith("paint"):
        return "Paint shop"
    return material_type


def to_float_or_zero(value: Any) -> float:
    if value in (None, "", NA_VALUE):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        raw = value.strip()
        if not raw or raw.upper() == NA_VALUE:
            return 0.0
        normalized = raw.replace(" ", "")
        if "," in normalized and "." not in normalized:
            normalized = normalized.replace(",", ".")
        try:
            return float(normalized)
        except ValueError:
            return 0.0
    return 0.0


def int_if_whole(value: float) -> Any:
    if abs(value - round(value)) < 1e-12:
        return int(round(value))
    return value


def create_aggregated_table(
    conn: sqlite3.Connection,
    source_table_name: str,
    sqlite_headers: List[str],
) -> str:
    aggregate_table_name = f"{source_table_name}aggreated"

    if table_exists(conn, aggregate_table_name):
        raise ValueError(
            f"Table '{aggregate_table_name}' already exists. Import stopped to protect existing aggregation data."
        )

    cols_sql = ", ".join([f'"{h}"' for h in sqlite_headers])
    rows = conn.execute(f'SELECT {cols_sql} FROM "{source_table_name}";').fetchall()

    grouped: Dict[Tuple[str, str, str, str, str], Dict[str, Any]] = {}
    for db_row in rows:
        record = dict(zip(sqlite_headers, db_row))
        material_type_norm = normalize_material_type_for_aggregation(str(record.get("Material type", NA_VALUE)))

        key = (
            str(record.get("Description", NA_VALUE)),
            str(record.get("Material", NA_VALUE)),
            str(record.get("Type", NA_VALUE)),
            str(record.get("Size", NA_VALUE)),
            material_type_norm,
        )

        if key not in grouped:
            base = {header: NA_VALUE for header in sqlite_headers}
            base["Description"] = key[0]
            base["Material"] = key[1]
            base["Type"] = key[2]
            base["Size"] = key[3]
            base["Material type"] = key[4]
            for col in AGGREGATE_SUM_COLUMNS:
                base[col] = 0.0
            grouped[key] = base

        group_record = grouped[key]
        for col in AGGREGATE_SUM_COLUMNS:
            group_record[col] = to_float_or_zero(group_record.get(col, 0.0)) + to_float_or_zero(record.get(col, 0.0))

    sorted_groups = sorted(
        grouped.values(),
        key=lambda r: (
            str(r.get("Material", "")),
            str(r.get("Type", "")),
            str(r.get("Size", "")),
        ),
    )

    uid = 1
    final_rows: List[Dict[str, Any]] = []
    for row in sorted_groups:
        row["UID"] = uid
        uid += 1
        for col in AGGREGATE_SUM_COLUMNS:
            row[col] = int_if_whole(to_float_or_zero(row[col]))
        final_rows.append(row)

    create_table(conn, aggregate_table_name, sqlite_headers)
    insert_rows(conn, aggregate_table_name, sqlite_headers, final_rows)
    return aggregate_table_name


def run_import(revision_folder: Path, db_path: Path, headers_workbook: Path, input_subdir: str) -> int:
    input_dir = revision_folder / input_subdir
    if not input_dir.exists():
        print(f"ERROR: Input directory not found: {input_dir}")
        return 1

    sqlite_headers = load_sqlite_headers(headers_workbook)
    table_name = sanitize_table_name(revision_folder.name)

    found_files, file_warnings, paint_files, ignored_files = locate_required_files(input_dir)

    if ignored_files:
        for path in ignored_files:
            print(f"INFO: Ignoring non-target file: {path.name}")

    if file_warnings:
        for warning in file_warnings:
            print(f"WARNING: {warning}")
        print("ERROR: Required input file validation failed. Import stopped.")
        return 1

    db_path.parent.mkdir(parents=True, exist_ok=True)
    db_exists_before = db_path.exists()

    with sqlite3.connect(db_path) as conn:
        conn.execute("PRAGMA journal_mode=WAL;")

        if db_exists_before:
            has_data = db_has_any_rows(conn)
            print(f"INFO: Shared database exists: {db_path}")
            print(f"INFO: Existing database has data: {'yes' if has_data else 'no'}")
        else:
            print(f"INFO: Shared database does not exist. It will be created: {db_path}")

        if table_exists(conn, table_name):
            print(f"WARNING: Table '{table_name}' already exists. Import stopped to protect existing revision data.")
            return 1

        all_rows: List[Dict[str, Any]] = []
        uid_counter = 1

        for category in sorted(REQUIRED_FILE_CATEGORIES):
            classified = found_files[category]
            print(f"INFO: Processing {classified.path.name}")
            try:
                rows, mapping_warnings = build_rows_for_file(classified, sqlite_headers, uid_counter)
            except ValueError as exc:
                print(f"ERROR: {exc}")
                print("ERROR: Import stopped due to missing required column/data.")
                return 1

            uid_counter += len(rows)
            all_rows.extend(rows)
            for warning in mapping_warnings:
                print(f"WARNING: {classified.path.name}: {warning}")

        for paint_file in sorted(paint_files):
            print(f"INFO: Processing {paint_file.name}")
            try:
                rows, mapping_warnings = build_rows_for_paint_file(paint_file, sqlite_headers, uid_counter)
            except ValueError as exc:
                print(f"ERROR: {exc}")
                print("ERROR: Import stopped due to missing required column/data.")
                return 1

            uid_counter += len(rows)
            all_rows.extend(rows)
            for warning in mapping_warnings:
                print(f"WARNING: {paint_file.name}: {warning}")

        create_table(conn, table_name, sqlite_headers)
        insert_rows(conn, table_name, sqlite_headers, all_rows)

        try:
            aggregate_table_name = create_aggregated_table(conn, table_name, sqlite_headers)
        except ValueError as exc:
            print(f"WARNING: {exc}")
            return 1

        conn.commit()

        print(f"SUCCESS: Imported {len(all_rows)} rows into table '{table_name}' in {db_path}.")
        agg_count = conn.execute(f'SELECT COUNT(*) FROM "{aggregate_table_name}";').fetchone()[0]
        print(f"SUCCESS: Created aggregated table '{aggregate_table_name}' with {agg_count} rows.")

    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Import Price Sheet revision input material files to SQLite with strict validation.",
    )
    parser.add_argument(
        "--revision-folder",
        type=Path,
        default=Path("Price_Sheet_New_Format/rev0"),
        help="Path to the revision folder (contains input subfolder).",
    )
    parser.add_argument(
        "--input-subdir",
        type=str,
        default="input",
        help="Input subfolder name under revision folder.",
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=Path("Price_Sheet_New_Format/price_sheet.db"),
        help="Path to the shared SQLite database file.",
    )
    parser.add_argument(
        "--headers-workbook",
        type=Path,
        default=Path("Price_Sheet_New_Format/sqlite_table_headers.xlsx"),
        help="Workbook containing the mandatory SQLite table column headers in row 1.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    return run_import(
        revision_folder=args.revision_folder,
        db_path=args.db_path,
        headers_workbook=args.headers_workbook,
        input_subdir=args.input_subdir,
    )


if __name__ == "__main__":
    raise SystemExit(main())
