from __future__ import annotations

import argparse
import logging
import math
import re
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Tuple

from Price_Sheet_New_Format.logging_utils import DEFAULT_LOG_DIR, setup_csv_logging
from Price_Sheet_New_Format.settings_markdown import load_markdown_settings, sections_hint

NA_VALUE = "N/A"
logger = logging.getLogger(__name__)

TUBI_CS_SHOP_PRE_PAINT_COL = "Shop order calculated len"
TUBI_CS_SHOP_POST_PAINT_COL = "With paint update order length"
SURFACE_W_SPARE_COL = "Surface m2 w spare"
PAINT_BASIC_BUCKET = "basic only"
PAINT_BASIC_OR_PRIMER_NAMES = {"basic only"}


def parse_float(value: Any, default: float = 0.0) -> float:
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or text.upper() == NA_VALUE:
        return default
    text = text.replace(" ", "")
    if "," in text and "." not in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return default


def parse_int(value: Any, default: int = 0) -> int:
    return int(round(parse_float(value, float(default))))


def parse_keywords(value: Any) -> List[str]:
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    return [v.strip().lower() for v in text.split(",") if v.strip()]


def contains_any(text: str, keywords: List[str]) -> bool:
    text_l = text.lower()
    return any(k in text_l for k in keywords)


def contains_all(text: str, keywords: List[str]) -> bool:
    text_l = text.lower()
    return all(k in text_l for k in keywords)


def extract_numeric(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value)
    m = re.search(r"(\d+(?:[\.,]\d+)?)", text)
    if not m:
        return 0.0
    return parse_float(m.group(1), 0.0)


def ceil_int(value: float) -> int:
    return int(math.ceil(value))


def round_up_to_multiple(value: float, multiple: int) -> int:
    if multiple <= 0:
        return ceil_int(value)
    return int(math.ceil(value / multiple) * multiple)


def round_up_to_next_10(value: float) -> int:
    return round_up_to_multiple(value, 10)


def round_up_nut_washer(value: float) -> int:
    if value <= 100:
        return round_up_to_multiple(value, 10)
    if value <= 500:
        return round_up_to_multiple(value, 50)
    return round_up_to_multiple(value, 100)


def uid_category(material_type: str) -> Tuple[int, str]:
    mt = str(material_type or "").strip().lower()
    if "mapress shop" in mt:
        return (1, "MA")
    if "ss shop" in mt:
        return (2, "SS")
    if "cs shop" in mt:
        return (3, "CS")
    if "04 erection" in mt or ("erection" in mt and "paint" not in mt):
        return (4, "ER")
    if "paint shop" in mt:
        return (5, "PAS")
    if "paint erection" in mt:
        return (6, "PAE")
    if "blind disks" in mt:
        return (7, "BD")
    if "orifice plates" in mt:
        return (8, "OP")
    if "flange guards" in mt:
        return (9, "FG")
    return (99, "OT")


def rewrite_aggregated_uids(conn: sqlite3.Connection, agg_table: str, mt_col: str) -> None:
    cols = conn.execute(f'PRAGMA table_info("{agg_table}");').fetchall()
    existing = {str(c[1]) for c in cols}
    if "UID" not in existing:
        return

    rows = conn.execute(
        f'SELECT rowid, "{mt_col}", "Material", "Type", "Size" FROM "{agg_table}"'
    ).fetchall()

    sortable: List[Tuple[int, str, str, int, str, float, str, int]] = []
    for rowid, mt, material, type_value, size in rows:
        cat_num, code = uid_category(str(mt or ""))
        sortable.append(
            (
                cat_num,
                code,
                str(type_value or "").strip().lower(),
                0 if str(material or "").strip().upper() != NA_VALUE else 1,
                str(material or "").strip().lower(),
                extract_numeric(size),
                str(size or "").strip().lower(),
                int(rowid),
            )
        )

    sortable.sort(key=lambda t: (t[0], t[2], t[3], t[4], t[5], t[6], t[7]))

    counters: Dict[Tuple[int, str], int] = {}
    updates: List[Tuple[str, int]] = []
    for cat_num, code, _, _, _, _, _, rowid in sortable:
        key = (cat_num, code)
        next_index = counters.get(key, 0) + 1
        counters[key] = next_index
        uid_text = f"{cat_num:02d}-{code}-{next_index:03d}"
        updates.append((uid_text, rowid))

    if updates:
        conn.executemany(f'UPDATE "{agg_table}" SET "UID"=? WHERE rowid=?', updates)


def load_spare_settings(settings_workbook: Path) -> Dict[str, Any]:
    settings_doc = load_markdown_settings(settings_workbook)

    general_records = settings_doc.get_records("SpareGeneral")
    if not general_records:
        raise ValueError(f"Missing settings table. {sections_hint('SpareGeneral')}")
    general = {str(r.get("Key", "")).strip(): r.get("Value") for r in general_records}

    main_records = settings_doc.get_records("SpareMainRules")
    if not main_records:
        raise ValueError(f"Missing settings table. {sections_hint('SpareMainRules')}")
    main_map = {str(r.get("Key", "")).strip(): r.get("Value") for r in main_records}

    erection_records = settings_doc.get_records("SpareErectionRules")
    if not erection_records:
        raise ValueError(f"Missing settings table. {sections_hint('SpareErectionRules')}")
    erection_map = {str(r.get("Key", "")).strip(): r.get("Value") for r in erection_records}

    keyword_records = settings_doc.get_records("SpareKeywords")
    if not keyword_records:
        raise ValueError(f"Missing settings table. {sections_hint('SpareKeywords')}")

    kw: Dict[str, List[str]] = {}
    for rec in keyword_records:
        name = str(rec.get("Name", "")).strip()
        if name:
            kw[name] = parse_keywords(rec.get("Keywords"))

    blind_disk_records: List[Dict[str, object]] = []
    try:
        blind_disk_records = settings_doc.get_records("BlindDiskThickness")
    except ValueError:
        blind_disk_records = []

    blind_disk_thickness: Dict[str, str] = {}
    for rec in blind_disk_records:
        dn_raw = str(rec.get("DN", "")).strip()
        s1_raw = str(rec.get("S1 mm", "")).strip()
        if dn_raw and s1_raw:
            blind_disk_thickness[dn_raw] = s1_raw

    paint_erection_defaults_records = settings_doc.get_records("PaintErectionDefaults")
    if not paint_erection_defaults_records:
        raise ValueError(f"Missing settings table. {sections_hint('PaintErectionDefaults')}")
    paint_erection_defaults = records_to_column_map(paint_erection_defaults_records)

    blind_disk_defaults_records = settings_doc.get_records("BlindDiskDefaults")
    if not blind_disk_defaults_records:
        raise ValueError(f"Missing settings table. {sections_hint('BlindDiskDefaults')}")
    blind_disk_defaults = records_to_column_map(blind_disk_defaults_records)

    flange_guard_system_records = settings_doc.get_records("FlangeGuardSystems")
    if not flange_guard_system_records:
        raise ValueError(f"Missing settings table. {sections_hint('FlangeGuardSystems')}")
    flange_guard_system_keywords = [
        str(rec.get("SystemKeyword", "")).strip().lower()
        for rec in flange_guard_system_records
        if str(rec.get("SystemKeyword", "")).strip()
    ]

    flange_guard_defaults_records = settings_doc.get_records("FlangeGuardDefaults")
    if not flange_guard_defaults_records:
        raise ValueError(f"Missing settings table. {sections_hint('FlangeGuardDefaults')}")
    flange_guard_defaults = records_to_column_map(flange_guard_defaults_records)

    return {
        "aggregate_table_suffix": str(general.get("aggregate_table_suffix", "aggreated") or "aggreated"),
        "spare_column": str(general.get("spare_column", "Spare") or "Spare"),
        "spare_info_column": str(general.get("spare_info_column", "Spare info") or "Spare info"),
        "material_type_column": str(general.get("material_type_column", "Material type") or "Material type"),
        "description_column": str(general.get("description_column", "Description") or "Description"),
        "type_label_column": str(general.get("type_label_column", "Type label") or "Type label"),
        "type_column": str(general.get("type_column", "Type") or "Type"),
        "size_column": str(general.get("size_column", "Size") or "Size"),
        "qty_pcs_column": str(general.get("qty_pcs_column", "Qty pcs") or "Qty pcs"),
        "qty_m_column": str(general.get("qty_m_column", "Qty m") or "Qty m"),
        "weight_column": str(general.get("weight_column", "Weight kg") or "Weight kg"),
        "total_qty_column": str(general.get("total_qty_column", "Order qty") or "Order qty"),
        "weight_w_spare_column": str(general.get("weight_w_spare_column", "Weight w spare") or "Weight w spare"),
        "paint_prefix": str(general.get("paint_prefix", "paint") or "paint").lower(),
        "erection_exact": str(general.get("erection_exact", "Erection") or "Erection"),
        "shop_suffix": str(general.get("shop_suffix", "shop") or "shop").lower(),
        "main": {
            "seal_qty_threshold": parse_int(main_map.get("seal_qty_threshold", 10)),
            "seal_min_spare_low": parse_int(main_map.get("seal_min_spare_low", 5)),
            "seal_min_spare_high": parse_int(main_map.get("seal_min_spare_high", 5)),
            "seal_percent": parse_float(main_map.get("seal_percent", 0.10)),
            "critical_qty_threshold": parse_int(main_map.get("critical_qty_threshold", 10)),
            "critical_min_spare": parse_int(main_map.get("critical_min_spare", 3)),
            "critical_percent_low": parse_float(main_map.get("critical_percent_low", 0.50)),
            "critical_percent_high": parse_float(main_map.get("critical_percent_high", 0.10)),
            "threshold_1": parse_int(main_map.get("threshold_1", 5)),
            "threshold_2": parse_int(main_map.get("threshold_2", 20)),
            "size_threshold": parse_float(main_map.get("size_threshold", 50.0)),
            "percent_pcs_small": parse_float(main_map.get("percent_pcs_small", 0.075)),
            "percent_pcs_large": parse_float(main_map.get("percent_pcs_large", 0.05)),
            "percent_m_small": parse_float(main_map.get("percent_m_small", 0.15)),
            "percent_m_large": parse_float(main_map.get("percent_m_large", 0.10)),
            "tubi_round_multiple": parse_int(main_map.get("tubi_round_multiple", 6)),
        },
        "erection": {
            "nwb_threshold_1": parse_int(erection_map.get("nwb_threshold_1", 30)),
            "nwb_threshold_2": parse_int(erection_map.get("nwb_threshold_2", 100)),
            "nwb_threshold_3": parse_int(erection_map.get("nwb_threshold_3", 400)),
            "nwb_percent_low": parse_float(erection_map.get("nwb_percent_low", 1.0)),
            "nwb_percent_mid": parse_float(erection_map.get("nwb_percent_mid", 0.5)),
            "nwb_percent_high": parse_float(erection_map.get("nwb_percent_high", 0.35)),
            "nwb_percent_very_high": parse_float(erection_map.get("nwb_percent_very_high", 0.2)),
            "gasket_threshold_1": parse_int(erection_map.get("gasket_threshold_1", 30)),
            "gasket_threshold_2": parse_int(erection_map.get("gasket_threshold_2", 500)),
            "gasket_percent_low": parse_float(erection_map.get("gasket_percent_low", 1.0)),
            "gasket_percent_mid": parse_float(erection_map.get("gasket_percent_mid", 0.5)),
            "gasket_percent_high": parse_float(erection_map.get("gasket_percent_high", 0.3)),
            "standard_threshold": parse_int(erection_map.get("standard_threshold", 400)),
            "standard_percent_low": parse_float(erection_map.get("standard_percent_low", 0.35)),
            "standard_percent_high": parse_float(erection_map.get("standard_percent_high", 0.2)),
        },
        "keywords": {
            "seal_required": kw.get("seal_required", ["seal", "ring", "fkm"]),
            "seal_excluded": kw.get("seal_excluded", ["elbow", "flange", "tee", "reducer"]),
            "critical": kw.get(
                "critical",
                ["nipple", "nipples", "union", "connection", "connections", "welding", "adaptor", "coupling", "threaded"],
            ),
            "tubi": kw.get("tubi", ["tubi"]),
            "nwb": kw.get("nwb", ["nut", "washer"]),
            "bolt": kw.get("bolt", ["bolt"]),
            "gasket": kw.get("gasket", ["gasket"]),
        },
        "blind_disk": {
            "thickness_by_dn": blind_disk_thickness,
            "columns": blind_disk_defaults,
        },
        "paint_erection": {
            "columns": paint_erection_defaults,
        },
        "flange_guard": {
            "system_keywords": flange_guard_system_keywords,
            "columns": flange_guard_defaults,
        },
    }


def ensure_table_column(conn: sqlite3.Connection, table_name: str, column_name: str, column_type: str = "TEXT") -> None:
    cols = conn.execute(f'PRAGMA table_info("{table_name}");').fetchall()
    existing = {str(c[1]) for c in cols}
    if column_name not in existing:
        conn.execute(f'ALTER TABLE "{table_name}" ADD COLUMN "{column_name}" {column_type};')


def to_numeric_cell(value: float) -> Any:
    if abs(value - round(value)) < 1e-12:
        return int(round(value))
    return value


def append_note(existing_note: Any, addition: str) -> str:
    base = "" if existing_note is None else str(existing_note).strip()
    if not base or base.upper() == NA_VALUE:
        return addition
    return f"{base} | {addition}"


def fmt_num(value: float) -> str:
    return str(to_numeric_cell(float(value)))


def records_to_column_map(records: List[Dict[str, object]]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for rec in records:
        col = str(rec.get("Column", "")).strip()
        val = str(rec.get("Value", "")).strip()
        if col:
            out[col] = val
    return out


def render_template(template: str, variables: Dict[str, Any]) -> str:
    rendered = str(template)
    for key, value in variables.items():
        rendered = rendered.replace("{" + str(key) + "}", str(value))
    return rendered


def calculate_touch_up_cans(surface_m2: float) -> int:
    if surface_m2 <= 0:
        return 0
    liters = surface_m2 / 5.0
    cans = ceil_int(liters / 20.0)
    if cans <= 0:
        return 0
    liters_in_not_full_can = liters - (20.0 * (cans - 1))
    if 15.0 <= liters_in_not_full_can < 20.0:
        cans += 1
    return cans


def normalize_dn_key(size_value: str) -> str:
    text = str(size_value or "").strip()
    if not text:
        return ""
    num = extract_numeric(text)
    if num > 0:
        return str(int(round(num))) if abs(num - round(num)) < 1e-12 else str(num)
    return text


def calculate_main_spare(row: sqlite3.Row, settings: Dict[str, Any]) -> Tuple[float, str]:
    c = settings
    m = settings["main"]
    kw = settings["keywords"]

    description = str(row[c["description_column"]] or "")
    type_label_text = str(row[c["type_label_column"]] or "") if c["type_label_column"] in row.keys() else ""
    type_text = str(row[c["type_column"]] or "")
    # VBA parses the numeric threshold value from the Type column.
    type_value = extract_numeric(type_text)

    qty_pcs = max(0.0, parse_float(row[c["qty_pcs_column"]], 0.0))
    qty_m = max(0.0, parse_float(row[c["qty_m_column"]], 0.0))

    is_seal = contains_all(description, kw["seal_required"]) and not contains_any(description, kw["seal_excluded"])
    is_critical = contains_any(description, kw["critical"])
    is_tubi = (
        contains_any(type_label_text, kw["tubi"])
        or contains_any(type_text, kw["tubi"])
        or contains_any(description, kw["tubi"])
    )

    if qty_pcs <= 0 and qty_m <= 0:
        return 0, "No quantity available (Qty pcs and Qty m are both 0)."

    qty = qty_pcs if qty_pcs > 0 else qty_m
    unit = "pcs" if qty_pcs > 0 else "m"

    if is_seal:
        if qty <= m["seal_qty_threshold"]:
            spare = m["seal_min_spare_low"]
            return (
                spare,
                f"Rule #1: Seal components qty <= {m['seal_qty_threshold']} -> spare = fixed minimum of {m['seal_min_spare_low']} {unit}.",
            )
        spare = max(m["seal_min_spare_high"], ceil_int(qty * m["seal_percent"]))
        return (
            spare,
            f"Rule #2: Seal components qty > {m['seal_qty_threshold']} -> spare = max({m['seal_min_spare_high']}, ceiling(qty * {m['seal_percent']:.0%})) {unit}.",
        )

    if is_critical:
        if qty <= m["critical_qty_threshold"]:
            spare = max(m["critical_min_spare"], ceil_int(qty * m["critical_percent_low"]))
            return (
                spare,
                f"Rule #3: Critical components qty <= {m['critical_qty_threshold']} -> spare = max({m['critical_min_spare']}, ceiling(qty * {m['critical_percent_low']:.0%})) {unit}.",
            )
        spare = max(m["critical_min_spare"], ceil_int(qty * m["critical_percent_high"]))
        return (
            spare,
            f"Rule #4: Critical components qty > {m['critical_qty_threshold']} -> spare = max({m['critical_min_spare']}, ceiling(qty * {m['critical_percent_high']:.0%})) {unit}.",
        )

    if is_tubi:
        pct = m["percent_m_small"] if 0 < type_value < m["size_threshold"] else m["percent_m_large"]
        base_spare = ceil_int(qty * pct)
        pre_rounded_total = qty + base_spare
        total = round_up_to_multiple(pre_rounded_total, m["tubi_round_multiple"])
        spare = max(0.0, float(total - qty))
        return (
            spare,
            f"Rule #T1: TUBI components -> base spare = ceiling(qty * {pct:.0%}) = {base_spare} {unit}; pre-rounded total = qty + base spare = {pre_rounded_total:.3f}; order qty = next multiple of {m['tubi_round_multiple']} -> {total}; spare = order qty - original qty = {spare:.3f} {unit}.",
        )

    if qty_pcs > 0:
        if qty <= m["threshold_1"]:
            return 0, f"Rule #5: Standard components qty <= {m['threshold_1']} -> spare = 0 pcs."
        if qty <= m["threshold_2"]:
            return 1, f"Rule #6: Standard components qty between {m['threshold_1'] + 1} and {m['threshold_2']} -> spare = 1 pcs."
        if type_value < m["size_threshold"]:
            spare = ceil_int(qty * m["percent_pcs_small"])
            return spare, f"Rule #7a: Standard components qty > {m['threshold_2']} and DN < {m['size_threshold']} -> spare = ceiling(qty * {m['percent_pcs_small']:.1%}) pcs."
        spare = ceil_int(qty * m["percent_pcs_large"])
        return spare, f"Rule #7b: Standard components qty > {m['threshold_2']} and DN >= {m['size_threshold']} -> spare = ceiling(qty * {m['percent_pcs_large']:.0%}) pcs."

    if type_value < m["size_threshold"]:
        spare = ceil_int(qty * m["percent_m_small"])
        return spare, f"Rule #8a: Standard meter qty with DN < {m['size_threshold']} -> spare = ceiling(qty * {m['percent_m_small']:.0%}) m."
    spare = ceil_int(qty * m["percent_m_large"])
    return spare, f"Rule #8b: Standard meter qty with DN >= {m['size_threshold']} -> spare = ceiling(qty * {m['percent_m_large']:.0%}) m."


def calculate_erection_spare(row: sqlite3.Row, settings: Dict[str, Any]) -> Tuple[int, str]:
    c = settings
    e = settings["erection"]
    kw = settings["keywords"]

    description = str(row[c["description_column"]] or "")
    qty = max(0.0, parse_float(row[c["qty_pcs_column"]], 0.0))

    if qty <= 0:
        return 0, "No quantity available (Qty pcs is 0)."

    is_nwb = contains_any(description, kw["nwb"])
    is_bolt = contains_any(description, kw["bolt"])
    is_gasket = contains_any(description, kw["gasket"])

    if is_nwb:
        if qty <= e["nwb_threshold_1"]:
            base = ceil_int(qty * e["nwb_percent_low"])
            total = round_up_nut_washer(qty + base)
            spare = int(total - qty)
            return spare, f"Rule #1: Nut/Washer qty <= {e['nwb_threshold_1']} -> base spare = {e['nwb_percent_low']:.0%} of qty = {base}; total rounded (<=100->10, 101-500->50, >500->100) = {total}; spare = {spare}."
        if qty <= e["nwb_threshold_2"]:
            base = ceil_int(qty * e["nwb_percent_mid"])
            total = round_up_nut_washer(qty + base)
            spare = int(total - qty)
            return spare, f"Rule #2: Nut/Washer qty between {e['nwb_threshold_1'] + 1} and {e['nwb_threshold_2']} -> base spare = {e['nwb_percent_mid']:.0%} of qty = {base}; total rounded (<=100->10, 101-500->50, >500->100) = {total}; spare = {spare}."
        if qty <= e["nwb_threshold_3"]:
            base = ceil_int(qty * e["nwb_percent_high"])
            total = round_up_nut_washer(qty + base)
            spare = int(total - qty)
            return spare, f"Rule #3: Nut/Washer qty between {e['nwb_threshold_2'] + 1} and {e['nwb_threshold_3']} -> base spare = {e['nwb_percent_high']:.0%} of qty = {base}; total rounded (<=100->10, 101-500->50, >500->100) = {total}; spare = {spare}."
        base = ceil_int(qty * e["nwb_percent_very_high"])
        total = round_up_nut_washer(qty + base)
        spare = int(total - qty)
        return spare, f"Rule #4: Nut/Washer qty > {e['nwb_threshold_3']} -> base spare = {e['nwb_percent_very_high']:.0%} of qty = {base}; total rounded (<=100->10, 101-500->50, >500->100) = {total}; spare = {spare}."

    if is_bolt:
        if qty <= e["nwb_threshold_1"]:
            base = ceil_int(qty * e["nwb_percent_low"])
            total = round_up_to_next_10(qty + base)
            spare = int(total - qty)
            return spare, f"Rule #1B: Bolt qty <= {e['nwb_threshold_1']} -> base spare = {e['nwb_percent_low']:.0%} of qty = {base}; total rounded to next 10 = {total}; spare = {spare}."
        if qty <= e["nwb_threshold_2"]:
            base = ceil_int(qty * e["nwb_percent_mid"])
            total = round_up_to_next_10(qty + base)
            spare = int(total - qty)
            return spare, f"Rule #2B: Bolt qty between {e['nwb_threshold_1'] + 1} and {e['nwb_threshold_2']} -> base spare = {e['nwb_percent_mid']:.0%} of qty = {base}; total rounded to next 10 = {total}; spare = {spare}."
        if qty <= e["nwb_threshold_3"]:
            base = ceil_int(qty * e["nwb_percent_high"])
            total = round_up_to_next_10(qty + base)
            spare = int(total - qty)
            return spare, f"Rule #3B: Bolt qty between {e['nwb_threshold_2'] + 1} and {e['nwb_threshold_3']} -> base spare = {e['nwb_percent_high']:.0%} of qty = {base}; total rounded to next 10 = {total}; spare = {spare}."
        base = ceil_int(qty * e["nwb_percent_very_high"])
        total = round_up_to_next_10(qty + base)
        spare = int(total - qty)
        return spare, f"Rule #4B: Bolt qty > {e['nwb_threshold_3']} -> base spare = {e['nwb_percent_very_high']:.0%} of qty = {base}; total rounded to next 10 = {total}; spare = {spare}."

    if is_gasket:
        if qty <= e["gasket_threshold_1"]:
            spare = ceil_int(qty * e["gasket_percent_low"])
            return spare, f"Rule #5: Gasket qty <= {e['gasket_threshold_1']} -> spare = {e['gasket_percent_low']:.0%} of qty rounded up to the next integer."
        if qty <= e["gasket_threshold_2"]:
            base = ceil_int(qty * e["gasket_percent_mid"])
            total = round_up_to_next_10(qty + base)
            spare = int(total - qty)
            return spare, f"Rule #6: Gasket qty between {e['gasket_threshold_1'] + 1} and {e['gasket_threshold_2']} -> base spare = {e['gasket_percent_mid']:.0%} of qty = {base}; total rounded to next 10 = {total}; spare = {spare}."
        base = ceil_int(qty * e["gasket_percent_high"])
        total = round_up_to_next_10(qty + base)
        spare = int(total - qty)
        return spare, f"Rule #7: Gasket qty > {e['gasket_threshold_2']} -> base spare = {e['gasket_percent_high']:.0%} of qty = {base}; total rounded to next 10 = {total}; spare = {spare}."

    if qty < e["standard_threshold"]:
        spare = ceil_int(qty * e["standard_percent_low"])
        return spare, f"Rule #8: Standard components qty < {e['standard_threshold']} -> spare = {e['standard_percent_low']:.0%} of qty rounded up to next integer."
    spare = ceil_int(qty * e["standard_percent_high"])
    return spare, f"Rule #9: Standard components qty >= {e['standard_threshold']} -> spare = {e['standard_percent_high']:.0%} of qty rounded up to next integer."


def resolve_aggregate_table_name(revision_table: str, settings: Dict[str, Any]) -> str:
    return f"{revision_table}{settings['aggregate_table_suffix']}"


def classify_material_type(material_type: str, settings: Dict[str, Any]) -> str:
    mt = material_type.strip()
    mt_l = mt.lower()
    if "paint shop" in mt_l:
        return "paint"
    if mt_l == settings["erection_exact"].lower() or ("erection" in mt_l and "paint" not in mt_l):
        return "erection"
    if settings["paint_prefix"] in mt_l and "shop" in mt_l:
        return "paint"
    if mt_l.endswith(settings["shop_suffix"]):
        return "shop"
    return "other"


def run_spare_calcs(db_path: Path, revision_table: str, settings_workbook: Path) -> int:
    settings = load_spare_settings(settings_workbook)
    agg_table = resolve_aggregate_table_name(revision_table, settings)

    if not db_path.exists():
        logger.error("Database not found: %s", db_path)
        return 1

    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row

        exists = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1;",
            (agg_table,),
        ).fetchone()
        if not exists:
            logger.error("Aggregated table not found: %s", agg_table)
            return 1

        rows = conn.execute(f'SELECT rowid, * FROM "{agg_table}"').fetchall()
        if not rows:
            logger.info("No rows to process in %s", agg_table)
            return 0

        spare_col = settings["spare_column"]
        note_col = settings["spare_info_column"]
        mt_col = settings["material_type_column"]
        qty_pcs_col = settings["qty_pcs_column"]
        qty_m_col = settings["qty_m_column"]
        weight_col = settings["weight_column"]
        total_qty_col = settings["total_qty_column"]
        weight_w_spare_col = settings["weight_w_spare_column"]
        size_col = settings["size_column"]
        type_col = settings["type_column"]
        desc_col = settings["description_column"]
        tubi_cs_pre_col = TUBI_CS_SHOP_PRE_PAINT_COL
        tubi_cs_post_col = TUBI_CS_SHOP_POST_PAINT_COL
        surface_w_spare_col = SURFACE_W_SPARE_COL

        ensure_table_column(conn, agg_table, spare_col, "NUMERIC")
        ensure_table_column(conn, agg_table, note_col, "TEXT")
        ensure_table_column(conn, agg_table, total_qty_col, "NUMERIC")
        ensure_table_column(conn, agg_table, weight_w_spare_col, "NUMERIC")
        ensure_table_column(conn, agg_table, tubi_cs_pre_col, "NUMERIC")
        ensure_table_column(conn, agg_table, tubi_cs_post_col, "NUMERIC")
        ensure_table_column(conn, agg_table, surface_w_spare_col, "NUMERIC")

        row_states: List[Dict[str, Any]] = []
        shop_count = 0
        paint_count = 0
        erection_count = 0
        skipped_count = 0

        for row in rows:
            material_type = str(row[mt_col] or "")
            cls = classify_material_type(material_type, settings)

            if cls in ("shop", "paint"):
                spare, note = calculate_main_spare(row, settings)
                if cls == "shop":
                    shop_count += 1
                else:
                    paint_count += 1
            elif cls == "erection":
                spare, note = calculate_erection_spare(row, settings)
                erection_count += 1
            else:
                spare = parse_float(row[spare_col], 0.0) if spare_col in row.keys() else 0.0
                note = row[note_col] if note_col in row.keys() and row[note_col] is not None else NA_VALUE
                skipped_count += 1

            qty_pcs = max(0.0, parse_float(row[qty_pcs_col], 0.0))
            qty_m = max(0.0, parse_float(row[qty_m_col], 0.0))
            qty_base = qty_pcs if qty_pcs > 0 else qty_m
            spare_num = max(0.0, parse_float(spare, 0.0))
            total_qty = qty_base + spare_num
            weight_value = max(0.0, parse_float(row[weight_col], 0.0))
            unit_weight = (weight_value / qty_base) if qty_base > 0 else 0.0
            weight_w_spare = total_qty * unit_weight

            row_states.append(
                {
                    "rowid": int(row["rowid"]),
                    "cls": cls,
                    "material_type": material_type,
                    "size": str(row[size_col] or "").strip(),
                    "type_text": str(row[type_col] or ""),
                    "description": str(row[desc_col] or ""),
                    "paint_colour": str(row["Paint colour"] or "") if "Paint colour" in row.keys() else "",
                    "surface_m2": max(0.0, parse_float(row["Surface m2"], 0.0)) if "Surface m2" in row.keys() else 0.0,
                    "qty_base": qty_base,
                    "spare_num": spare_num,
                    "note": note,
                    "total_qty": total_qty,
                    "unit_weight": unit_weight,
                    "weight_w_spare": weight_w_spare,
                    "tubi_cs_shop_pre_paint": None,
                    "tubi_cs_shop_post_paint": None,
                    "surface_w_spare": None,
                }
            )

        for state in row_states:
            if state["cls"] != "paint":
                continue
            qty_base = float(state.get("qty_base", 0.0) or 0.0)
            total_qty = float(state.get("total_qty", 0.0) or 0.0)
            surface_raw = float(state.get("surface_m2", 0.0) or 0.0)
            if surface_raw <= 0 or total_qty <= 0:
                state["surface_w_spare"] = 0.0
                continue
            scale = (total_qty / qty_base) if qty_base > 0 else 1.0
            state["surface_w_spare"] = surface_raw * scale

        tubi_keywords = settings.get("keywords", {}).get("tubi", ["tubi"])

        def is_tubi_state(state: Dict[str, Any]) -> bool:
            return contains_any(state["type_text"], tubi_keywords) or contains_any(state["description"], tubi_keywords)

        def size_key(value: str) -> str:
            return value.strip().lower()

        for state in row_states:
            is_cs_shop_tubi = (
                state["cls"] == "shop"
                and "cs" in state["material_type"].lower()
                and is_tubi_state(state)
            )
            if is_cs_shop_tubi:
                state["tubi_cs_shop_pre_paint"] = float(state["total_qty"])
                state["tubi_cs_shop_post_paint"] = float(state["total_qty"])

        paint_totals_by_size: Dict[str, float] = {}
        paint_rows_by_size: Dict[str, List[Dict[str, Any]]] = {}
        size_display: Dict[str, str] = {}
        for state in row_states:
            if state["cls"] != "paint" or not is_tubi_state(state):
                continue
            key = size_key(state["size"])
            paint_totals_by_size[key] = paint_totals_by_size.get(key, 0.0) + float(state["total_qty"])
            paint_rows_by_size.setdefault(key, []).append(state)
            if key and key not in size_display:
                size_display[key] = state["size"]

        cs_shop_by_size: Dict[str, List[Dict[str, Any]]] = {}
        for state in row_states:
            if state["cls"] != "shop":
                continue
            if "cs" not in state["material_type"].lower():
                continue
            if not is_tubi_state(state):
                continue
            key = size_key(state["size"])
            cs_shop_by_size.setdefault(key, []).append(state)
            if key and key not in size_display:
                size_display[key] = state["size"]

        adjusted_size_count = 0
        warnings: List[str] = []

        for key, paint_total in paint_totals_by_size.items():
            cs_rows = cs_shop_by_size.get(key, [])
            paint_rows = paint_rows_by_size.get(key, [])
            shown_size = size_display.get(key, key if key else "<empty>")

            if not cs_rows:
                warnings.append(
                    f"TUBI size '{shown_size}': paint Order qty sum is {fmt_num(paint_total)} but no CS shop TUBI row exists."
                )
                for p_row in paint_rows:
                    p_row["note"] = append_note(
                        p_row["note"],
                        (
                            f"Post-process TUBI size check: size {shown_size}; paint total Order qty={fmt_num(paint_total)}; "
                            "CS shop total Order qty before=missing; no CS adjustment possible."
                        ),
                    )
                continue

            cs_total = sum(float(r["total_qty"]) for r in cs_rows)
            cs_total_before = cs_total
            diff = float(paint_total - cs_total)

            if diff > 1e-12:
                target = max(cs_rows, key=lambda r: float(r["total_qty"]))
                previous_cs_total = cs_total
                target["spare_num"] = float(target["spare_num"]) + diff
                target["total_qty"] = float(target["total_qty"]) + diff
                target["weight_w_spare"] = float(target["total_qty"]) * float(target["unit_weight"])
                target["note"] = append_note(
                    target["note"],
                    (
                        f"Post-process TUBI size match: summed paint Order qty ({fmt_num(paint_total)}) "
                        f"was higher than CS shop ({fmt_num(previous_cs_total)}). "
                        f"Added spare +{fmt_num(diff)} to align CS Order qty with paint."
                    ),
                )
                for cs_row in cs_rows:
                    cs_row["tubi_cs_shop_post_paint"] = float(cs_row["total_qty"])
                    if cs_row is target:
                        continue
                    cs_row["note"] = append_note(
                        cs_row["note"],
                        (
                            f"Post-process TUBI size check: size {shown_size}; paint total Order qty={fmt_num(paint_total)}; "
                            f"CS shop total Order qty before={fmt_num(cs_total_before)}; "
                            f"delta(paint-CS)={fmt_num(diff)}; adjustment applied on another CS row for this size."
                        ),
                    )
                for p_row in paint_rows:
                    p_row["note"] = append_note(
                        p_row["note"],
                        (
                            f"Post-process TUBI size check: size {shown_size}; paint total Order qty={fmt_num(paint_total)}; "
                            f"CS shop total Order qty before={fmt_num(cs_total_before)}; "
                            f"delta(paint-CS)={fmt_num(diff)}; CS spare adjusted to match paint total."
                        ),
                    )
                adjusted_size_count += 1
            elif diff < -1e-12:
                for cs_row in cs_rows:
                    cs_row["tubi_cs_shop_post_paint"] = float(cs_row["total_qty"])
                warnings.append(
                    f"TUBI size '{shown_size}': paint Order qty sum ({fmt_num(paint_total)}) is lower than CS shop ({fmt_num(cs_total)})."
                )
                for cs_row in cs_rows:
                    cs_row["note"] = append_note(
                        cs_row["note"],
                        (
                            f"Post-process TUBI size check: size {shown_size}; paint total Order qty={fmt_num(paint_total)}; "
                            f"CS shop total Order qty before={fmt_num(cs_total_before)}; "
                            f"delta(paint-CS)={fmt_num(diff)}; no CS spare increase applied because paint is lower."
                        ),
                    )
                for p_row in paint_rows:
                    p_row["note"] = append_note(
                        p_row["note"],
                        (
                            f"Post-process TUBI size check: size {shown_size}; paint total Order qty={fmt_num(paint_total)}; "
                            f"CS shop total Order qty before={fmt_num(cs_total_before)}; "
                            f"delta(paint-CS)={fmt_num(diff)}; no CS spare increase applied because paint is lower."
                        ),
                    )
            else:
                for cs_row in cs_rows:
                    cs_row["tubi_cs_shop_post_paint"] = float(cs_row["total_qty"])
                for cs_row in cs_rows:
                    cs_row["note"] = append_note(
                        cs_row["note"],
                        (
                            f"Post-process TUBI size check: size {shown_size}; paint total Order qty={fmt_num(paint_total)}; "
                            f"CS shop total Order qty before={fmt_num(cs_total_before)}; delta(paint-CS)=0; already matched."
                        ),
                    )
                for p_row in paint_rows:
                    p_row["note"] = append_note(
                        p_row["note"],
                        (
                            f"Post-process TUBI size check: size {shown_size}; paint total Order qty={fmt_num(paint_total)}; "
                            f"CS shop total Order qty before={fmt_num(cs_total_before)}; delta(paint-CS)=0; already matched."
                        ),
                    )

        paint_erection_by_colour: Dict[str, float] = {}
        for state in row_states:
            if state["cls"] != "paint":
                continue

            paint_colour_raw = str(state.get("paint_colour", "") or "").strip()
            if not paint_colour_raw or paint_colour_raw.upper() == NA_VALUE:
                continue

            surface_with_spare = float(state.get("surface_w_spare", 0.0) or 0.0)
            if surface_with_spare <= 0:
                continue

            factor = 0.05 if is_tubi_state(state) else 1.0
            surface = surface_with_spare * factor
            if surface <= 0:
                continue

            paint_colour_l = paint_colour_raw.lower()
            if paint_colour_l in PAINT_BASIC_OR_PRIMER_NAMES:
                paint_erection_by_colour[PAINT_BASIC_BUCKET] = paint_erection_by_colour.get(PAINT_BASIC_BUCKET, 0.0) + (2.0 * surface)
            else:
                paint_erection_by_colour[paint_colour_raw] = paint_erection_by_colour.get(paint_colour_raw, 0.0) + surface
                paint_erection_by_colour[PAINT_BASIC_BUCKET] = paint_erection_by_colour.get(PAINT_BASIC_BUCKET, 0.0) + (2.0 * surface)

        existing_columns = [str(c[1]) for c in conn.execute(f'PRAGMA table_info("{agg_table}");').fetchall()]
        col_set = set(existing_columns)

        uid_values = conn.execute(f'SELECT "UID" FROM "{agg_table}";').fetchall() if "UID" in col_set else []
        max_uid = 0
        for (uid_val,) in uid_values:
            uid_num = parse_float(uid_val, 0.0)
            if uid_num > max_uid:
                max_uid = int(uid_num)
        next_uid = max_uid + 1

        paint_erection_rows: List[Dict[str, Any]] = []
        paint_erection_defaults = settings.get("paint_erection", {}).get("columns", {})
        for colour, total_surface in sorted(paint_erection_by_colour.items(), key=lambda kv: kv[0].lower()):
            cans = calculate_touch_up_cans(total_surface)
            if cans <= 0:
                continue

            new_row: Dict[str, Any] = {col: NA_VALUE for col in existing_columns if col != "UID"}
            if "UID" in col_set:
                new_row["UID"] = next_uid
                next_uid += 1
            liters = total_surface / 5.0
            liters_in_not_full_can = liters - (20.0 * (ceil_int(liters / 20.0) - 1)) if liters > 0 else 0.0
            paint_ctx = {
                "paint_colour": colour,
                "surface_m2": fmt_num(total_surface),
                "liters": fmt_num(liters),
                "cans": cans,
                "liters_in_not_full_can": fmt_num(liters_in_not_full_can),
            }
            for column_name, template_value in paint_erection_defaults.items():
                target_col = mt_col if column_name == "Material type" else column_name
                if target_col in col_set:
                    new_row[target_col] = render_template(template_value, paint_ctx)
            if "Surface m2" in col_set:
                new_row["Surface m2"] = to_numeric_cell(total_surface)
            if qty_pcs_col in col_set:
                new_row[qty_pcs_col] = cans
            if total_qty_col in col_set:
                new_row[total_qty_col] = cans

            paint_erection_rows.append(new_row)

        flan_qty_by_size: Dict[str, float] = {}
        size_display_by_key: Dict[str, str] = {}
        for state in row_states:
            if not (contains_any(state["type_text"], ["flan"]) or contains_any(state["description"], ["flan"])):
                continue
            size_text = str(state.get("size", "") or "").strip()
            size_key = normalize_dn_key(size_text)
            if not size_key:
                continue
            flan_qty_by_size[size_key] = flan_qty_by_size.get(size_key, 0.0) + float(state.get("qty_base", 0.0) or 0.0)
            if size_key not in size_display_by_key:
                size_display_by_key[size_key] = size_text if size_text else size_key

        thickness_map = settings.get("blind_disk", {}).get("thickness_by_dn", {})
        blind_disk_defaults = settings.get("blind_disk", {}).get("columns", {})
        blind_disk_rows: List[Dict[str, Any]] = []
        for size_key, qty_sum in sorted(flan_qty_by_size.items(), key=lambda kv: (extract_numeric(kv[0]), kv[0])):
            if qty_sum > 50:
                order_pcs = 10
            elif qty_sum > 30:
                order_pcs = 4
            else:
                order_pcs = 2

            thickness = str(thickness_map.get(size_key, NA_VALUE)).strip() or NA_VALUE
            add_info_key = "Add. Info" if thickness != NA_VALUE else "Add. Info Missing Thickness"

            row: Dict[str, Any] = {col: NA_VALUE for col in existing_columns if col != "UID"}
            if "UID" in col_set:
                row["UID"] = next_uid
                next_uid += 1
            blind_ctx = {
                "dn": size_key,
                "size": size_display_by_key.get(size_key, size_key),
                "thickness": thickness,
                "flange_qty": fmt_num(qty_sum),
                "order_qty": order_pcs,
            }
            for column_name, template_value in blind_disk_defaults.items():
                if column_name in {"Add. Info", "Add. Info Missing Thickness"}:
                    continue
                target_col = mt_col if column_name == "Material type" else column_name
                target_col = note_col if column_name == "Spare info" else target_col
                target_col = size_col if column_name == "Size" else target_col
                if target_col in col_set:
                    row[target_col] = render_template(template_value, blind_ctx)
            if qty_pcs_col in col_set:
                row[qty_pcs_col] = order_pcs
            if total_qty_col in col_set:
                row[total_qty_col] = order_pcs
            if "Add. Info" in col_set:
                add_info_template = blind_disk_defaults.get(add_info_key, "")
                row["Add. Info"] = render_template(add_info_template, blind_ctx) if add_info_template else NA_VALUE

            blind_disk_rows.append(row)

        updates: List[Tuple[Any, Any, Any, Any, Any, Any, Any, int]] = []
        for state in row_states:
            if state["cls"] == "other":
                tubi_pre_val: Any = NA_VALUE
                tubi_post_val: Any = NA_VALUE
                surface_w_spare_val: Any = NA_VALUE
            else:
                tubi_pre_val = (
                    to_numeric_cell(float(state["tubi_cs_shop_pre_paint"]))
                    if state["tubi_cs_shop_pre_paint"] is not None
                    else None
                )
                tubi_post_val = (
                    to_numeric_cell(float(state["tubi_cs_shop_post_paint"]))
                    if state["tubi_cs_shop_post_paint"] is not None
                    else None
                )
                surface_w_spare_val = (
                    to_numeric_cell(float(state["surface_w_spare"]))
                    if state["surface_w_spare"] is not None
                    else None
                )

            updates.append(
                (
                    to_numeric_cell(float(state["spare_num"])),
                    state["note"],
                    to_numeric_cell(float(state["total_qty"])),
                    to_numeric_cell(float(state["weight_w_spare"])),
                    tubi_pre_val,
                    tubi_post_val,
                    surface_w_spare_val,
                    int(state["rowid"]),
                )
            )

        if updates:
            conn.executemany(
                f'UPDATE "{agg_table}" SET "{spare_col}"=?, "{note_col}"=?, "{total_qty_col}"=?, "{weight_w_spare_col}"=?, "{tubi_cs_pre_col}"=?, "{tubi_cs_post_col}"=?, "{surface_w_spare_col}"=? WHERE rowid=?',
                updates,
            )

        # Keep helper calculated columns explicit in exports: fill non-applicable NULLs with N/A.
        conn.execute(f'UPDATE "{agg_table}" SET "{tubi_cs_pre_col}"=? WHERE "{tubi_cs_pre_col}" IS NULL', (NA_VALUE,))
        conn.execute(f'UPDATE "{agg_table}" SET "{tubi_cs_post_col}"=? WHERE "{tubi_cs_post_col}" IS NULL', (NA_VALUE,))
        conn.execute(f'UPDATE "{agg_table}" SET "{surface_w_spare_col}"=? WHERE "{surface_w_spare_col}" IS NULL', (NA_VALUE,))

        if paint_erection_rows:
            insert_cols = [col for col in existing_columns if any(col in row for row in paint_erection_rows)]
            placeholders = ", ".join(["?" for _ in insert_cols])
            cols_sql = ", ".join([f'"{c}"' for c in insert_cols])
            insert_sql = f'INSERT INTO "{agg_table}" ({cols_sql}) VALUES ({placeholders})'
            insert_values = [[row.get(col, None) for col in insert_cols] for row in paint_erection_rows]
            conn.executemany(insert_sql, insert_values)

        if blind_disk_rows:
            insert_cols = [col for col in existing_columns if any(col in row for row in blind_disk_rows)]
            placeholders = ", ".join(["?" for _ in insert_cols])
            cols_sql = ", ".join([f'"{c}"' for c in insert_cols])
            insert_sql = f'INSERT INTO "{agg_table}" ({cols_sql}) VALUES ({placeholders})'
            insert_values = [[row.get(col, None) for col in insert_cols] for row in blind_disk_rows]
            conn.executemany(insert_sql, insert_values)

        flange_guard_settings = settings.get("flange_guard", {})
        flange_guard_keywords = [str(v).strip().lower() for v in flange_guard_settings.get("system_keywords", []) if str(v).strip()]
        flange_guard_defaults = flange_guard_settings.get("columns", {})
        flange_guard_rows: List[Dict[str, Any]] = []

        base_rows = conn.execute(
            f'SELECT "Material type", "Material", "Type", "Description", "Size", "System", "Qty pcs", "Qty m" FROM "{revision_table}"'
        ).fetchall()

        flange_group_totals: Dict[Tuple[str, str, str], float] = {}
        for base_row in base_rows:
            material_type_text = str(base_row[0] or "")
            if "paint shop" in material_type_text.lower():
                continue

            type_text = str(base_row[2] or "")
            if not contains_any(type_text, ["flan"]):
                continue

            system_text = str(base_row[5] or "").strip()
            system_l = system_text.lower()
            if not flange_guard_keywords or not any(k in system_l for k in flange_guard_keywords):
                continue

            material_text = str(base_row[1] or "").strip() or NA_VALUE
            size_text = str(base_row[4] or "").strip() or NA_VALUE
            qty_pcs = max(0.0, parse_float(base_row[6], 0.0))
            qty_m = max(0.0, parse_float(base_row[7], 0.0))
            qty_base = qty_pcs if qty_pcs > 0 else qty_m
            if qty_base <= 0:
                continue

            key = (material_text, size_text, system_text or NA_VALUE)
            flange_group_totals[key] = flange_group_totals.get(key, 0.0) + qty_base

        for (material_text, size_text, system_text), qty_sum in sorted(
            flange_group_totals.items(),
            key=lambda kv: (str(kv[0][2]).lower(), str(kv[0][0]).lower(), extract_numeric(kv[0][1]), str(kv[0][1])),
        ):
            spare_fg = max(1, ceil_int(qty_sum * 0.05))
            order_fg = qty_sum + spare_fg

            row: Dict[str, Any] = {col: NA_VALUE for col in existing_columns if col != "UID"}
            if "UID" in col_set:
                row["UID"] = next_uid
                next_uid += 1

            fg_ctx = {
                "material": material_text,
                "size": size_text,
                "system": system_text,
                "flange_qty": fmt_num(qty_sum),
                "spare_qty": fmt_num(spare_fg),
                "order_qty": fmt_num(order_fg),
            }

            for column_name, template_value in flange_guard_defaults.items():
                target_col = mt_col if column_name == "Material type" else column_name
                target_col = note_col if column_name == "Spare info" else target_col
                target_col = size_col if column_name == "Size" else target_col
                if target_col in col_set:
                    row[target_col] = render_template(template_value, fg_ctx)

            if size_col in col_set:
                row[size_col] = size_text
            if "System" in col_set:
                row["System"] = system_text
            if qty_pcs_col in col_set:
                row[qty_pcs_col] = to_numeric_cell(qty_sum)
            if spare_col in col_set:
                row[spare_col] = to_numeric_cell(spare_fg)
            if total_qty_col in col_set:
                row[total_qty_col] = to_numeric_cell(order_fg)

            flange_guard_rows.append(row)

        if flange_guard_rows:
            insert_cols = [col for col in existing_columns if any(col in row for row in flange_guard_rows)]
            placeholders = ", ".join(["?" for _ in insert_cols])
            cols_sql = ", ".join([f'"{c}"' for c in insert_cols])
            insert_sql = f'INSERT INTO "{agg_table}" ({cols_sql}) VALUES ({placeholders})'
            insert_values = [[row.get(col, None) for col in insert_cols] for row in flange_guard_rows]
            conn.executemany(insert_sql, insert_values)

        rewrite_aggregated_uids(conn, agg_table, mt_col)

        conn.commit()

        logger.info("Updated spare values in table '%s'.", agg_table)
        logger.info("Shop rows processed: %s", shop_count)
        logger.info("Paint rows processed with main rules: %s", paint_count)
        logger.info("Erection rows processed: %s", erection_count)
        logger.info("Skipped rows (other): %s", skipped_count)
        logger.info("TUBI post-process adjusted CS sizes: %s", adjusted_size_count)
        logger.info("Paint erection rows inserted: %s", len(paint_erection_rows))
        logger.info("Blind disk rows inserted: %s", len(blind_disk_rows))
        logger.info("Flange guard rows inserted: %s", len(flange_guard_rows))
        for warning_text in warnings:
            logger.warning("%s", warning_text)

    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Apply spare calculations to aggregated revision table (shop + paint + erection).",
    )
    parser.add_argument(
        "--db-path",
        type=Path,
        default=Path("Price_Sheet_New_Format/price_sheet.db"),
        help="Path to SQLite database.",
    )
    parser.add_argument(
        "--revision-table",
        type=str,
        default="rev0",
        help="Base revision table name (without aggregate suffix), e.g. rev0.",
    )
    parser.add_argument(
        "--settings-workbook",
        type=Path,
        default=Path("Price_Sheet_New_Format/import_settings.md"),
        help="Path to Markdown settings file containing spare rule sections.",
    )
    parser.add_argument(
        "--log-dir",
        type=Path,
        default=DEFAULT_LOG_DIR,
        help="Directory where timestamped CSV run logs are written.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    log_file = setup_csv_logging(run_name="run_spare_calcs", log_dir=args.log_dir)
    logger.info("CSV logging initialized. log_file=%s", log_file)
    try:
        return run_spare_calcs(
            db_path=args.db_path,
            revision_table=args.revision_table,
            settings_workbook=args.settings_workbook,
        )
    except Exception:
        logger.exception(
            "Spare calculation failed. "
            + sections_hint("SpareGeneral", "SpareMainRules", "SpareErectionRules", "SpareKeywords")
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
