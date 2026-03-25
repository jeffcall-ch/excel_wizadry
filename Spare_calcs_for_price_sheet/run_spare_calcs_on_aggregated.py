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
    if mt_l == settings["erection_exact"].lower():
        return "erection"
    if mt_l.startswith(settings["paint_prefix"]):
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

        ensure_table_column(conn, agg_table, spare_col, "NUMERIC")
        ensure_table_column(conn, agg_table, note_col, "TEXT")
        ensure_table_column(conn, agg_table, total_qty_col, "NUMERIC")
        ensure_table_column(conn, agg_table, weight_w_spare_col, "NUMERIC")

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
                    "spare_num": spare_num,
                    "note": note,
                    "total_qty": total_qty,
                    "unit_weight": unit_weight,
                    "weight_w_spare": weight_w_spare,
                }
            )

        tubi_keywords = settings.get("keywords", {}).get("tubi", ["tubi"])

        def is_tubi_state(state: Dict[str, Any]) -> bool:
            return contains_any(state["type_text"], tubi_keywords) or contains_any(state["description"], tubi_keywords)

        def size_key(value: str) -> str:
            return value.strip().lower()

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

        updates: List[Tuple[Any, Any, Any, Any, int]] = []
        for state in row_states:
            updates.append(
                (
                    to_numeric_cell(float(state["spare_num"])),
                    state["note"],
                    to_numeric_cell(float(state["total_qty"])),
                    to_numeric_cell(float(state["weight_w_spare"])),
                    int(state["rowid"]),
                )
            )

        if updates:
            conn.executemany(
                f'UPDATE "{agg_table}" SET "{spare_col}"=?, "{note_col}"=?, "{total_qty_col}"=?, "{weight_w_spare_col}"=? WHERE rowid=?',
                updates,
            )
            conn.commit()

        logger.info("Updated spare values in table '%s'.", agg_table)
        logger.info("Shop rows processed: %s", shop_count)
        logger.info("Paint rows processed with main rules: %s", paint_count)
        logger.info("Erection rows processed: %s", erection_count)
        logger.info("Skipped rows (other): %s", skipped_count)
        logger.info("TUBI post-process adjusted CS sizes: %s", adjusted_size_count)
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
