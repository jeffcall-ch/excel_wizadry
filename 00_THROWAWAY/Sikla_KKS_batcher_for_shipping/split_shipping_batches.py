from __future__ import annotations

import argparse
import math
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_INPUT_FILENAME = (
    "CA100-KVI-50296373_2.0 - BoM General Hangers and support material "
    "LS_erection_split.xlsm"
)
DEFAULT_SHEET_NAME = "Per_Support_Structure"
EXCLUDED_WORKING_AREAS = {"", "NO_AREA", "BQ NOT EXIST"}
REQUIRED_COLUMNS = ["Support Structure", "Working area", "Total Weight [kg]"]
DEFAULT_SECTION_ORDER = ["UHA", "UMA"]

SECTION_ORDER: list[str] = DEFAULT_SECTION_ORDER.copy()
SECTION_ORDER_MAP: dict[str, int] = {
    section: index for index, section in enumerate(SECTION_ORDER)
}

DETAIL_COLUMN_WIDTHS = {
    "A": 18.57,
    "B": 27.14,
    "C": 24.86,
    "D": 19.86,
    "E": 18.00,
    "F": 13.00,
    "G": 13.00,
    "H": 12.00,
    "I": 12.00,
}

SUMMARY_COLUMN_WIDTHS = {
    "A": 13.43,
    "B": 16.14,
    "C": 30.43,
    "D": 13.00,
    "E": 10.57,
    "F": 17.43,
    "G": 16.14,
    "H": 16.43,
    "I": 18.43,
    "J": 17.00,
}


@dataclass
class SectionChunk:
    building_section: str
    chunk_id: int
    working_areas: list[str]
    weight_kg: float


@dataclass
class BatchBin:
    batch_number: int
    weight_kg: float = 0.0
    chunks: list[SectionChunk] = field(default_factory=list)
    sections: set[str] = field(default_factory=set)


def derive_building_section(working_area: str) -> str:
    match = re.search(r"U[A-Z]{2}", str(working_area).upper())
    return match.group(0) if match else "UNKNOWN"


def parse_section_order(value: str) -> list[str]:
    sections = [part.strip().upper() for part in str(value).split(",") if part.strip()]
    deduped: list[str] = []
    seen: set[str] = set()
    for section in sections:
        if section not in seen:
            seen.add(section)
            deduped.append(section)
    return deduped


def set_section_order(section_order: list[str]) -> None:
    global SECTION_ORDER
    global SECTION_ORDER_MAP

    SECTION_ORDER = section_order if section_order else DEFAULT_SECTION_ORDER.copy()
    SECTION_ORDER_MAP = {section: index for index, section in enumerate(SECTION_ORDER)}


def natural_sort_token(text: str) -> str:
    value = str(text).strip().upper()
    return re.sub(r"\d+", lambda m: f"{int(m.group()):06d}", value)


def section_sort_token(section: str) -> str:
    key = str(section).strip().upper()
    rank = SECTION_ORDER_MAP.get(key)
    if rank is not None:
        return f"{rank:03d}_{key}"
    return f"900_{key}"


def batch_min_working_area_token(batch: BatchBin, section: str | None = None) -> str:
    areas: list[str] = []
    for chunk in batch.chunks:
        if section is None or chunk.building_section == section:
            areas.extend(chunk.working_areas)

    if not areas:
        return "ZZZZZZ"
    return min(natural_sort_token(area) for area in areas)


def load_scope(input_file: Path, sheet_name: str) -> pd.DataFrame:
    data = pd.read_excel(input_file, sheet_name=sheet_name, engine="openpyxl")

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in data.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {missing_columns}")

    scope = data.loc[:, REQUIRED_COLUMNS].copy()
    scope["Total Weight [kg]"] = pd.to_numeric(
        scope["Total Weight [kg]"], errors="coerce"
    ).fillna(0.0)

    scope = scope[scope["Working area"].notna()].copy()
    scope["Working area"] = scope["Working area"].astype(str).str.strip()
    scope = scope[~scope["Working area"].isin(EXCLUDED_WORKING_AREAS)].copy()

    scope = scope[scope["Support Structure"].notna()].copy()
    scope["Support Structure"] = scope["Support Structure"].astype(str).str.strip()
    scope = scope[
        (scope["Support Structure"] != "")
        & (scope["Support Structure"].str.lower() != "nan")
    ].copy()

    return scope


def split_section_into_chunks(
    section_name: str, section_rows: pd.DataFrame, split_count: int
) -> list[SectionChunk]:
    if split_count < 1:
        split_count = 1

    bins = [{"weight": 0.0, "areas": []} for _ in range(split_count)]
    rows = section_rows.sort_values(["Total Weight [kg]", "Working area"], ascending=[False, True])

    for _, row in rows.iterrows():
        lightest_idx = min(range(split_count), key=lambda i: bins[i]["weight"])
        bins[lightest_idx]["areas"].append(row["Working area"])
        bins[lightest_idx]["weight"] += float(row["Total Weight [kg]"])

    chunks: list[SectionChunk] = []
    for idx, bucket in enumerate(bins, start=1):
        if not bucket["areas"]:
            continue
        chunks.append(
            SectionChunk(
                building_section=section_name,
                chunk_id=idx,
                working_areas=sorted(bucket["areas"]),
                weight_kg=float(bucket["weight"]),
            )
        )

    return chunks


def create_chunks(wa_summary: pd.DataFrame, max_batch_weight: float) -> list[SectionChunk]:
    chunks: list[SectionChunk] = []

    for section_name, section_rows in wa_summary.groupby("Building section"):
        section_total = float(section_rows["Total Weight [kg]"].sum())
        split_count = max(1, math.ceil(section_total / max_batch_weight))

        chunks.extend(
            split_section_into_chunks(
                section_name=section_name,
                section_rows=section_rows,
                split_count=split_count,
            )
        )

    # Bigger chunks first for stable balancing.
    chunks.sort(key=lambda c: (-c.weight_kg, c.building_section, c.chunk_id))
    return chunks


def select_batch(
    chunk: SectionChunk,
    batches: list[BatchBin],
    section_batches: dict[str, set[int]],
    max_batch_weight: float,
) -> BatchBin:
    assigned_batches = section_batches.get(chunk.building_section, set())

    if assigned_batches:
        preferred = [b for b in batches if b.batch_number in assigned_batches]
    else:
        preferred = batches

    def lightest_fit(candidates: list[BatchBin]) -> BatchBin:
        return min(candidates, key=lambda b: (b.weight_kg, b.batch_number))

    non_overflow = [b for b in preferred if (b.weight_kg + chunk.weight_kg) <= max_batch_weight]
    if non_overflow:
        return lightest_fit(non_overflow)

    if assigned_batches:
        non_overflow_all = [
            b for b in batches if (b.weight_kg + chunk.weight_kg) <= max_batch_weight
        ]
        if non_overflow_all:
            return lightest_fit(non_overflow_all)

    return lightest_fit(preferred)


def assign_chunks_to_batches(
    chunks: list[SectionChunk],
    batch_count: int,
    max_batch_weight: float,
) -> list[BatchBin]:
    batches = [BatchBin(batch_number=i + 1) for i in range(batch_count)]
    section_batches: dict[str, set[int]] = {}

    for chunk in chunks:
        batch = select_batch(
            chunk=chunk,
            batches=batches,
            section_batches=section_batches,
            max_batch_weight=max_batch_weight,
        )

        batch.chunks.append(chunk)
        batch.weight_kg += chunk.weight_kg
        batch.sections.add(chunk.building_section)
        section_batches.setdefault(chunk.building_section, set()).add(batch.batch_number)

    return batches


def reorder_batches_for_output(batches: list[BatchBin]) -> list[BatchBin]:
    if not batches:
        return batches

    ordered: list[BatchBin] = []
    selected_ids: set[int] = set()

    for section in SECTION_ORDER:
        section_batches = sorted(
            [
                batch
                for batch in batches
                if section in batch.sections and id(batch) not in selected_ids
            ],
            key=lambda b: (batch_min_working_area_token(b, section=section), b.batch_number),
        )
        ordered.extend(section_batches)
        selected_ids.update(id(batch) for batch in section_batches)

    remaining = [batch for batch in batches if id(batch) not in selected_ids]
    remaining_sorted = sorted(
        remaining,
        key=lambda b: (
            min(section_sort_token(section) for section in b.sections)
            if b.sections
            else "99_UNKNOWN",
            batch_min_working_area_token(b),
            b.batch_number,
        ),
    )

    ordered.extend(remaining_sorted)

    for index, batch in enumerate(ordered, start=1):
        batch.batch_number = index

    return ordered


def build_output_tables(
    scope: pd.DataFrame,
    batches: list[BatchBin],
    batch_count: int,
    tolerance: float,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, float]]:
    kks_summary = (
        scope.groupby(["Support Structure", "Working area"], as_index=False)["Total Weight [kg]"]
        .sum()
        .rename(
            columns={
                "Support Structure": "KKS code",
                "Total Weight [kg]": "KKS Total Weight [kg]",
            }
        )
    )
    kks_summary["Building section"] = kks_summary["Working area"].map(derive_building_section)

    wa_to_batch: dict[str, int] = {}
    for batch in batches:
        for chunk in batch.chunks:
            for area in chunk.working_areas:
                wa_to_batch[area] = batch.batch_number

    kks_summary["Batch number"] = kks_summary["Working area"].map(wa_to_batch)
    if kks_summary["Batch number"].isna().any():
        unresolved = kks_summary[kks_summary["Batch number"].isna()]["Working area"].unique()
        raise ValueError(f"Unassigned working areas: {sorted(unresolved)}")

    kks_summary["_section_sort"] = kks_summary["Building section"].map(section_sort_token)
    kks_summary["_wa_sort"] = kks_summary["Working area"].map(natural_sort_token)

    detail = kks_summary.sort_values(
        ["Batch number", "_section_sort", "_wa_sort", "KKS code"],
        ascending=[True, True, True, True],
    ).reset_index(drop=True)
    detail = detail.drop(columns=["_section_sort", "_wa_sort"])

    total_weight = float(detail["KKS Total Weight [kg]"].sum())
    target_weight = total_weight / batch_count
    min_weight = target_weight * (1.0 - tolerance)
    max_weight = target_weight * (1.0 + tolerance)

    summary = (
        detail.groupby("Batch number", as_index=False)
        .agg(
            Total_Weight_kg=("KKS Total Weight [kg]", "sum"),
            Building_Sections=(
                "Building section",
                lambda s: ", ".join(sorted(set(s), key=section_sort_token)),
            ),
            Working_Areas=("Working area", "nunique"),
            KKS_Count=("KKS code", "nunique"),
        )
        .sort_values("Batch number")
        .reset_index(drop=True)
    )
    summary["Target_Weight_kg"] = target_weight
    summary["Min_Allowed_kg"] = min_weight
    summary["Max_Allowed_kg"] = max_weight
    summary["Delta_vs_Target_kg"] = summary["Total_Weight_kg"] - target_weight
    summary["Within_Tolerance"] = summary["Total_Weight_kg"].between(min_weight, max_weight)

    area_summary = (
        detail.groupby(["Batch number", "Building section", "Working area"], as_index=False)[
            "KKS Total Weight [kg]"
        ]
        .sum()
        .assign(
            _section_sort=lambda d: d["Building section"].map(section_sort_token),
            _wa_sort=lambda d: d["Working area"].map(natural_sort_token),
        )
        .sort_values(["Batch number", "_section_sort", "_wa_sort"])
        .reset_index(drop=True)
        .drop(columns=["_section_sort", "_wa_sort"])
        .rename(columns={"KKS Total Weight [kg]": "Working_Area_Total_Weight_kg"})
    )

    stats = {
        "total_weight": total_weight,
        "target_weight": target_weight,
        "min_weight": min_weight,
        "max_weight": max_weight,
    }
    return detail, summary, area_summary, stats


def apply_detail_formatting(worksheet: Worksheet, detail_rows: int) -> None:
    last_data_row = detail_rows + 1
    subtotal_formula = f"=SUBTOTAL(109,C2:C{last_data_row})"

    worksheet.auto_filter.ref = f"A1:E{last_data_row}"
    worksheet.freeze_panes = "A2"

    for column, width in DETAIL_COLUMN_WIDTHS.items():
        worksheet.column_dimensions[column].width = width

    worksheet["H1"] = "Total weight"
    worksheet["H1"].style = worksheet["A1"].style

    worksheet["I1"] = subtotal_formula
    worksheet["I1"].number_format = "0.00"

    bottom_total_row = last_data_row + 3
    worksheet.cell(row=bottom_total_row, column=3, value=subtotal_formula)
    worksheet.cell(row=bottom_total_row, column=3).style = worksheet["C2"].style


def apply_summary_formatting(worksheet: Worksheet) -> None:
    for column, width in SUMMARY_COLUMN_WIDTHS.items():
        worksheet.column_dimensions[column].width = width


def write_output(
    output_file: Path,
    detail: pd.DataFrame,
    summary: pd.DataFrame,
    area_summary: pd.DataFrame,
) -> None:
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        detail.to_excel(writer, sheet_name="Detail", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False, startrow=0)
        area_summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
            startrow=len(summary) + 3,
        )

        apply_detail_formatting(
            worksheet=writer.sheets["Detail"],
            detail_rows=len(detail),
        )
        apply_summary_formatting(worksheet=writer.sheets["Summary"])


def parse_args() -> argparse.Namespace:
    script_dir = Path(__file__).resolve().parent
    default_input_path = script_dir / DEFAULT_INPUT_FILENAME

    parser = argparse.ArgumentParser(
        description="Split support scope into 5 weight-balanced batches by building section and working area."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=default_input_path,
        help="Input xlsm/xlsx path.",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET_NAME,
        help="Worksheet name to process.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output xlsx path. Defaults next to input file.",
    )
    parser.add_argument(
        "--batches",
        type=int,
        default=5,
        help="Number of batches.",
    )
    parser.add_argument(
        "--tolerance",
        type=float,
        default=0.25,
        help="Relative tolerance around target batch weight.",
    )
    parser.add_argument(
        "--section-order",
        default=",".join(DEFAULT_SECTION_ORDER),
        help=(
            "Comma-separated section priority for batch sequence and output sort, "
            "for example UHA,UMA,N/A"
        ),
    )

    args = parser.parse_args()
    if args.batches <= 0:
        raise ValueError("--batches must be a positive integer")
    if not (0.0 < args.tolerance < 1.0):
        raise ValueError("--tolerance must be between 0 and 1")
    return args


def main() -> None:
    args = parse_args()
    set_section_order(parse_section_order(args.section_order))

    input_file = args.input.resolve()
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    output_file = (
        args.output.resolve()
        if args.output is not None
        else input_file.with_name(f"{input_file.stem}_batched.xlsx")
    )

    scope = load_scope(input_file=input_file, sheet_name=args.sheet)

    wa_summary = (
        scope.assign(**{"Building section": scope["Working area"].map(derive_building_section)})
        .groupby(["Building section", "Working area"], as_index=False)["Total Weight [kg]"]
        .sum()
    )

    total_weight = float(wa_summary["Total Weight [kg]"].sum())
    target_weight = total_weight / args.batches
    max_batch_weight = target_weight * (1.0 + args.tolerance)

    if wa_summary["Total Weight [kg]"].max() > max_batch_weight:
        overweight_areas = wa_summary[wa_summary["Total Weight [kg]"] > max_batch_weight]
        area_list = overweight_areas["Working area"].tolist()
        raise ValueError(
            "At least one working area exceeds max allowed batch weight. "
            f"Cannot keep working areas whole with current settings: {area_list}"
        )

    chunks = create_chunks(wa_summary=wa_summary, max_batch_weight=max_batch_weight)
    batches = assign_chunks_to_batches(
        chunks=chunks,
        batch_count=args.batches,
        max_batch_weight=max_batch_weight,
    )
    batches = reorder_batches_for_output(batches)

    detail, summary, area_summary, stats = build_output_tables(
        scope=scope,
        batches=batches,
        batch_count=args.batches,
        tolerance=args.tolerance,
    )

    write_output(
        output_file=output_file,
        detail=detail,
        summary=summary,
        area_summary=area_summary,
    )

    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print(f"Processed rows: {len(scope)}")
    print(f"Total weight [kg]: {stats['total_weight']:.2f}")
    print(
        "Target/min/max [kg]: "
        f"{stats['target_weight']:.2f} / {stats['min_weight']:.2f} / {stats['max_weight']:.2f}"
    )
    print("Batch totals [kg]:")
    for _, row in summary.iterrows():
        print(
            f"  Batch {int(row['Batch number'])}: "
            f"{row['Total_Weight_kg']:.2f} "
            f"(within tolerance: {bool(row['Within_Tolerance'])})"
        )


if __name__ == "__main__":
    main()
