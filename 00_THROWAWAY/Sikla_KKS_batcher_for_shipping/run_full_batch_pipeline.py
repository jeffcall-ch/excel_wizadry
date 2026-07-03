from __future__ import annotations

import argparse
import configparser
import subprocess
import sys
from pathlib import Path

import pandas as pd


DEFAULT_CONFIG_FILENAME = "batch_pipeline.ini"


def normalize_kks(value: object) -> str | None:
    if pd.isna(value):
        return None

    text = str(value).strip().upper().replace(" ", "")
    if not text or text == "NAN":
        return None

    if "/SU" in text and not text.startswith("/"):
        text = f"/{text}"

    return text


def parse_bool(value: str, default: bool = False) -> bool:
    text = str(value).strip().lower()
    if not text:
        return default
    return text in {"1", "true", "yes", "y", "on"}


def resolve_path(base_dir: Path, value: str) -> Path:
    candidate = Path(str(value).strip())
    if not candidate.is_absolute():
        candidate = base_dir / candidate
    return candidate.resolve()


def run_step(name: str, command: list[str], cwd: Path, dry_run: bool) -> None:
    print(f"\n=== {name} ===")
    print("COMMAND:")
    print(" ".join(command))

    if dry_run:
        print("Dry run enabled, step not executed.")
        return

    result = subprocess.run(
        command,
        cwd=str(cwd),
        capture_output=True,
        text=True,
        check=False,
    )

    if result.stdout.strip():
        print("STDOUT:")
        print(result.stdout.strip())
    if result.stderr.strip():
        print("STDERR:")
        print(result.stderr.strip())

    if result.returncode != 0:
        raise RuntimeError(f"Step failed: {name} (exit code {result.returncode})")


def read_config(config_file: Path) -> configparser.ConfigParser:
    cfg = configparser.ConfigParser(interpolation=None)
    if not cfg.read(config_file, encoding="utf-8"):
        raise FileNotFoundError(f"Could not read config file: {config_file}")
    return cfg


def require_path_exists(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")


def make_reconciliation_report(
    status_file: Path,
    status_sheet: str,
    c5_file: Path,
    final_batch_file: Path,
    final_detail_sheet: str,
    deduct_deleted: bool,
    deduct_c5: bool,
    deduct_solid: bool,
    deduct_bq_not_exist: bool,
    output_file: Path,
) -> dict[str, int | bool]:
    status_df = pd.read_excel(status_file, sheet_name=status_sheet, engine="openpyxl")

    status_col = None
    for col in status_df.columns:
        label = str(col).replace("\n", " ").strip().lower()
        if "support" in label and "drawing" in label and "number" in label:
            status_col = col
            break
    if not status_col:
        raise ValueError("Could not find Support Drawing Number column in status file")

    status_df["KKS"] = status_df[status_col].map(normalize_kks)
    status_df = status_df[status_df["KKS"].notna()].copy()
    status_df = status_df[status_df["KKS"].str.contains("/SU", na=False)].copy()

    status_set = set(status_df["KKS"])

    text_frame = status_df.astype(str).apply(lambda col: col.str.lower())
    deleted_set = set(
        status_df.loc[
            text_frame.apply(
                lambda col: col.str.contains("deleted|withdraw", na=False, regex=True)
            ).any(axis=1),
            "KKS",
        ]
    )
    solid_set = set(
        status_df.loc[
            text_frame.apply(
                lambda col: col.str.contains("solid handling support not in gp lot", na=False)
            ).any(axis=1),
            "KKS",
        ]
    )
    bq_not_exist_set = set(
        status_df.loc[
            text_frame.apply(lambda col: col.str.contains("bq not exist", na=False)).any(axis=1),
            "KKS",
        ]
    )

    c5_df = pd.read_excel(c5_file, sheet_name=0, engine="openpyxl")
    c5_col = "KKS C5" if "KKS C5" in c5_df.columns else c5_df.columns[0]
    c5_set = set(
        k for k in (normalize_kks(value) for value in c5_df[c5_col]) if k and "/SU" in k
    )

    final_df = pd.read_excel(final_batch_file, sheet_name=final_detail_sheet, engine="openpyxl")
    if "KKS code" not in final_df.columns:
        raise ValueError(f"Could not find KKS code column in final detail sheet: {final_detail_sheet}")
    final_set = set(
        k for k in (normalize_kks(value) for value in final_df["KKS code"]) if k and "/SU" in k
    )

    deduction_parts: list[set[str]] = []
    if deduct_deleted:
        deduction_parts.append(deleted_set)
    if deduct_c5:
        deduction_parts.append(c5_set)
    if deduct_solid:
        deduction_parts.append(solid_set)
    if deduct_bq_not_exist:
        deduction_parts.append(bq_not_exist_set)

    if deduction_parts:
        deduction_union = set().union(*deduction_parts)
    else:
        deduction_union = set()

    deduction_union = deduction_union & status_set
    expected_set = status_set - deduction_union

    expected_minus_final = sorted(expected_set - final_set)
    final_minus_expected = sorted(final_set - expected_set)

    summary = {
        "status_total": len(status_set),
        "final_total": len(final_set),
        "deleted_total": len(deleted_set & status_set),
        "c5_total": len(c5_set & status_set),
        "solid_total": len(solid_set & status_set),
        "bq_not_exist_total": len(bq_not_exist_set & status_set),
        "deduct_union_total": len(deduction_union),
        "expected_after_deduction": len(expected_set),
        "expected_minus_final": len(expected_minus_final),
        "final_minus_expected": len(final_minus_expected),
        "sets_match": expected_set == final_set,
    }

    summary_rows = [
        ("Status unique /SU", summary["status_total"]),
        ("Final unique /SU", summary["final_total"]),
        ("Deleted in status", summary["deleted_total"]),
        ("C5", summary["c5_total"]),
        ("Solid handling not in GP", summary["solid_total"]),
        ("BQ not exist", summary["bq_not_exist_total"]),
        ("Deduction union", summary["deduct_union_total"]),
        ("Expected after deduction", summary["expected_after_deduction"]),
        ("Expected minus final", summary["expected_minus_final"]),
        ("Final minus expected", summary["final_minus_expected"]),
        ("Set match", summary["sets_match"]),
    ]

    summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])
    expected_minus_df = pd.DataFrame({"KKS missing in final": expected_minus_final})
    final_minus_df = pd.DataFrame({"KKS extra in final": final_minus_expected})

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        expected_minus_df.to_excel(writer, sheet_name="Expected_minus_Final", index=False)
        final_minus_df.to_excel(writer, sheet_name="Final_minus_Expected", index=False)

    return summary


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Run full Sikla KKS batching pipeline: split, reconcile missing, "
            "integrate missing, then reconcile final counts."
        )
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=Path(__file__).resolve().parent / DEFAULT_CONFIG_FILENAME,
        help="Path to INI config file.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print commands only, do not execute.",
    )
    args = parser.parse_args()

    config_file = args.config.resolve()
    cfg = read_config(config_file)

    script_dir = Path(__file__).resolve().parent
    base_dir = resolve_path(script_dir, cfg.get("paths", "base_dir", fallback="."))

    source_bom_file = resolve_path(base_dir, cfg.get("paths", "source_bom_file"))
    source_bom_sheet = cfg.get("paths", "source_bom_sheet", fallback="Per_Support_Structure")

    status_file = resolve_path(base_dir, cfg.get("paths", "status_file"))
    status_sheet = cfg.get("paths", "status_sheet", fallback="Pipe support status list")

    c5_file = resolve_path(base_dir, cfg.get("paths", "c5_file"))
    original_data_file = resolve_path(base_dir, cfg.get("paths", "original_data_file"))

    batched_output = resolve_path(base_dir, cfg.get("paths", "step1_batched_output"))
    missing_report_output = resolve_path(base_dir, cfg.get("paths", "step2_missing_report_output"))
    final_output = resolve_path(base_dir, cfg.get("paths", "step3_final_output"))
    reconciliation_output = resolve_path(base_dir, cfg.get("paths", "step4_reconciliation_output"))

    batch_count = cfg.getint("batching", "batch_count", fallback=5)
    tolerance = cfg.getfloat("batching", "tolerance", fallback=0.25)
    section_order = cfg.get("batching", "section_order", fallback="UHA,UMA,N/A")

    exclude_deleted = parse_bool(cfg.get("integration", "exclude_deleted", fallback="true"))
    exclude_bq_not_exist = parse_bool(
        cfg.get("integration", "exclude_bq_not_exist", fallback="false")
    )
    exclude_solid_not_gp = parse_bool(
        cfg.get("integration", "exclude_solid_not_gp", fallback="true")
    )
    na_working_areas = cfg.get(
        "integration", "na_working_areas", fallback="NO_AREA,BQ NOT EXIST"
    )
    na_section_label = cfg.get("integration", "na_section_label", fallback="N/A")
    merge_unknown_to_na = parse_bool(
        cfg.get("integration", "merge_unknown_to_na", fallback="true")
    )

    deduct_deleted = parse_bool(cfg.get("reconciliation", "deduct_deleted", fallback="true"))
    deduct_c5 = parse_bool(cfg.get("reconciliation", "deduct_c5", fallback="true"))
    deduct_solid = parse_bool(cfg.get("reconciliation", "deduct_solid_not_gp", fallback="true"))
    deduct_bq_not_exist = parse_bool(
        cfg.get("reconciliation", "deduct_bq_not_exist", fallback="false")
    )

    for path_obj, label in [
        (source_bom_file, "Source BoM file"),
        (status_file, "Status file"),
        (c5_file, "C5 file"),
        (original_data_file, "Original data file"),
    ]:
        require_path_exists(path_obj, label)

    run_step(
        name="Step 1 - Split Base Scope Into Batches",
        command=[
            sys.executable,
            str(script_dir / "split_shipping_batches.py"),
            "--input",
            str(source_bom_file),
            "--sheet",
            source_bom_sheet,
            "--output",
            str(batched_output),
            "--batches",
            str(batch_count),
            "--tolerance",
            str(tolerance),
            "--section-order",
            section_order,
        ],
        cwd=base_dir,
        dry_run=args.dry_run,
    )

    run_step(
        name="Step 2 - Build Missing KKS Report",
        command=[
            sys.executable,
            str(script_dir / "reconcile_missing_kks.py"),
            "--status-file",
            str(status_file),
            "--status-sheet",
            status_sheet,
            "--batched-file",
            str(batched_output),
            "--batched-sheet",
            "Detail",
            "--c5-file",
            str(c5_file),
            "--output",
            str(missing_report_output),
        ],
        cwd=base_dir,
        dry_run=args.dry_run,
    )

    run_step(
        name="Step 3 - Integrate In-Scope Missing Supports",
        command=[
            sys.executable,
            str(script_dir / "integrate_missing_into_batches.py"),
            "--missing-report",
            str(missing_report_output),
            "--batched-file",
            str(batched_output),
            "--original-file",
            str(original_data_file),
            "--output",
            str(final_output),
            "--section-order",
            section_order,
            "--na-working-areas",
            na_working_areas,
            "--na-section-label",
            na_section_label,
            "--merge-unknown-to-na",
            str(merge_unknown_to_na).lower(),
            "--exclude-deleted",
            str(exclude_deleted).lower(),
            "--exclude-bq-not-exist",
            str(exclude_bq_not_exist).lower(),
            "--exclude-solid-not-gp",
            str(exclude_solid_not_gp).lower(),
        ],
        cwd=base_dir,
        dry_run=args.dry_run,
    )

    if args.dry_run:
        print("\nDry run finished. No files were changed.")
        return

    final_detail_sheet = "Detail_Updated"
    summary = make_reconciliation_report(
        status_file=status_file,
        status_sheet=status_sheet,
        c5_file=c5_file,
        final_batch_file=final_output,
        final_detail_sheet=final_detail_sheet,
        deduct_deleted=deduct_deleted,
        deduct_c5=deduct_c5,
        deduct_solid=deduct_solid,
        deduct_bq_not_exist=deduct_bq_not_exist,
        output_file=reconciliation_output,
    )

    print("\n=== Step 4 - Final Reconciliation ===")
    print(f"Reconciliation report: {reconciliation_output}")
    for key, value in summary.items():
        print(f"{key}: {value}")


if __name__ == "__main__":
    main()
