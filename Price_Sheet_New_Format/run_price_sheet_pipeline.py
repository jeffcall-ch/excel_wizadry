from __future__ import annotations

import argparse
import logging
from datetime import datetime
from pathlib import Path

from logging_utils import DEFAULT_LOG_DIR, setup_csv_logging
from price_sheet_import import load_import_settings, run_import
from sqlite_export_tables_to_xlsm import export_sqlite_to_xlsm


logger = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run Price Sheet import+spare+export pipeline (test mode uses timestamped DB files).",
    )
    parser.add_argument(
        "--test",
        action="store_true",
        help="Use a timestamped test DB in the test folder.",
    )
    parser.add_argument(
        "--revision-folder",
        type=Path,
        default=Path("Price_Sheet_New_Format/rev0"),
        help="Path to revision folder (contains input subfolder).",
    )
    parser.add_argument(
        "--input-subdir",
        type=str,
        default="input",
        help="Input subfolder name under revision folder.",
    )
    parser.add_argument(
        "--production-db-path",
        type=Path,
        default=Path("Price_Sheet_New_Format/price_sheet.db"),
        help="Production SQLite DB path (used when --test is not set).",
    )
    parser.add_argument(
        "--test-db-path",
        type=Path,
        default=Path("Price_Sheet_New_Format/price_sheet_test.db"),
        help="Legacy argument. Ignored when --test is set.",
    )
    parser.add_argument(
        "--test-db-dir",
        type=Path,
        default=Path("Price_Sheet_New_Format/test_runs"),
        help="Folder where timestamped test SQLite DB files are created when --test is set.",
    )
    parser.add_argument(
        "--test-db-prefix",
        type=str,
        default="price_sheet_test",
        help="Filename prefix for timestamped test DB files.",
    )
    parser.add_argument(
        "--headers-workbook",
        type=Path,
        default=Path("Price_Sheet_New_Format/sqlite_table_headers.xlsx"),
        help="Workbook with required SQLite column headers.",
    )
    parser.add_argument(
        "--settings-workbook",
        type=Path,
        default=Path("Price_Sheet_New_Format/import_settings.xlsx"),
        help="Workbook with import and spare settings.",
    )
    parser.add_argument(
        "--export-path",
        type=Path,
        default=Path("Price_Sheet_New_Format/price_sheet_export.xlsx"),
        help="Export XLSX path. This file is always overwritten.",
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
    log_file = setup_csv_logging(run_name="run_price_sheet_pipeline", log_dir=args.log_dir)
    logger.info("CSV logging initialized. log_file=%s", log_file)

    if args.test:
        args.test_db_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        db_path = args.test_db_dir / f"{args.test_db_prefix}_{ts}.db"
        logger.info("Test mode database path=%s", db_path)
    else:
        db_path = args.production_db_path
        logger.info("Production mode database path=%s", db_path)

    try:
        load_import_settings(args.settings_workbook)
    except Exception as exc:
        logger.exception("Failed to load settings workbook path=%s", args.settings_workbook)
        return 1

    import_result = run_import(
        revision_folder=args.revision_folder,
        db_path=db_path,
        headers_workbook=args.headers_workbook,
        input_subdir=args.input_subdir,
        settings_workbook=args.settings_workbook,
    )
    if int(import_result) != 0:
        logger.error("Import phase failed. Export was not started.")
        return 1

    export_result = export_sqlite_to_xlsm(
        db_path=db_path,
        output_path=args.export_path,
        overwrite=True,
    )
    if int(export_result) != 0:
        logger.error("Export phase failed.")
        return 1

    if args.test:
        logger.info("Test pipeline completed (import + spare + export).")
        logger.info("Test DB retained at=%s", db_path)
    else:
        logger.info("Production pipeline completed (import + spare + export).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
