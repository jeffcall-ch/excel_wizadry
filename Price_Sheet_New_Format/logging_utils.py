from __future__ import annotations

import csv
import io
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

DEFAULT_LOG_DIR = Path("Price_Sheet_New_Format/logs")


class CsvLogFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        timestamp = datetime.fromtimestamp(record.created).isoformat(timespec="milliseconds")
        message = record.getMessage()
        exc_text = ""
        if record.exc_info:
            exc_text = self.formatException(record.exc_info)

        row = [
            timestamp,
            record.levelname,
            record.name,
            record.module,
            record.funcName,
            str(record.lineno),
            message,
            exc_text,
        ]

        buffer = io.StringIO()
        writer = csv.writer(buffer)
        writer.writerow(row)
        return buffer.getvalue().rstrip("\r\n")


def _find_existing_csv_handler(root_logger: logging.Logger) -> Optional[logging.Handler]:
    for handler in root_logger.handlers:
        if getattr(handler, "_is_csv_run_handler", False):
            return handler
    return None


def setup_csv_logging(
    run_name: str,
    log_dir: Path = DEFAULT_LOG_DIR,
    level: int = logging.INFO,
) -> Path:
    root_logger = logging.getLogger()
    existing_csv_handler = _find_existing_csv_handler(root_logger)
    if existing_csv_handler is not None:
        return Path(getattr(existing_csv_handler, "_log_file_path"))

    log_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"{run_name}_{ts}.csv"

    with log_file.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["timestamp", "level", "logger", "module", "function", "line", "message", "exception"])

    root_logger.setLevel(level)

    csv_handler = logging.FileHandler(log_file, encoding="utf-8")
    csv_handler.setFormatter(CsvLogFormatter())
    csv_handler._is_csv_run_handler = True  # type: ignore[attr-defined]
    csv_handler._log_file_path = str(log_file)  # type: ignore[attr-defined]
    root_logger.addHandler(csv_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(
        logging.Formatter(
            fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    console_handler._is_csv_console_handler = True  # type: ignore[attr-defined]
    root_logger.addHandler(console_handler)

    return log_file
