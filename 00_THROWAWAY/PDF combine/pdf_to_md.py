import sys
import io
import warnings
import fitz  # PyMuPDF
from pathlib import Path

# Suppress the "Consider using pymupdf_layout" advisory from PyMuPDF
warnings.filterwarnings("ignore")

WORD_LIMIT = 450_000
SIZE_LIMIT = 180 * 1024 * 1024  # 180 MB in bytes

# Max vertical gap (px) between two tables to merge them into one logical table
TABLE_MERGE_GAP = 50
# Max vertical gap (px) to look for adjacent rows above/below a table bbox
TABLE_ADJ_GAP = 35


# ---------------------------------------------------------------------------
# Table helpers
# ---------------------------------------------------------------------------

def rows_to_md(rows: list) -> str:
    """Convert a list-of-lists to a Markdown table string."""
    if not rows:
        return ""
    col_count = max(len(r) for r in rows)
    if col_count == 0:
        return ""

    def fmt(row):
        padded = (list(row) + [""] * col_count)[:col_count]
        return "| " + " | ".join(str(c) for c in padded) + " |"

    lines = [fmt(rows[0]), "| " + " | ".join(["---"] * col_count) + " |"]
    for row in rows[1:]:
        lines.append(fmt(row))
    return "\n".join(lines)


def _group_cells_by_row(cells) -> dict:
    """Group table.cells into {round(y0): [cell_bbox, ...]}."""
    rows: dict = {}
    for c in cells:
        if c is None:
            continue
        k = round(c[1])
        rows.setdefault(k, []).append(c)
    return rows


def _is_full_width(cell, tbbox, tol: float = 5.0) -> bool:
    """True if the cell spans the full table width (a merged single-cell row)."""
    return abs(cell[0] - tbbox[0]) < tol and abs(cell[2] - tbbox[2]) < tol


def extract_table_rows(page: fitz.Page, table) -> list:
    """
    Extract all rows of a PyMuPDF Table object.

    For rows where PyMuPDF merged all cells into one (shaded/alternating rows),
    re-extracts the text via bbox clipping and splits on newlines to recover
    individual column values.
    """
    cc = table.col_count
    if cc == 0:
        return []

    tbbox = table.bbox
    extracted = table.extract()  # best-effort from PyMuPDF
    row_map = _group_cells_by_row(table.cells)
    result = []

    for ri, y_key in enumerate(sorted(row_map)):
        cells_in_row = sorted(row_map[y_key], key=lambda c: c[0])

        if len(cells_in_row) == 1 and _is_full_width(cells_in_row[0], tbbox):
            # ---- Merged row: clip the full-width cell and split by \n ----
            raw = page.get_text("text", clip=fitz.Rect(cells_in_row[0])).strip()
            parts = [p.strip() for p in raw.split("\n") if p.strip()]
            row = (parts + [""] * cc)[:cc]
        elif ri < len(extracted):
            ext = extracted[ri]
            good = sum(1 for v in ext if v)
            if good >= cc // 2:
                # PyMuPDF extracted this row cleanly
                row = [" ".join(str(v).split()) if v else "" for v in ext]
            else:
                # Fall back to per-cell clipping
                row = []
                for c in cells_in_row:
                    t = page.get_text("text", clip=fitz.Rect(c)).strip()
                    row.append(" ".join(t.split()))
                row = (row + [""] * cc)[:cc]
        else:
            row = [""] * cc

        result.append(row)

    return result


def adjacent_table_rows(page: fitz.Page, tbbox, col_count: int, text_blocks) -> tuple:
    """
    Find text blocks immediately above / below `tbbox` that look like table
    rows (have between col_count//2 and col_count+3 newline-separated values,
    and are horizontally within the table bounds).

    Returns (above_rows, below_rows) where each is a sorted list of
    (y0, [cell_values]).
    """
    tl, tt, tr, tb = tbbox
    lo, hi = max(1, col_count // 2), col_count + 3
    above, below = [], []

    for block in text_blocks:
        x0, y0, x1, y1, text, _, btype = block
        if btype != 0 or not text.strip():
            continue
        if x0 < tl - 10 or x1 > tr + 10:
            continue

        parts = [p.strip() for p in text.split("\n") if p.strip()]
        if not (lo <= len(parts) <= hi):
            continue

        vals = (parts + [""] * col_count)[:col_count]

        # Skip rows that have no numeric content – those are likely header
        # continuation lines, not data rows.
        if not any(any(c.isdigit() for c in v) for v in vals):
            continue

        if 0 < (tt - y1) <= TABLE_ADJ_GAP:
            above.append((y0, vals))
        elif 0 < (y0 - tb) <= TABLE_ADJ_GAP:
            below.append((y0, vals))

    return sorted(above), sorted(below)


# ---------------------------------------------------------------------------
# Page content extraction
# ---------------------------------------------------------------------------

def bboxes_overlap(a: tuple, b: tuple, tolerance: float = 2.0) -> bool:
    """Return True if rect a overlaps rect b (both as (x0, y0, x1, y1))."""
    return not (
        a[2] <= b[0] + tolerance
        or a[0] >= b[2] - tolerance
        or a[3] <= b[1] + tolerance
        or a[1] >= b[3] - tolerance
    )


def extract_page_content(page: fitz.Page) -> str:
    """
    Extract text and tables from a page in reading order.

    - Adjacent tables with the same column count are merged into one logical
      table so that "gap rows" (rows between split detections) are included.
    - Merged/shaded rows are recovered via bbox clipping + newline splitting.
    - Footer rows (Average, S&P 500, etc.) just below a table are included.
    - The PyMuPDF layout-analysis warning is suppressed.
    """
    # Suppress the "Consider using pymupdf_layout" stderr message
    _saved_stderr = sys.stderr
    sys.stderr = io.StringIO()
    try:
        tab_finder = page.find_tables()
    finally:
        sys.stderr = _saved_stderr

    tables = sorted(tab_finder.tables, key=lambda t: t.bbox[1])
    text_blocks = list(page.get_text("blocks"))

    # ---- Group adjacent same-column-count tables into logical groups ----
    groups: list = []
    if tables:
        cur = [tables[0]]
        for t in tables[1:]:
            gap = t.bbox[1] - cur[-1].bbox[3]
            if t.col_count == cur[-1].col_count and gap <= TABLE_MERGE_GAP:
                cur.append(t)
            else:
                groups.append(cur)
                cur = [t]
        groups.append(cur)

    # ---- Build MD table for each logical group ----
    claimed_y0: set = set()   # y0 keys of text blocks consumed by tables
    table_items: list = []

    for grp in groups:
        col_count = grp[0].col_count
        all_rows: list = []
        group_y0 = grp[0].bbox[1]

        for gi, tbl in enumerate(grp):
            if gi > 0:
                # Gap rows between this table and the previous one in the group
                above, _ = adjacent_table_rows(page, tbl.bbox, col_count, text_blocks)
                for y0_r, vals in above:
                    claimed_y0.add(round(y0_r))
                    all_rows.append(vals)

            all_rows.extend(extract_table_rows(page, tbl))

        # Footer rows below the last table in the group
        _, footer = adjacent_table_rows(page, grp[-1].bbox, col_count, text_blocks)
        for y0_r, vals in footer:
            claimed_y0.add(round(y0_r))
            all_rows.append(vals)

        md = rows_to_md(all_rows)
        if md:
            table_items.append((group_y0, md))

    # ---- Collect non-table plain text ----
    all_tbboxes = [t.bbox for t in tables]
    items: list = list(table_items)

    for block in text_blocks:
        x0, y0, x1, y1, text, _, btype = block
        if btype != 0 or not text.strip():
            continue
        if any(bboxes_overlap((x0, y0, x1, y1), tb) for tb in all_tbboxes):
            continue
        if round(y0) in claimed_y0:
            continue
        items.append((y0, text.strip()))

    items.sort(key=lambda item: item[0])
    return "\n\n".join(content for _, content in items)


class OutputManager:
    """Manages one or more output .md files, splitting when limits are reached."""

    def __init__(self, base_path: Path):
        self.base_path = base_path
        self.part = 1
        self.word_count = 0
        self._open_new_file()

    def _output_path(self) -> Path:
        return self.base_path.parent / f"{self.base_path.stem}_part{self.part}.md"

    def _open_new_file(self):
        self.current_path = self._output_path()
        self.current_file = open(self.current_path, "w", encoding="utf-8")
        print(f"Writing to: {self.current_path.name}")

    def _close_current_file(self):
        if not self.current_file.closed:
            self.current_file.flush()
            size = self.current_path.stat().st_size
            print(
                f"  Closed {self.current_path.name} "
                f"({self.word_count:,} words, {size / 1024 / 1024:.1f} MB)"
            )
            self.current_file.close()

    def write(self, text: str):
        self.current_file.write(text)
        self.current_file.flush()
        self.word_count += len(text.split())

        size = self.current_path.stat().st_size
        if self.word_count >= WORD_LIMIT or size >= SIZE_LIMIT:
            self._close_current_file()
            self.part += 1
            self.word_count = 0
            self._open_new_file()

    def close(self):
        self._close_current_file()

    @property
    def total_parts(self) -> int:
        return self.part


def main(pdf_path: str):
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        print(f"ERROR: File not found: {pdf_path}")
        sys.exit(1)

    print(f"Converting: {pdf_path.name}")
    out = OutputManager(pdf_path)

    try:
        with fitz.open(str(pdf_path)) as doc:
            total_pages = len(doc)
            for page_num, page in enumerate(doc, 1):
                print(f"  Processing page {page_num}/{total_pages}...", end="\r")
                content = extract_page_content(page)
                if content.strip():
                    out.write(
                        f"\n\n---\n*Page {page_num} of {total_pages}*\n\n{content}\n"
                    )
    finally:
        out.close()

    print(f"\nDone. {out.total_parts} output file(s) written.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pdf_to_md.py <path_to_pdf>")
        sys.exit(1)
    main(sys.argv[1])
