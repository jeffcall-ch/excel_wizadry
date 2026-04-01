"""
BEAM SECTION CUT OPTIMIZER — Python pseudocode
Mirrors the logic of BeamCutOptimizer_v4.bas exactly.

This is NOT production code — it is a readable translation of the VBA
so the calculation logic can be verified independently of Excel.

Input : CSV file with columns:
    [0] ID  [1] Item Number  [2] Qty  [3] Description
    [4] Cut Length [mm]  [5] Weight  [6] Total Wt  [7] Remarks  [8] Coating

Output: two lists of dicts — summary_rows and cut_plan_rows
"""

import math
import csv
import re
from collections import defaultdict


# ==============================================================================
#  CONSTANTS  (match VBA constants exactly)
# ==============================================================================

BAR_MM          = 5900    # effective packing length — conservative working assumption
PHYSICAL_BAR_MM = 6000    # actual bar length ordered from supplier
ORDER_BUFFER    = 1.05    # 5% workshop safety buffer
KERF_MM         = 3       # saw-cut allowance per cut in mm


# ==============================================================================
#  HELPERS
# ==============================================================================

def ceiling(value: float, multiple: int) -> int:
    """Round value UP to the nearest multiple of `multiple`."""
    return int(math.ceil(value / multiple) * multiple)


def profile_name(description: str) -> str:
    """
    Strip the cut-length number from a description string.
    'Beam Section MS 41/21/2,0 D x 500 HCP'  ->  'Beam Section MS 41/21/2,0 D HCP'
    """
    match = re.search(r" x ", description, re.IGNORECASE)
    if match:
        before = description[:match.start()]
        after  = description[match.end():]                 # e.g. "500 HCP"
        space  = after.find(" ")                           # find space after number
        suffix = after[space:] if space != -1 else ""      # e.g. " HCP"
        return (before + suffix).strip()
    return description.strip()


def parse_cut_length(raw: str) -> int:
    """Parse '500 mm' or '500' -> 500. Returns 0 if not parseable."""
    cleaned = raw.replace("mm", "").strip()
    try:
        return int(float(cleaned))
    except ValueError:
        return 0


# ==============================================================================
#  STEP 1 — READ AND GROUP CSV DATA
# ==============================================================================

def read_and_group(csv_path: str) -> tuple[dict, dict]:
    """
    Returns:
        group_desc   : { group_key -> profile description string }
        group_pieces : { group_key -> [cut_length, cut_length, ...] }

    group_key = "ItemNumber|Coating"
    Pieces list is already expanded by Qty
    (e.g. Qty=3, cut_length=500 -> [500, 500, 500])
    """
    group_desc   = {}
    group_pieces = defaultdict(list)

    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        next(reader)    # skip header row

        for row in reader:
            if len(row) < 9:
                continue

            item_num = row[1].strip()
            if not item_num:
                continue

            coating  = row[8].strip() or "Uncoated"
            group_key = f"{item_num}|{coating}"

            cut_len = parse_cut_length(row[4])
            if cut_len <= 0:
                continue

            try:
                qty = int(float(row[2].strip()))
            except ValueError:
                continue
            if qty <= 0:
                continue

            description = row[3].strip()
            if group_key not in group_desc:
                group_desc[group_key] = profile_name(description)

            # expand qty into individual piece lengths
            group_pieces[group_key].extend([cut_len] * qty)

    return group_desc, dict(group_pieces)


# ==============================================================================
#  STEP 2 — FFD BIN PACKING
#  Pieces MUST be sorted descending before calling.
#  Returns number of 5900mm bars needed.
# ==============================================================================

def run_ffd(pieces: list[int]) -> int:
    """
    First Fit Decreasing bin packing.
    Each bin = BAR_MM (5900 mm).
    Each piece consumes piece_length + KERF_MM from the bin.
    Returns the number of bins (bars) used.
    """
    bins = [BAR_MM]    # remaining capacity per bar

    for piece in pieces:
        piece_with_kerf = min(piece + KERF_MM, BAR_MM)   # oversized safety cap

        placed = False
        for i, remaining in enumerate(bins):
            if remaining >= piece_with_kerf:
                bins[i] -= piece_with_kerf
                placed = True
                break

        if not placed:
            bins.append(BAR_MM - piece_with_kerf)

    return len(bins)


# ==============================================================================
#  STEP 3 — REBUILD BIN LAYOUT FOR CUT PLAN
#  Same FFD logic but also records which pieces land in which bar.
# ==============================================================================

def rebuild_bin_layout(pieces: list[int]) -> list[list[int]]:
    """
    Re-runs FFD and returns a list of bars, each bar being a list of piece lengths.
    e.g. [[1200, 1100, 900], [1650, 1300], ...]
    """
    bin_remain = [BAR_MM]
    bin_pieces = [[]]

    for piece in pieces:
        piece_with_kerf = min(piece + KERF_MM, BAR_MM)

        placed = False
        for i, remaining in enumerate(bin_remain):
            if remaining >= piece_with_kerf:
                bin_remain[i] -= piece_with_kerf
                bin_pieces[i].append(piece)
                placed = True
                break

        if not placed:
            bin_remain.append(BAR_MM - piece_with_kerf)
            bin_pieces.append([piece])

    return bin_pieces


# ==============================================================================
#  STEP 4 — ORDER QUANTITY CALCULATION
#
#  FFD gives numBars using BAR_MM = 5900mm packing assumption.
#  Order quantity is based on physical 6000mm bars:
#
#    orderLenMM  = ceiling(numBars * 6000 * 1.1,  6000)
#    orderedBars = orderLenMM / 6000
#
#  Meaning: total physical length needed, buffered by 10%, rounded up
#  to a whole number of 6m bars.
# ==============================================================================

def calc_ordered_bars(num_bars: int) -> int:
    order_len_mm = ceiling(num_bars * PHYSICAL_BAR_MM * ORDER_BUFFER, PHYSICAL_BAR_MM)
    return order_len_mm // PHYSICAL_BAR_MM


# ==============================================================================
#  STEP 5 — PROCESS ALL GROUPS
# ==============================================================================

def process_groups(group_desc: dict, group_pieces: dict) -> tuple[list, list]:
    """
    Returns:
        summary_rows  : one dict per group
        cut_plan_rows : one dict per bar per group
    """
    summary_rows  = []
    cut_plan_rows = []

    for group_key in sorted(group_desc.keys()):

        pieces     = group_pieces[group_key]
        item_part, coating_part = group_key.split("|", 1)
        piece_count = len(pieces)

        if piece_count == 0:
            continue

        # --- sort largest first (FFD requirement) ---
        pieces_sorted = sorted(pieces, reverse=True)

        # --- oversize check ---
        has_oversize = pieces_sorted[0] > BAR_MM

        # --- FFD packing ---
        num_bars = run_ffd(pieces_sorted)

        # --- statistics ---
        net_len   = sum(pieces_sorted)
        kerf_loss = piece_count * KERF_MM
        used_len  = net_len + kerf_loss

        # --- order quantity ---
        ordered_bars  = calc_ordered_bars(num_bars)
        total_bar_len = ordered_bars * PHYSICAL_BAR_MM
        waste         = total_bar_len - used_len
        efficiency    = used_len / total_bar_len if total_bar_len > 0 else 0

        # --- summary row ---
        summary_rows.append({
            "item_number"   : item_part,
            "coating"       : coating_part,
            "profile"       : group_desc[group_key],
            "ordered_bars"  : ordered_bars,       # ORDER THIS MANY from supplier
            "ffd_minimum"   : num_bars,            # theoretical minimum for reference
            "total_pieces"  : piece_count,
            "net_cut_mm"    : net_len,
            "kerf_loss_mm"  : kerf_loss,
            "used_mm"       : used_len,
            "total_bar_mm"  : total_bar_len,
            "offcut_mm"     : waste,
            "efficiency_pct": round(efficiency * 100, 1),
            "warning"       : "PIECE > 5.9m" if has_oversize else "",
        })

        # --- cut plan rows (one per bar) ---
        bin_layout = rebuild_bin_layout(pieces_sorted)
        for bar_num, bar_pieces in enumerate(bin_layout, start=1):
            used_in_bar = sum(p + KERF_MM for p in bar_pieces)
            leftover    = BAR_MM - used_in_bar
            cut_plan_rows.append({
                "item_number" : item_part,
                "coating"     : coating_part,
                "bar_num"     : bar_num,
                "pieces_mm"   : bar_pieces,        # list of cut lengths in this bar
                "used_mm"     : used_in_bar,
                "leftover_mm" : leftover,
            })

        # blank separator between groups (mirrors VBA cutPlanRow += 1 spacer)
        cut_plan_rows.append({})

    return summary_rows, cut_plan_rows


# ==============================================================================
#  MAIN
# ==============================================================================

def main(csv_path: str):

    # read and group
    group_desc, group_pieces = read_and_group(csv_path)

    if not group_desc:
        print("ERROR: No valid rows found. Check CSV column layout.")
        return

    # process
    summary_rows, cut_plan_rows = process_groups(group_desc, group_pieces)

    # --- print summary ---
    print(f"\n{'='*80}")
    print(f"  BEAM CUT OPTIMIZER — SUMMARY")
    print(f"  BAR_MM={BAR_MM}  PHYSICAL={PHYSICAL_BAR_MM}  KERF={KERF_MM}  BUFFER={ORDER_BUFFER}")
    print(f"{'='*80}")
    header = f"{'Item':>10} {'Coating':<16} {'Order':>6} {'FFD':>5} {'Pieces':>7} {'Eff%':>6}  Profile"
    print(header)
    print("-" * 80)
    for row in summary_rows:
        print(
            f"{row['item_number']:>10} "
            f"{row['coating']:<16} "
            f"{row['ordered_bars']:>6} "
            f"{row['ffd_minimum']:>5} "
            f"{row['total_pieces']:>7} "
            f"{row['efficiency_pct']:>5.1f}%  "
            f"{row['profile']}"
            + (f"  *** {row['warning']}" if row['warning'] else "")
        )

    # --- print cut plan (abbreviated) ---
    print(f"\n{'='*80}")
    print("  CUT PLAN")
    print(f"{'='*80}")
    for row in cut_plan_rows:
        if not row:
            print()
            continue
        pieces_str = "  |  ".join(str(p) for p in row["pieces_mm"])
        print(
            f"{row['item_number']:>10} {row['coating']:<16} "
            f"Bar {row['bar_num']:>3}:  {pieces_str}  "
            f"(used={row['used_mm']} mm  leftover={row['leftover_mm']} mm)"
        )

    print(f"\n{len(summary_rows)} groups processed.")


if __name__ == "__main__":
    import sys
    csv_file = sys.argv[1] if len(sys.argv) > 1 else "beams_sections.csv"
    main(csv_file)
