"""
Generate a fully-populated test workbook with random scores.

  • All group-stage goal cells → =RANDBETWEEN(0,4)
  • All bracket goal cells    → =RANDBETWEEN(0,4)
  • All bracket penalty cells → =RANDBETWEEN(3,7)  (slightly wider range to
    keep ties rare but still exercise the penalty path)
  • The 8 best-3rd-place team slots → pre-filled with plausible team names

Run:
    poetry run python generate_randbetween.py

The output file is written next to this script as:
    quiniela_mundial_2026_randbetween.xlsx
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from src.data import GROUPS
from src.sheets.bracket import build_bracket
from src.sheets.clasificados import build_clasificados
from src.sheets.group_stage import build_group_stage
from src.sheets.references import build_references
from src.styles import MATCH_BG, WHITE, fill

OUTPUT_FILE = Path("quiniela_mundial_2026_randbetween.xlsx")

# ── 8 plausible 3rd-place qualifiers (one per bracket slot, in order) ──────
# The slots appear in the order they are drawn: R32-Left slots 2,5,6,7 then
# R32-Right slots 0,1,4,7  (all "3rd*" entries in data.py).
THIRD_PLACE_TEAMS = [
    "Czech Republic",   # R32-L slot 2  (vs E1 – Germany)
    "Norway",           # R32-L slot 5  (vs I1 – France)
    "South Africa",     # R32-L slot 6  (vs A1 – Mexico)
    "Ghana",            # R32-L slot 7  (vs L1 – England)
    "New Zealand",      # R32-R slot 0  (vs G1 – Belgium)
    "Turkey",           # R32-R slot 1  (vs D1 – USA)
    "Qatar",            # R32-R slot 4  (vs B1 – Canada)
    "DR Congo",         # R32-R slot 7  (vs K1 – Portugal)
]

def _is_input_yellow(cell) -> bool:
    """True for cells styled with INPUT_YELLOW (FFD966) fill."""
    try:
        return (
            cell.fill.fill_type == "solid"
            and "FFD966" in cell.fill.fgColor.rgb
        )
    except Exception:
        return False


def _fill_group_scores(ws) -> None:
    """Replace every blank INPUT_YELLOW cell in the group-stage sheet with RANDBETWEEN(0,4).
    Skips label cells (e.g. the legend row) that share the same fill but already have text.
    """
    for row in ws.iter_rows():
        for cell in row:
            if _is_input_yellow(cell) and cell.value is None:
                cell.value = "=RANDBETWEEN(0,4)"


def _fill_bracket_scores(ws, bracket_info: dict) -> None:
    """Fill bracket goal cells and penalty cells with random formulas."""
    # Regular-time goal cells (INPUT_YELLOW)
    for row in ws.iter_rows():
        for cell in row:
            if _is_input_yellow(cell):
                cell.value = "=RANDBETWEEN(0,4)"

    # Penalty score cells (tracked by build_bracket)
    for coord in bracket_info["pen_cells"]:
        ws[coord].value = "=RANDBETWEEN(3,7)"


def _fill_third_place_cells(ws, bracket_info: dict) -> None:
    """Write team names into the 8 manual 3rd-place qualifier slots."""
    coords = bracket_info["third_place_cells"]
    if len(coords) != len(THIRD_PLACE_TEAMS):
        raise RuntimeError(
            f"Expected {len(THIRD_PLACE_TEAMS)} third-place cells, "
            f"found {len(coords)}: {coords}"
        )
    for coord, team in zip(coords, THIRD_PLACE_TEAMS):
        cell = ws[coord]
        cell.value = team
        # Style as a normal team cell (not the placeholder grey)
        cell.fill = fill(MATCH_BG)
        cell.font = Font(name="Calibri", size=9, bold=True, color=WHITE)


def main() -> None:
    wb = Workbook()
    del wb["Sheet"]

    print("Building group stage tab…")
    standings_start = build_group_stage(wb)

    print("Building Clasificados tab…")
    qualified_refs = build_clasificados(wb, standings_start)

    print("Building bracket tab…")
    bracket_info = build_bracket(wb, qualified_refs)

    print("Building references tab…")
    build_references(wb)

    print("Filling random scores (group stage)…")
    _fill_group_scores(wb["Fase de Grupos"])

    print("Filling random scores (bracket goal cells)…")
    _fill_bracket_scores(wb["Bracket"], bracket_info)

    print("Filling 3rd-place team names…")
    _fill_third_place_cells(wb["Bracket"], bracket_info)

    wb.active = wb["Fase de Grupos"]
    wb.save(OUTPUT_FILE)
    print(f"✅  Saved → {OUTPUT_FILE}")
    print(f"   3rd-place cells filled: {bracket_info['third_place_cells']}")
    print(f"   Penalty cell count:     {len(bracket_info['pen_cells'])}")


if __name__ == "__main__":
    main()
