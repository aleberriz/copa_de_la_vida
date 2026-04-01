"""
FIFA World Cup 2026 — Quiniela generator.
Orchestrates all worksheet builders and saves the workbook.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from src.sheets.bracket import build_bracket
from src.sheets.clasificados import build_clasificados
from src.sheets.group_stage import build_group_stage
from src.sheets.references import build_references

OUTPUT_FILE = Path("quiniela_mundial_2026.xlsx")


def main() -> None:
    wb = Workbook()
    del wb["Sheet"]

    print("Building group stage tab…")
    standings_start = build_group_stage(wb)

    print("Building Clasificados tab…")
    qualified_refs = build_clasificados(wb, standings_start)

    print("Building bracket tab…")
    build_bracket(wb, qualified_refs)

    print("Building references tab…")
    build_references(wb)

    wb.active = wb["Fase de Grupos"]
    wb.save(OUTPUT_FILE)
    print(f"✅  Saved → {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
