"""
exporter.py — Excel workbook helper

Thin wrapper used by change_detector and enricher to write
styled rows into the master workbook without duplicating
formatting logic. Also exposes a standalone re-export function
that rewrites the Current Snapshot tab cleanly from the history.
"""

from __future__ import annotations

from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

import config

EXCEL_PATH = Path(config.EXCEL_PATH)

LINKEDIN_BLUE = "0A66C2"
HEADER_FILL   = PatternFill("solid", fgColor=LINKEDIN_BLUE)
HEADER_FONT   = Font(color="FFFFFF", bold=True, size=10)
ALT_FILL      = PatternFill("solid", fgColor="F3F6FA")
GREEN_FILL    = PatternFill("solid", fgColor="E4F3EC")


def style_header_row(ws):
    """Apply LinkedIn-blue header styling to row 1 of a worksheet."""
    for cell in ws[1]:
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def write_row(ws, row_num: int, headers: list[str], data: dict):
    """Write a data dict into the given row, applying alternating fill."""
    fill = ALT_FILL if row_num % 2 == 0 else None
    for i, col in enumerate(headers, 1):
        cell = ws.cell(row=row_num, column=i, value=data.get(col))
        if fill:
            cell.fill = fill


def append_row(ws, headers: list[str], data: dict):
    """Append a new row at the bottom of the sheet."""
    write_row(ws, ws.max_row + 1, headers, data)


def get_header_map(ws) -> dict[str, int]:
    """Return {column_name: 1-based column index} from row 1."""
    return {cell.value: cell.column for cell in ws[1] if cell.value}


def ensure_column(ws, col_name: str) -> int:
    """
    Ensure a column with col_name exists in row 1.
    Adds it at the end if missing. Returns its 1-based index.
    """
    hmap = get_header_map(ws)
    if col_name in hmap:
        return hmap[col_name]
    new_idx = ws.max_column + 1
    cell = ws.cell(row=1, column=new_idx, value=col_name)
    cell.font      = HEADER_FONT
    cell.fill      = HEADER_FILL
    cell.alignment = Alignment(horizontal="center")
    return new_idx


def auto_width(ws, min_width: int = 10, max_width: int = 40):
    """Auto-size column widths based on content."""
    for col in ws.columns:
        length = max(
            len(str(cell.value or "")) for cell in col
        )
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(min_width, min(max_width, length + 2))
