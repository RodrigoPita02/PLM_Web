"""Pós-processamento aditivo dos relatórios Excel: navegação, impressão e legibilidade.

Não altera o conteúdo nem o estilo já aplicado pelos engines; só acrescenta
(filtros, gridlines off, data bars, formato de números, configuração de impressão
e um guia de folhas com hiperlinks no Executive Summary, quando existe).
"""
from __future__ import annotations

import math

from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SUMMARY_SHEETS = {"Executive Summary", "Adjusted Summary"}

SHEET_GUIDE = {
    "Adjusted Summary": "Error counts per field and materials without an active view",
    "Field With Error": "One row per material/field error to be corrected",
    "All Errors": "Complete error list, including missing views",
    "View is Missing": "Materials missing required views",
    "Original Data": "Source data used for this validation run",
    "Field Diagnostics": "Input diagnostics for the validated fields",
}

_BRAND = "1F4E79"
_LIGHT = "E8F0FF"


def _is_data_sheet(ws) -> bool:
    if ws.sheet_state != "visible" or ws.title in SUMMARY_SHEETS or ws.max_row < 2:
        return False
    if any(rng.min_row <= 1 <= rng.max_row for rng in ws.merged_cells.ranges):
        return False
    header = [c.value for c in ws[1]]
    filled = [v for v in header if v not in (None, "")]
    return bool(filled) and all(isinstance(v, str) for v in filled)


def _page_setup(ws, *, title_rows: bool) -> None:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.6
    ws.oddFooter.left.text = "&A"
    ws.oddFooter.right.text = "Page &P of &N"
    if title_rows:
        ws.print_title_rows = "1:1"


def _chart_bottom_row(ws) -> int:
    bottom = 0
    for ch in getattr(ws, "_charts", []):
        anchor = ch.anchor
        try:
            row = anchor._from.row + 1 if hasattr(anchor, "_from") else int("".join(c for c in str(anchor) if c.isdigit()))
        except Exception:
            continue
        bottom = max(bottom, row + math.ceil((getattr(ch, "height", 10) or 10) / 0.5) + 1)
    return bottom


def _polish_exec_summary(wb, ws) -> None:
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    # milhares nos números
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                c.number_format = "#,##0"
    # data bars na coluna "Error Count" do Top 10
    for row in ws.iter_rows(min_col=1, max_col=3):
        if [row[0].value, row[1].value, row[2].value] == ["Rank", "Field", "Error Count"]:
            first = row[0].row + 1
            last = first
            while ws.cell(row=last + 1, column=3).value not in (None, ""):
                last += 1
            if ws.cell(row=first, column=3).value not in (None, ""):
                ws.conditional_formatting.add(
                    f"C{first}:C{last}", DataBarRule(start_type="num", start_value=0, end_type="max", color="F4B183", showValue=True)
                )
            break
    # guia de folhas com hiperlinks (no fim, abaixo dos gráficos)
    sheets = [s for s in wb.worksheets if s.sheet_state == "visible" and s.title != ws.title]
    if not sheets:
        return
    r = max(ws.max_row, _chart_bottom_row(ws)) + 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    h = ws.cell(row=r, column=1, value="SHEET GUIDE")
    h.font = Font(bold=True, color="FFFFFF")
    h.fill = PatternFill("solid", start_color="2E75B6", end_color="2E75B6")
    h.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[r].height = 22
    for s in sheets:
        r += 1
        desc = SHEET_GUIDE.get(s.title) or (
            f"Materials missing the {s.title.split('_', 1)[1]} view" if s.title.startswith("Missing_") else ""
        )
        link = ws.cell(row=r, column=1, value=s.title)
        link.hyperlink = f"#'{s.title}'!A1"
        link.font = Font(color="0563C1", underline="single")
        n = max(s.max_row - 1, 0)
        ws.cell(row=r, column=2, value=n).number_format = '#,##0" rows"'
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="right")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        d = ws.cell(row=r, column=3, value=desc)
        d.font = Font(italic=True, color="595959")
    # coluna B (nomes de campos do Top 10) sem truncar
    longest = max((len(str(c.value)) for c in ws["B"] if isinstance(c.value, str)), default=16)
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, min(longest + 4, 46))
    _page_setup(ws, title_rows=False)
    ws.page_setup.fitToHeight = 1       # Executive Summary: uma página (gráficos não se partem)


def polish_workbook(wb) -> None:
    """Aplica o polish a todas as folhas visíveis de um workbook openpyxl (in-place)."""
    for ws in list(wb.worksheets):
        if ws.sheet_state != "visible":
            continue
        try:
            if ws.title == "Executive Summary":
                _polish_exec_summary(wb, ws)
            elif ws.title in SUMMARY_SHEETS:
                ws.sheet_view.showGridLines = False
                _page_setup(ws, title_rows=False)
            elif _is_data_sheet(ws):
                if not ws.auto_filter.ref:
                    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
                _page_setup(ws, title_rows=True)
            else:
                _page_setup(ws, title_rows=False)
        except Exception:
            # o polish nunca pode impedir a geração do relatório
            continue
