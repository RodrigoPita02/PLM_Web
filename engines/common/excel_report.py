"""Estilização e formatação condicional dos relatórios Excel gerados."""

import io
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList

from engines.assignments.store import style_assignments_sheet
from engines.common.excel_polish import polish_workbook

def _style_adjusted_summary_sheet(xlsx_bytes: io.BytesIO) -> io.BytesIO:
    """Legacy wrapper — delegates to _style_sap_report."""
    return _style_sap_report(xlsx_bytes)


def _style_sap_report(
    xlsx_bytes: io.BytesIO,
    *,
    errors_df: pd.DataFrame | None = None,
    views_missing_output_df: pd.DataFrame | None = None,
    contagem_final_df: pd.DataFrame | None = None,
    coverage_df: pd.DataFrame | None = None,
    view_tabs_dict: dict | None = None,
    sap_df: pd.DataFrame | None = None,
    plant_filter: str | None = None,
) -> io.BytesIO:
    """
    Full professional styling for the SAP Material Validation report.
    Improvements:
      1. Global styling: freeze, bold headers, auto-width, borders on every sheet
      2. Error_Type colouring in All Errors (Missing View=yellow, Field Error by severity)
      3. Error keyword colouring in Field With Error
      4. Adjusted Summary: % of Total, sorted descending, freeze, auto-width
      5. Executive Summary sheet: KPI overview, top fields, plant breakdown
      6. Tab colours by sheet type
      7. Sheet order: Exec Summary first, Original Data last
      8. Missing_* sheets: freeze + auto-width
    """
    from openpyxl.utils import get_column_letter as _gcl
    xlsx_bytes.seek(0)
    wb = load_workbook(xlsx_bytes)

    CLR = {
        "header_blue":   "1F4E79",
        "header_fill":   "BDD7EE",
        "tab_exec":      "1F4E79",
        "tab_summary":   "AEC6FF",
        "tab_field_err": "FF7200",
        "tab_all_err":   "C00000",
        "tab_view_miss": "FFD700",
        "tab_missing":   "A6A6A6",
        "tab_original":  "70AD47",
        "row_red":       "FFD7D7",
        "row_orange":    "FFE5B4",
        "row_yellow":    "FFFACD",
        "zebra":         "F5F9FF",
        "border_col":    "B8CCE4",
    }

    def _fill(h):
        return PatternFill(start_color=h, end_color=h, fill_type="solid")

    def _bdr():
        t = Side(style="thin", color=CLR["border_col"])
        return Border(left=t, right=t, top=t, bottom=t)

    def _style_sheet(ws, *, tab_color=None, freeze="B2", zebra=True):
        """Apply global styling: header, freeze, auto-width, borders."""
        if tab_color:
            ws.sheet_properties.tabColor = tab_color
        if ws.max_row < 1:
            return
        hf = _fill(CLR["header_fill"])
        hfont = Font(bold=True, color=CLR["header_blue"])
        halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 28
        for cell in ws[1]:
            cell.fill = hf; cell.font = hfont
            cell.alignment = halign; cell.border = _bdr()
        if freeze:
            ws.freeze_panes = freeze
        zf = _fill(CLR["zebra"])
        for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            for cell in row:
                cell.border = _bdr()
                if zebra and r_idx % 2 == 0:
                    if cell.fill.fgColor.rgb in ("00000000", "FFFFFFFF", "00FFFFFF"):
                        cell.fill = zf
        for col_idx, col in enumerate(ws.columns, start=1):
            max_len = 0
            for cell in col:
                try:
                    v = len(str(cell.value)) if cell.value is not None else 0
                    if v > max_len: max_len = v
                except Exception:
                    pass
            ws.column_dimensions[_gcl(col_idx)].width = min(max_len + 3, 55)

    # ── Executive Summary ────────────────────────────────────────────────────
    if "Executive Summary" in wb.sheetnames:
        ws_exec = wb["Executive Summary"]
        ws_exec.sheet_properties.tabColor = CLR["tab_exec"]

        total_errors    = len(errors_df) if errors_df is not None and not errors_df.empty else 0
        total_materials = int(errors_df['Material'].nunique()) if (errors_df is not None and not errors_df.empty and 'Material' in errors_df.columns) else 0
        mv_cnt = int((errors_df['Error_Type'] == 'Missing View').sum()) if (errors_df is not None and not errors_df.empty and 'Error_Type' in errors_df.columns) else 0
        fe_cnt = int((errors_df['Error_Type'] == 'Field Error').sum())  if (errors_df is not None and not errors_df.empty and 'Error_Type' in errors_df.columns) else 0
        # FIX: count distinct fields only from Field Error rows (views_missing_output_df)
        # errors_df includes Missing View errors which inflate the count vs Adjusted Summary
        _fe_df = errors_df[errors_df['Error_Type'] == 'Field Error'] if (errors_df is not None and not errors_df.empty and 'Error_Type' in errors_df.columns) else errors_df
        total_fields = int(_fe_df['Field'].nunique()) if (_fe_df is not None and not _fe_df.empty and 'Field' in _fe_df.columns) else 0

        scope_label = f"Plant: {plant_filter}" if plant_filter else "Scope: All Plants"
        run_date    = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %Z")

        title_font  = Font(bold=True, size=16, color="FFFFFF")
        title_fill  = _fill(CLR["header_blue"])
        title_align = Alignment(horizontal="center", vertical="center")
        kpi_lf  = Font(bold=True, color=CLR["header_blue"])
        kpi_vf  = Font(bold=True, size=14, color="1F4E79")
        sec_font = Font(bold=True, color="FFFFFF")
        sec_fill = _fill("2E75B6")
        nb = _bdr()

        ws_exec.merge_cells("A1:F1")
        c1 = ws_exec["A1"]
        c1.value = "SAP Material Validation — Executive Summary"
        c1.font = title_font; c1.fill = title_fill; c1.alignment = title_align
        ws_exec.row_dimensions[1].height = 36

        ws_exec.merge_cells("A2:F2")
        c2 = ws_exec["A2"]
        c2.value = f"{scope_label}   |   Generated: {run_date}"
        c2.font = Font(italic=True, color="595959")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        ws_exec.row_dimensions[2].height = 20

        ws_exec.merge_cells("A4:F4")
        s4 = ws_exec["A4"]
        s4.value = "KEY METRICS"; s4.font = sec_font; s4.fill = sec_fill
        s4.alignment = Alignment(horizontal="left", vertical="center")
        ws_exec.row_dimensions[4].height = 22

        # KPIs: removed Total Errors (redundant with Field Errors) and
        # Justified by Missing View. Field Errors is the actionable number.
        kpi_rows = [
            ("Field Errors (Actionable)", fe_cnt,          "FCE4D6"),
            ("Materials with Errors",     total_materials,  "FCE4D6"),
            ("Distinct Fields with Error",total_fields,     "DEEAF1"),
        ]
        for i, (label, value, hex_fill) in enumerate(kpi_rows, start=5):
            ws_exec.row_dimensions[i].height = 22
            lc = ws_exec.cell(row=i, column=1, value=label)
            vc = ws_exec.cell(row=i, column=2, value=value)
            lc.font = kpi_lf; lc.fill = _fill(hex_fill); lc.border = nb
            lc.alignment = Alignment(vertical="center")
            vc.font = kpi_vf; vc.fill = _fill(hex_fill); vc.border = nb
            vc.alignment = Alignment(horizontal="center", vertical="center")

        row_cursor = len(kpi_rows) + 6

        # ── TOP 10 FIELDS table (no % column) ────────────────────────────────
        ws_exec.merge_cells(f"A{row_cursor}:C{row_cursor}")
        sh = ws_exec.cell(row=row_cursor, column=1, value="TOP 10 FIELDS WITH MOST ERRORS")
        sh.font = sec_font; sh.fill = sec_fill
        sh.alignment = Alignment(horizontal="left", vertical="center")
        ws_exec.row_dimensions[row_cursor].height = 22
        row_cursor += 1

        top_fields_data = None
        fields_table_header_row = row_cursor
        if errors_df is not None and not errors_df.empty and 'Field' in errors_df.columns:
            top_fields = errors_df['Field'].value_counts().head(10).reset_index()
            top_fields.columns = ['Field', 'Count']
            top_fields_data = top_fields
            for ci, hdr in enumerate(['Rank', 'Field', 'Error Count'], start=1):
                hc = ws_exec.cell(row=row_cursor, column=ci, value=hdr)
                hc.font = Font(bold=True, color=CLR["header_blue"])
                hc.fill = _fill(CLR["header_fill"]); hc.border = nb
                hc.alignment = Alignment(horizontal="center", vertical="center")
            ws_exec.row_dimensions[row_cursor].height = 20
            row_cursor += 1
            for rank, (_, tf) in enumerate(top_fields.iterrows(), start=1):
                rf = _fill("F5F9FF") if rank % 2 == 0 else _fill("FFFFFF")
                for ci, val in enumerate([rank, tf['Field'], tf['Count']], start=1):
                    rc = ws_exec.cell(row=row_cursor, column=ci, value=val)
                    rc.fill = rf; rc.border = nb
                    rc.alignment = Alignment(
                        horizontal="center" if ci != 2 else "left",
                        vertical="center"
                    )
                row_cursor += 1

        # ── LINE CHART: Weekly Error Evolution (from 9 Feb) ──────────────────
        # Build weekly data from the SAP log history and write it to a hidden
        # area of the sheet (column H onwards), then chart it.
        from engines.common.run_logs import list_run_logs as _lrl
        _sap_logs = _lrl("sap", limit=500) or []

        # Filter logs from 9 Feb onwards and group by ISO week
        _weekly: dict[str, int] = {}
        for _log in _sap_logs:
            try:
                _ts = pd.to_datetime(_log.get("timestamp", ""))
                if _ts < pd.Timestamp("2025-02-09"):
                    continue
                _week_key = _ts.strftime("%Y-W%V")  # e.g. "2025-W07"
                _val = _log.get("metrics", {}).get("total_errors", 0)
                # Keep the latest run per week
                if _week_key not in _weekly or _val > _weekly[_week_key]:
                    _weekly[_week_key] = _val
            except Exception:
                continue

        # Sort weeks chronologically
        _weeks_sorted = sorted(_weekly.keys())

        if _weeks_sorted:
            # Write data to a hidden sheet — Excel renders charts from hidden
            # sheets without issues, unlike hidden columns which block rendering
            _data_sheet_name = "_ChartData"
            if _data_sheet_name in wb.sheetnames:
                del wb[_data_sheet_name]
            _ws_data = wb.create_sheet(_data_sheet_name)
            _ws_data.sheet_state = "hidden"

            _data_start_row = 1
            _ws_data.cell(row=_data_start_row, column=1, value="Week")
            _ws_data.cell(row=_data_start_row, column=2, value="Total Errors")
            for _ri, _wk in enumerate(_weeks_sorted, start=_data_start_row + 1):
                _ws_data.cell(row=_ri, column=1, value=_wk)
                _ws_data.cell(row=_ri, column=2, value=_weekly[_wk])
            _data_end_row = _data_start_row + len(_weeks_sorted)

            from openpyxl.chart import LineChart
            _lc = LineChart()
            _lc.title  = "Weekly Error Evolution (from 9 Feb)"
            _lc.style  = 10
            _lc.y_axis.title = "Total Errors"
            _lc.x_axis.title = "Week"
            _lc.width  = 18
            _lc.height = 12
            _lc.y_axis.numFmt = '0'
            _lc.x_axis.tickLblPos = "low"

            # Reference the hidden sheet for data
            _vals_ref = Reference(_ws_data,
                                   min_col=2, min_row=_data_start_row,
                                   max_row=_data_end_row)
            _cats_ref = Reference(_ws_data,
                                   min_col=1, min_row=_data_start_row + 1,
                                   max_row=_data_end_row)
            _lc.add_data(_vals_ref, titles_from_data=True)
            _lc.set_categories(_cats_ref)

            if _lc.series:
                _s = _lc.series[0]
                _s.graphicalProperties.line.solidFill = "2E75B6"
                _s.graphicalProperties.line.width     = 20000  # 2pt in EMU
                _s.marker.symbol   = "circle"
                _s.marker.size     = 6
                _s.marker.graphicalProperties.solidFill = "2E75B6"
                _s.marker.graphicalProperties.line.solidFill = "1F4E79"

            # Place chart on Executive Summary sheet
            _line_anchor = f"E{fields_table_header_row - 1}"
            ws_exec.add_chart(_lc, _line_anchor)

        row_cursor += 2  # spacing after table + chart area

        # ── ERRORS BY PLANT table (no % column) ──────────────────────────────
        ws_exec.merge_cells(f"A{row_cursor}:B{row_cursor}")
        bp_h = ws_exec.cell(row=row_cursor, column=1, value="ERRORS BY PLANT")
        bp_h.font = sec_font; bp_h.fill = sec_fill
        bp_h.alignment = Alignment(horizontal="left", vertical="center")
        ws_exec.row_dimensions[row_cursor].height = 22
        row_cursor += 1

        by_plant_data = None
        plant_table_header_row = row_cursor
        if errors_df is not None and not errors_df.empty and 'Plant' in errors_df.columns:
            by_plant = (
                errors_df['Plant'].astype(str).str.strip()
                .replace({'': 'General', 'nan': 'General'})
                .value_counts().reset_index()
            )
            by_plant.columns = ['Plant', 'Errors']
            by_plant['Plant'] = by_plant['Plant'].replace('No Plant', 'General')
            by_plant_data = by_plant

            for ci, hdr in enumerate(['Plant', 'Error Count'], start=1):
                hc = ws_exec.cell(row=row_cursor, column=ci, value=hdr)
                hc.font = Font(bold=True, color=CLR["header_blue"])
                hc.fill = _fill(CLR["header_fill"]); hc.border = nb
                hc.alignment = Alignment(horizontal="center", vertical="center")
            ws_exec.row_dimensions[row_cursor].height = 20
            row_cursor += 1
            for rank, (_, bp) in enumerate(by_plant.iterrows(), start=1):
                rf = _fill("F5F9FF") if rank % 2 == 0 else _fill("FFFFFF")
                for ci, val in enumerate([bp['Plant'], bp['Errors']], start=1):
                    rc = ws_exec.cell(row=row_cursor, column=ci, value=val)
                    rc.fill = rf; rc.border = nb
                    rc.alignment = Alignment(horizontal="center", vertical="center")
                row_cursor += 1

        # ── PIE CHART: Errors by Plant — placed to the right of plant table ──
        if by_plant_data is not None and len(by_plant_data) > 0:
            plant_data_min_row = plant_table_header_row + 1
            plant_data_max_row = plant_data_min_row + len(by_plant_data) - 1

            pie_chart = PieChart()
            pie_chart.title = "Errors by Plant"
            pie_chart.style = 10
            pie_chart.width  = 12
            pie_chart.height = 10

            pie_data = Reference(ws_exec,
                                  min_col=2, min_row=plant_table_header_row,
                                  max_row=plant_data_max_row)
            pie_cats = Reference(ws_exec,
                                  min_col=1, min_row=plant_data_min_row,
                                  max_row=plant_data_max_row)
            pie_chart.add_data(pie_data, titles_from_data=True)
            pie_chart.set_categories(pie_cats)
            pie_chart.dataLabels             = DataLabelList()
            pie_chart.dataLabels.showPercent  = True
            pie_chart.dataLabels.showVal      = False
            pie_chart.dataLabels.showCatName  = False  # category shown in legend
            pie_chart.dataLabels.showSerName  = False  # removes "Error Count," prefix

            pie_anchor = f"D{plant_table_header_row - 1}"
            ws_exec.add_chart(pie_chart, pie_anchor)

        for col_letter, width in [("A",38),("B",16),("C",14),("D",12),("E",14),("F",14)]:
            ws_exec.column_dimensions[col_letter].width = width

    # ── Adjusted Summary ─────────────────────────────────────────────────────
    if "Adjusted Summary" in wb.sheetnames:
        ws_sum = wb["Adjusted Summary"]
        ws_sum.sheet_properties.tabColor = CLR["tab_summary"]
        ws_sum.freeze_panes = "A3"

        tf_s = _fill("E8F0FF")
        tf_font = Font(bold=True, size=12, color=CLR["header_blue"])
        tf_align = Alignment(horizontal="center", vertical="center")
        tf_side = Side(style="thin", color="BBD0FF")
        tf_bdr = Border(left=tf_side, right=tf_side, top=tf_side, bottom=tf_side)

        first_header_row = next(
            (r for r in range(1, ws_sum.max_row + 1)
             if any(ws_sum.cell(row=r, column=cc).value not in (None, "")
                    for cc in range(1, ws_sum.max_column + 1))), None)
        if first_header_row:
            first_empty = next(
                (r for r in range(first_header_row + 1, ws_sum.max_row + 1)
                 if all(ws_sum.cell(row=r, column=cc).value in (None, "")
                        for cc in range(1, ws_sum.max_column + 1))), None)
            second_header_row = None
            if first_empty:
                second_header_row = next(
                    (r for r in range(first_empty + 1, ws_sum.max_row + 1)
                     if any(ws_sum.cell(row=r, column=cc).value not in (None, "")
                            for cc in range(1, ws_sum.max_column + 1))), None)

            last_c1 = max((cc for cc in range(1, ws_sum.max_column + 1)
                           if ws_sum.cell(row=first_header_row, column=cc).value not in (None, "")), default=1)
            t1 = first_header_row - 1
            if t1 >= 1:
                ws_sum.merge_cells(start_row=t1, start_column=1, end_row=t1, end_column=last_c1)
                ct1 = ws_sum.cell(row=t1, column=1)
                ct1.value = "Field Errors"; ct1.fill = tf_s; ct1.font = tf_font
                ct1.alignment = tf_align; ct1.border = tf_bdr
                ws_sum.row_dimensions[t1].height = 24

            t2 = None
            if second_header_row:
                last_c2 = max((cc for cc in range(1, ws_sum.max_column + 1)
                               if ws_sum.cell(row=second_header_row, column=cc).value not in (None, "")), default=1)
                t2 = second_header_row - 1
                ws_sum.merge_cells(start_row=t2, start_column=1, end_row=t2, end_column=last_c2)
                ct2 = ws_sum.cell(row=t2, column=1)
                ct2.value = "Materials Without Active View"; ct2.fill = tf_s; ct2.font = tf_font
                ct2.alignment = tf_align; ct2.border = tf_bdr
                ws_sum.row_dimensions[t2].height = 24

            for r in range(1, ws_sum.max_row + 1):
                row_vals = [ws_sum.cell(row=r, column=cc).value for cc in range(1, ws_sum.max_column + 1)]
                is_hdr = (r == first_header_row) or (second_header_row and r == second_header_row)
                is_ttl = (t1 and r == t1) or (t2 and r == t2)
                if is_ttl or all(v in (None, "") for v in row_vals):
                    continue
                for ci in range(1, ws_sum.max_column + 1):
                    cell = ws_sum.cell(row=r, column=ci)
                    if is_hdr:
                        cell.font = Font(bold=True, color=CLR["header_blue"])
                        cell.fill = _fill(CLR["header_fill"])
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    else:
                        cell.alignment = Alignment(vertical="center")
                    cell.border = _bdr()

        for col_idx, col in enumerate(ws_sum.columns, start=1):
            max_len = max((len(str(cell.value)) if cell.value else 0 for cell in col), default=8)
            ws_sum.column_dimensions[_gcl(col_idx)].width = min(max_len + 3, 50)

    # ── Field With Error — no row colour coding, clean zebra only ───────────
    if "Field With Error" in wb.sheetnames:
        _style_sheet(wb["Field With Error"], tab_color=None, freeze="A2", zebra=True)

    # ── All Errors — no row colour coding, clean zebra only ──────────────────
    if "All Errors" in wb.sheetnames:
        _style_sheet(wb["All Errors"], tab_color=None, freeze="A2", zebra=True)

    # ── View is Missing ──────────────────────────────────────────────────────
    if "View is Missing" in wb.sheetnames:
        _style_sheet(wb["View is Missing"], tab_color=None, freeze="A2")

    # ── Missing_* sheets ─────────────────────────────────────────────────────
    for sn in wb.sheetnames:
        if sn.startswith("Missing_"):
            _style_sheet(wb[sn], tab_color=None, freeze="A2")

    # ── Assignments ──────────────────────────────────────────────────────────
    if "Assignments" in wb.sheetnames:
        ws_asgn = wb["Assignments"]
        ws_asgn.sheet_properties.tabColor = "FF7200"
        # Build a minimal df from the sheet to pass to style_assignments_sheet
        try:
            style_assignments_sheet(ws_asgn, pd.DataFrame())
        except Exception:
            _style_sheet(ws_asgn, tab_color="FF7200", freeze="A2")

    # ── Original Data ────────────────────────────────────────────────────────
    if "Original Data" in wb.sheetnames:
        _style_sheet(wb["Original Data"], tab_color=None, freeze="B2")

    polish_workbook(wb)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def _apply_conditional_formatting_in_memory(xlsx_bytes: io.BytesIO,
                                              sheet_names=("PT01", "PT02")) -> io.BytesIO:
    """
    Applies styling to PLM Workflow sheets:
      - Header row: bold, light-blue fill, freeze pane
      - Auto-width columns (capped at 60)
      - Aging column: green<=30, yellow<=90, orange<=180, red>180
    """
    xlsx_bytes.seek(0)
    wb = load_workbook(xlsx_bytes)

    age_fills = {
        "green":  PatternFill(start_color="92D050", fill_type="solid"),
        "yellow": PatternFill(start_color="FFFF00", fill_type="solid"),
        "orange": PatternFill(start_color="FFA500", fill_type="solid"),
        "red":    PatternFill(start_color="FF0000", fill_type="solid"),
    }
    header_fill   = PatternFill(start_color="BDD7EE", fill_type="solid")
    header_font   = Font(bold=True)
    header_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin          = Side(style="thin", color="9DC3E6")
    header_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    cell_border   = Border(left=thin, right=thin, top=thin, bottom=thin)

    for sheet in sheet_names:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]

        # 1) Style header row + freeze
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 30
        for cell in ws[1]:
            cell.fill      = header_fill
            cell.font      = header_font
            cell.alignment = header_align
            cell.border    = header_border

        # 2) Auto-width for all columns
        from openpyxl.utils import get_column_letter as _gcl
        for col_idx, col in enumerate(ws.columns, start=1):
            max_len = 0
            for cell in col:
                try:
                    cell_len = len(str(cell.value)) if cell.value is not None else 0
                    if cell_len > max_len:
                        max_len = cell_len
                except Exception:
                    pass
            ws.column_dimensions[_gcl(col_idx)].width = min(max_len + 4, 60)

        # 3) Light border on data rows
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = cell_border

        # 4) Aging colour coding
        aging_col_idx = next(
            (idx for idx, cell in enumerate(ws[1], start=1) if cell.value == "Aging"),
            None
        )
        if aging_col_idx is not None:
            for row in ws.iter_rows(min_row=2, min_col=aging_col_idx, max_col=aging_col_idx):
                for cell in row:
                    try:
                        val = int(cell.value)
                    except (TypeError, ValueError):
                        continue
                    if val <= 30:
                        cell.fill = age_fills["green"]
                    elif val <= 90:
                        cell.fill = age_fills["yellow"]
                    elif val <= 180:
                        cell.fill = age_fills["orange"]
                    else:
                        cell.fill = age_fills["red"]

    polish_workbook(wb)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
