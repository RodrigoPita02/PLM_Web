"""Pipeline: SSA Validation."""

import io
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from engines.common.fast_loader import load_xlsx_fast
from engines.ssa.diagnostics import _ssa_print_input_diagnostics, _ssa_run_input_diagnostics
from engines.ssa.helpers import (
    _ssa_add_order_type_column,
    _ssa_build_doc_plant_map,
    _ssa_doc_type_map,
    _ssa_drop_rejected_items,
    _ssa_fill_missing_plant,
    _ssa_filter_docs,
    _ssa_filter_non_s_items,
    _ssa_filter_vbpa_non_s_items,
    _ssa_fully_rejected_docs,
    _ssa_load,
    _ssa_pick,
    _ssa_rejected_item_keys,
)
from engines.ssa.rules import (
    ssa_validar_customer_material,
    ssa_validar_customer_reference,
    ssa_validar_incompletion_log,
    ssa_validar_incoterms_ddp_gb,
    ssa_validar_mandatory_fields,
    ssa_validar_mrp_for_ds,
    ssa_validar_order_without_items,
    ssa_validar_rejected_material_status,
    ssa_validar_shipping_point,
    ssa_validar_terms_of_payment,
    ssa_validar_unloading_point_consistency,
    ssa_validar_weights,
)
from engines.ssa.state import _SSA_ACTIVE_DOC_TYPE_MAP
from engines.common.excel_polish import polish_workbook

def run_ssa_validation(
    vbak_file, vbap_file, vbkd_file, kna1_file,
    marc_file=None, vbpa_file=None, vbfa_file=None, vbep_file=None, mvke_file=None,
    plant_filter=None,
    apply_styles=True,
):
    """
    Run all SSA validations and return (excel_bytes, results_dict, summary_df).

    Business rules implemented here:
      1. Rule applicability is based on Sales Document Type (AUART), not on
         Sales Document number ranges.
      2. A document is closed when every VBAP item is 'done': Overall
         Status = C, or Reason for Rejection (VBAP-ABGRU) filled. Closed
         documents are excluded from all validations except the Material
         Status 70/75/99 rule. Within that rule (and every other
         item-level VBAP check), individual items that are complete
         (Overall Status = C) or rejected (VBAP-ABGRU filled) are always
         skipped, even inside otherwise-active documents.
      3. S-* materials are considered only by Mandatory Fields, Terms of
         Payment, Customer Reference and Incompletion Log. They are excluded
         from all other validations.
      4. Within Mandatory Fields, the S-materials in SSA_S_MATERIAL_LIST have
         their Acct Assmt Grp Mat. (KTGRM) checked against the value carried
         for them in MVKE, when MVKE is supplied.
    """
    from openpyxl.utils import get_column_letter as _gcl

    # Load source tables.
    vbak = load_xlsx_fast(vbak_file, sheet_name=0) if vbak_file else None
    vbap = load_xlsx_fast(vbap_file, sheet_name=0) if vbap_file else None
    vbkd = load_xlsx_fast(vbkd_file, sheet_name=0) if vbkd_file else None
    kna1 = load_xlsx_fast(kna1_file, sheet_name=0) if kna1_file else None
    marc = load_xlsx_fast(marc_file, sheet_name=0) if marc_file else None
    vbpa = load_xlsx_fast(vbpa_file, sheet_name=0) if vbpa_file else None
    vbfa = load_xlsx_fast(vbfa_file, sheet_name=0) if vbfa_file else None
    vbep = load_xlsx_fast(vbep_file, sheet_name=0) if vbep_file else None
    mvke = load_xlsx_fast(mvke_file, sheet_name=0) if mvke_file else None

    if any(df is None or df.empty for df in [vbak, vbap, vbkd, kna1]):
        raise ValueError("Missing required files: VBAK, VBAP, VBKD and KNA1 are mandatory.")

    vbak = _ssa_load(vbak)
    vbap = _ssa_load(vbap)
    vbkd = _ssa_load(vbkd)
    kna1 = _ssa_load(kna1)
    marc = _ssa_load(marc) if marc is not None else None
    vbpa = _ssa_load(vbpa) if vbpa is not None else None
    vbfa = _ssa_load(vbfa) if vbfa is not None else None
    vbep = _ssa_load(vbep) if vbep is not None else None
    mvke = _ssa_load(mvke) if mvke is not None else None

    # Validate the complete structural contract before filtering any dataset.
    # This makes silent rule skipping visible in the console and in the output.
    field_diagnostics_df = _ssa_run_input_diagnostics(
        vbak=vbak,
        vbap=vbap,
        vbkd=vbkd,
        kna1=kna1,
        marc=marc,
        vbpa=vbpa,
        vbfa=vbfa,
        vbep=vbep,
        mvke=mvke,
    )
    _ssa_print_input_diagnostics(field_diagnostics_df)

    # Plant scope must be applied before closed-document evaluation so the
    # report reflects the selected operational scope.
    if plant_filter:
        plt_col = _ssa_pick(vbap, ["Plant", "WERKS"])
        if plt_col:
            vbap = vbap[
                vbap[plt_col].astype(str).str.strip().str.upper()
                == str(plant_filter).strip().upper()
            ].copy()

    # Make AUART available to every rule that checks Sales Document Type
    # applicability via _ssa_doc_type_in. Rule applicability is determined
    # solely by AUART — the Sales Document number is never inspected.
    _SSA_ACTIVE_DOC_TYPE_MAP.clear()
    _SSA_ACTIVE_DOC_TYPE_MAP.update(_ssa_doc_type_map(vbak))

    closed_docs = _ssa_fully_rejected_docs(vbap)

    vbak_doc_col = _ssa_pick(vbak, ["Sales Document", "VBELN"])
    all_vbak_docs = set(vbak[vbak_doc_col].astype(str).str.strip()) if vbak_doc_col else set()
    active_docs = all_vbak_docs - closed_docs

    # Active datasets are used by every normal validation. The status rule is
    # intentionally executed against the unfiltered VBAP dataset.
    active_vbak = _ssa_filter_docs(vbak, active_docs, ["Sales Document", "VBELN"])
    active_vbap = _ssa_filter_docs(vbap, active_docs, ["Sales Document", "VBELN"])
    active_vbkd = _ssa_filter_docs(vbkd, active_docs, ["SD Document", "Sales Document", "VBELN"])
    active_vbpa = _ssa_filter_docs(vbpa, active_docs, ["SD Document", "Sales Document", "VBELN"])
    active_vbfa = _ssa_filter_docs(vbfa, active_docs, ["Preceding Document", "VBELV"])
    active_vbep = _ssa_filter_docs(vbep, active_docs, ["Sales Document", "SD Document", "VBELN"])

    # S-* is included only in the four explicitly approved rules.
    non_s_vbap = _ssa_filter_non_s_items(active_vbap)
    non_s_all_vbap = _ssa_filter_non_s_items(vbap)

    non_s_doc_col = _ssa_pick(non_s_vbap, ["Sales Document", "VBELN"]) if non_s_vbap is not None else None
    non_s_docs = (
        set(non_s_vbap[non_s_doc_col].astype(str).str.strip())
        if non_s_vbap is not None and not non_s_vbap.empty and non_s_doc_col
        else set()
    )
    non_s_vbak = _ssa_filter_docs(active_vbak, non_s_docs, ["Sales Document", "VBELN"])
    non_s_vbkd = _ssa_filter_docs(active_vbkd, non_s_docs, ["SD Document", "Sales Document", "VBELN"])
    non_s_vbpa = _ssa_filter_vbpa_non_s_items(active_vbpa, active_vbap)
    non_s_vbfa = _ssa_filter_docs(active_vbfa, non_s_docs, ["Preceding Document", "VBELV"])
    non_s_vbep = _ssa_filter_docs(active_vbep, non_s_docs, ["Sales Document", "SD Document", "VBELN"])

    all_frames = [
        # S-* included.
        ssa_validar_mandatory_fields(active_vbak, active_vbap, marc, mvke),
        # S-* excluded.
        ssa_validar_incoterms_ddp_gb(non_s_vbak, non_s_vbkd, kna1, non_s_vbpa, non_s_vbap),
        # S-* included.
        ssa_validar_incompletion_log(active_vbak, active_vbap),
        # Closed documents are not globally filtered for this exception rule;
        # S-* items remain excluded.
        ssa_validar_rejected_material_status(None, non_s_all_vbap, marc),
        # S-* included. This is header-level and uses all active orders.
        ssa_validar_terms_of_payment(active_vbak, active_vbkd),
        # S-* excluded.
        ssa_validar_shipping_point(non_s_vbap),
        ssa_validar_unloading_point_consistency(non_s_vbpa, non_s_vbak),
        # S-* included.
        ssa_validar_customer_reference(active_vbap, active_vbkd, active_vbak),
        # S-* excluded.
        ssa_validar_weights(non_s_vbap, non_s_vbfa, non_s_vbep),
        ssa_validar_order_without_items(non_s_vbak, non_s_vbap),
        ssa_validar_customer_material(non_s_vbap),
        ssa_validar_mrp_for_ds(non_s_vbak),
    ]

    all_frames = [_ssa_add_order_type_column(frame) for frame in all_frames]

    # ZLZ items with Reason for Rejection (VBAP-ABGRU) filled are never
    # reported, whatever the rule. Applied once, centrally, so no individual
    # rule can leak them (and future rules are covered automatically).
    _rejected_keys = _ssa_rejected_item_keys(vbap)
    all_frames = [_ssa_drop_rejected_items(frame, _rejected_keys) for frame in all_frames]

    non_empty_frames = [frame for frame in all_frames if frame is not None and not frame.empty]
    all_errors_df = (
        pd.concat(non_empty_frames, ignore_index=True)
        if non_empty_frames
        else pd.DataFrame(columns=[
            "Table", "Sales Document", "Order Type", "Item", "Material",
            "Plant", "Field", "Value", "Error",
        ])
    )

    # Backfill blank Plant values using the Sales Document -> Plant
    # relationship in VBAP, so 'Field With Error' always shows Plant when
    # it can be derived, even for header-level rules that don't report it.
    doc_plant_map = _ssa_build_doc_plant_map(vbap)
    all_errors_df = _ssa_fill_missing_plant(all_errors_df, doc_plant_map)

    rule_names = [
        "Mandatory Fields",
        "Incoterms",
        "Incompletion Log",
        "Rejected + Status 70-75-99",
        "Terms of Payment",
        "Shipping Point",
        "Unloading Point Consistency",
        "Customer Reference",
        "Weights",
        "Order Without Items",
        "Customer Material",
        "MRP for DS Type",
    ]
    results = {
        name: frame
        for name, frame in zip(rule_names, all_frames)
        if frame is not None and not frame.empty
    }

    if not all_errors_df.empty and "Field" in all_errors_df.columns:
        summary_df = (
            all_errors_df["Field"]
            .value_counts()
            .rename_axis("Field")
            .reset_index(name="Nº of Errors")
        )
    else:
        summary_df = pd.DataFrame(columns=["Field", "Nº of Errors"])

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Adjusted Summary", index=False)
        all_errors_df.to_excel(writer, sheet_name="Field With Error", index=False)
        field_diagnostics_df.to_excel(writer, sheet_name="Field Diagnostics", index=False)

    if not apply_styles:
        out.seek(0)
        return out, results, summary_df

    out.seek(0)
    wb = load_workbook(out)

    thin = Side(style="thin", color="B8CCE4")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    zebra = PatternFill(start_color="F5F9FF", fill_type="solid")

    tab_colors = {
        "Adjusted Summary": "2E75B6",
        "Field With Error": "FF7200",
        "Field Diagnostics": "FFD966",
    }
    header_colors = {
        "Adjusted Summary": "BDD7EE",
        "Field With Error": "FFE0CC",
        "Field Diagnostics": "FFF2CC",
    }

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.sheet_properties.tabColor = tab_colors.get(sheet_name, "4F81BD")
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 26

        header_fill = PatternFill(
            start_color=header_colors.get(sheet_name, "BDD7EE"),
            fill_type="solid",
        )
        header_font = Font(bold=True, color="1F4E79")
        header_alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = bdr

        error_col_idx = next(
            (
                index
                for index, cell in enumerate(ws[1], start=1)
                if str(cell.value) == "Error"
            ),
            None,
        )

        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            for cell in row:
                cell.border = bdr
                cell.alignment = Alignment(
                    vertical="center",
                    wrap_text=(cell.column == error_col_idx),
                )
                if row_idx % 2 == 0 and cell.fill.fgColor.rgb in (
                    "00000000", "FFFFFFFF", "00FFFFFF"
                ):
                    cell.fill = zebra

        for col_idx, column in enumerate(ws.columns, start=1):
            max_len = max((len(str(cell.value or "")) for cell in column), default=8)
            ws.column_dimensions[_gcl(col_idx)].width = min(max_len + 3, 55)

        if error_col_idx:
            ws.column_dimensions[_gcl(error_col_idx)].width = 60

    desired_order = ["Adjusted Summary", "Field With Error", "Field Diagnostics"]
    existing = [name for name in desired_order if name in wb.sheetnames]
    wb._sheets.sort(
        key=lambda ws: existing.index(ws.title) if ws.title in existing else 99
    )

    polish_workbook(wb)
    out2 = io.BytesIO()
    wb.save(out2)
    out2.seek(0)
    return out2, results, summary_df
