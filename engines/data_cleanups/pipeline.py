"""Pipeline: Data Cleanups."""

import io
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from engines.common.fast_loader import load_xlsx_fast
from engines.common.helpers import _normalize_code_as_str_basic
from engines.material_rules.cleanup_rules import (
    validar_ekpo_delivery_completion,
    validar_fp50_eop_expired,
    validar_fp55_eop_vs_service_liability,
    validar_fp55_service_liability_expired,
    validar_fp_existence_zplm_ta_ifs,
    validar_fp_pt01_pt02_status55,
    validar_fppp70_standard_price_mbew,
    validar_ltb_fields,
    validar_mard_total_stock,
)
from engines.material_rules.existence import (
    validar_mara_vs_zplm_mara_existence,
    validar_marc_vs_zplm_marc_existence,
    validar_zplm_mara_vs_mara_existence,
    validar_zplm_marc_vs_marc_existence,
)
from engines.sap_material.fp_status import run_fp_status_validation
from engines.common.excel_polish import polish_workbook

def run_data_cleanups_validation(
    mara_file, marc_file, zplm_marc_file, zplm_mara_file,
    mbew_file, ekpo_file, mard_file, zplm_ta_ifs_file=None, plant_filter=None,
):
    """
    Data Cleanup validation engine:
      - FP Level Status classification
      - FP/PT01+PT02 status 55 check
      - EOP vs Service Liability date check
      - Past Service Liability date check
      - Standard Price (MBEW) check
      - Cross-existence: MARA ↔ ZPLM MARA, MARC ↔ ZPLM MARC
      - LTB field validation
      - EKPO delivery completion
      - MARD total stock check
    Returns: (excel_bytes, results_dict, kpi_fp)
    """
    mara      = load_xlsx_fast(mara_file,      sheet_name=0)
    marc      = load_xlsx_fast(marc_file,      sheet_name=0)
    zplm_marc = load_xlsx_fast(zplm_marc_file, sheet_name=0)
    zplm_mara = load_xlsx_fast(zplm_mara_file, sheet_name=0)
    zplm_ta_ifs = load_xlsx_fast(zplm_ta_ifs_file, sheet_name=0) if zplm_ta_ifs_file else None
    mbew      = load_xlsx_fast(mbew_file,      sheet_name=0)
    ekpo      = load_xlsx_fast(ekpo_file,      sheet_name=0)
    mard      = load_xlsx_fast(mard_file,      sheet_name=0)

    if mara is None or marc is None or zplm_marc is None or zplm_ta_ifs is None:
        raise ValueError("Missing required files for Data Cleanups (MARA, MARC, ZPLM MARC, ZPLM_TA_IFS).")

    # ===============================================================
    # ✅ Fix #1 — When filtering by plant, DO NOT REMOVE materials
    #             without Plant (Plant = "" or NaN)
    # ===============================================================
    def filter_preserving_noPlant(df, pf):
        if df is None or "Plant" not in df.columns:
            return df
        df["_PL"] = df["Plant"].astype(str).str.strip().str.upper()
        withPlant     = df[df["_PL"] == pf].copy()
        withoutPlant  = df[df["_PL"] == ""].copy()   # preserve No Plant
        df = pd.concat([withPlant, withoutPlant], ignore_index=True)
        df.drop(columns=["_PL"], inplace=True)
        return df

    if plant_filter:
        _pf = str(plant_filter).strip().upper()
        marc      = filter_preserving_noPlant(marc,      _pf)
        zplm_marc = filter_preserving_noPlant(zplm_marc, _pf)
        mbew      = filter_preserving_noPlant(mbew,      _pf)
        ekpo      = filter_preserving_noPlant(ekpo,      _pf)
        mard      = filter_preserving_noPlant(mard,      _pf)

    # ===============================================================
    # CLEANUP GLOBAL EXCLUSION: ignore PSMS 75 and 99 in every cleanup rule.
    # The exclusion is applied per Material+Plant, preserving other plants of
    # the same material when their Part Status is different.
    # ===============================================================
    def _cleanup_status_context(marc_df):
        work = marc_df.copy()
        work.columns = [str(c).strip().upper() for c in work.columns]
        pick = lambda candidates: next((c for c in candidates if c in work.columns), None)
        mat = pick(["MATERIAL", "MATNR"])
        plant = pick(["PLANT", "WERKS"])
        status = pick(["PLANT-SPECIFIC MATERIAL STATUS", "PLANT SPECIFIC MATERIAL STATUS", "MTSTB"])
        if not all([mat, plant, status]):
            raise ValueError("MARC must contain Material, Plant and Plant-Specific Material Status for Cleanup validation.")
        work[mat] = work[mat].map(_normalize_code_as_str_basic).str.strip()
        work[plant] = work[plant].astype(str).str.strip().str.upper()
        work[status] = work[status].map(_normalize_code_as_str_basic)
        status_map = work.drop_duplicates([mat, plant]).set_index([mat, plant])[status].to_dict()
        ignored_pairs = {key for key, value in status_map.items() if value in {"75", "99"}}
        active_pairs = set(status_map) - ignored_pairs
        active_materials = {material for material, _ in active_pairs}
        return status_map, ignored_pairs, active_materials

    part_status_map, ignored_pairs, active_materials = _cleanup_status_context(marc)

    def _filter_pairs(df, plant_candidates=("Plant", "WERKS", "Valuation Area", "BWKEY")):
        if df is None or df.empty:
            return df
        out = df.copy()
        out.columns = [str(c).strip() for c in out.columns]
        upper = {str(c).strip().upper(): c for c in out.columns}
        mat_col = next((upper[c] for c in ("MATERIAL", "MATNR") if c in upper), None)
        plant_col = next((upper[c.upper()] for c in plant_candidates if c.upper() in upper), None)
        if mat_col is None:
            return out
        mats = out[mat_col].map(_normalize_code_as_str_basic).str.strip()
        if plant_col is None:
            return out[mats.isin(active_materials)].copy()
        plants = out[plant_col].astype(str).str.strip().str.upper()
        keep = [((m, p) not in ignored_pairs) for m, p in zip(mats, plants)]
        return out.loc[keep].copy()

    marc = _filter_pairs(marc)
    zplm_marc = _filter_pairs(zplm_marc)
    mbew = _filter_pairs(mbew)
    ekpo = _filter_pairs(ekpo)
    mard = _filter_pairs(mard)
    mara = _filter_pairs(mara)
    zplm_mara = _filter_pairs(zplm_mara)
    # ===============================================================
    # VALIDATIONS
    # ===============================================================
    results = {}

    # FP Status
    def _df_to_xlsx_bytes(df):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        buf.seek(0)
        return buf

    df_fp, kpi_fp = run_fp_status_validation(
        mara_file=_df_to_xlsx_bytes(mara),
        marc_file=_df_to_xlsx_bytes(marc),
        zplm_marc_file=_df_to_xlsx_bytes(zplm_marc),
        zplm_mara_file=_df_to_xlsx_bytes(zplm_mara),
        plant_filter=plant_filter
    )
    results["FP Level Status"] = df_fp
    results["FP missing in ZPLM_TA_IFS"] = validar_fp_existence_zplm_ta_ifs(zplm_marc, zplm_ta_ifs, marc)

    results["FP in PT02"]                   = validar_fp_pt01_pt02_status55(zplm_marc, marc)
    results["EOPvsSL (Blank or Swapped)"]   = validar_fp55_eop_vs_service_liability(zplm_marc, marc)
    results["Past Service Liability"]       = validar_fp55_service_liability_expired(zplm_marc, marc)
    results["EOP Expired"]                  = validar_fp50_eop_expired(zplm_marc, marc)
    results["No Standard Cost"]             = validar_fppp70_standard_price_mbew(zplm_marc, marc, mbew)
    results["Missing in ZPLM_MARC"]         = validar_marc_vs_zplm_marc_existence(marc, zplm_marc, mara)
    results["Missing in ZPLM_MARA"]         = validar_mara_vs_zplm_mara_existence(mara, zplm_mara)
    results["Not in MARC"]                  = validar_zplm_marc_vs_marc_existence(zplm_marc, marc, mara)
    results["Not in MARA"]                  = validar_zplm_mara_vs_mara_existence(zplm_mara, mara)
    results["FP LTB Fields"]                = validar_ltb_fields(zplm_marc, marc, mara)
    results["PO's to Close"]                = validar_ekpo_delivery_completion(marc, mara, zplm_marc, ekpo)
    results["OBSOLETE with No Stock"]       = validar_mard_total_stock(mard, marc, mara)

    # Add plant-specific Part Status to every final cleanup dataset.
    def _add_part_status(df):
        if df is None or df.empty or "Material" not in df.columns:
            return df
        out = df.copy()
        if "Part Status" in out.columns:
            return out
        materials = out["Material"].map(_normalize_code_as_str_basic).str.strip()
        if "Plant" in out.columns:
            plants = out["Plant"].fillna("").astype(str).str.strip().str.upper()
            out["Part Status"] = [part_status_map.get((m, p), "") for m, p in zip(materials, plants)]
        else:
            statuses_by_material = {}
            for (material, plant), status in part_status_map.items():
                if (material, plant) in ignored_pairs:
                    continue
                statuses_by_material.setdefault(material, set()).add(status)
            out["Part Status"] = materials.map(
                lambda material: ", ".join(sorted(s for s in statuses_by_material.get(material, set()) if s))
            )
        # Keep Part Status close to Plant for readability.
        column = out.pop("Part Status")
        insert_at = out.columns.get_loc("Plant") + 1 if "Plant" in out.columns else min(1, len(out.columns))
        out.insert(insert_at, "Part Status", column)
        return out

    for result_name in list(results):
        results[result_name] = _add_part_status(results[result_name])
    # ===============================================================
    # Strip datetime (remove time component)
    # ===============================================================
    def strip_datetime(df):
        if df is None or df.empty:
            return df
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.date
        return df

    for k in list(results.keys()):
        results[k] = strip_datetime(results[k])

    # ===============================================================
    # ✅ Fix #2 — DO NOT remove "No Plant" results
    # ===============================================================
    cleaned = {}
    for name, df in results.items():
        if df is None or df.empty:
            continue
        cleaned[name] = df
    results = cleaned

    # ===============================================================
    # ✅ Fix #3 — Write Excel including "No Plant" sheets
    # ===============================================================
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:

        # SUMMARY
        summary_rows = []
        for sheet_name, df in results.items():

            if "Plant" not in df.columns:
                # For FP Level Status, only count non-Correct rows as issues
                if sheet_name == "FP Level Status" and "Status" in df.columns:
                    count = int((df["Status"] != "Correct").sum())
                else:
                    count = len(df)
                summary_rows.append({"Validation": sheet_name, "Plant": "ALL", "Issues": count})
                continue

            unique_plants = sorted(df["Plant"].fillna("").astype(str).unique().tolist())

            for plant in unique_plants:
                label = plant if plant != "" else "No Plant"
                dfp = df[df["Plant"] == plant]
                if not dfp.empty:
                    # For FP Level Status, only count non-Correct rows as issues
                    if sheet_name == "FP Level Status" and "Status" in dfp.columns:
                        count = int((dfp["Status"] != "Correct").sum())
                    else:
                        count = len(dfp)
                    summary_rows.append({"Validation": sheet_name, "Plant": label, "Issues": count})

        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="SUMMARY", index=False)

        # DETAIL SHEETS
        for sheet_name, df in results.items():
            if "Plant" not in df.columns:
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                continue

            unique_plants = sorted(df["Plant"].fillna("").astype(str).unique().tolist())

            for plant in unique_plants:
                dfp = df[df["Plant"] == plant]
                if not dfp.empty:
                    label = plant if plant != "" else "No Plant"
                    dfp.to_excel(writer, sheet_name=f"{sheet_name}_{label}"[:31], index=False)

    # ===============================================================
    # Styling SUMMARY
    # ===============================================================
    out.seek(0)
    wb = load_workbook(out)
    if "SUMMARY" in wb.sheetnames:
        ws = wb["SUMMARY"]

        link_col = ws.max_column + 1
        ws.cell(row=1, column=link_col, value="Link")

        header_fill = PatternFill(start_color="D9D9D9", fill_type="solid")
        zebra       = PatternFill(start_color="F7F7F7", fill_type="solid")
        border      = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"),  bottom=Side(style="thin")
        )

        # Header styling
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border

        # Rows
        for r in range(2, ws.max_row + 1):
            validation = str(ws.cell(row=r, column=1).value)
            plant      = str(ws.cell(row=r, column=2).value)
            sn         = f"{validation}_{plant}" if plant not in ("ALL") else validation
            sn = sn.replace(" ", "_")

            c = ws.cell(row=r, column=link_col)
            if sn in wb.sheetnames:
                c.value = f'=HYPERLINK("#\'{sn}\'!A1","Open")'
                c.font = Font(color="0563C1", underline="single")
            else:
                c.value = "Not found"
            c.alignment = Alignment(horizontal="center")
            c.border = border

            for c2 in ws[r]:
                if r % 2 == 0:
                    c2.fill = zebra
                c2.border = border

        # Auto column width
        from openpyxl.utils import get_column_letter
        for col_idx, col in enumerate(ws.columns, start=1):
            max_len = max((len(str(cell.value or "")) for cell in col), default=8)
            ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2

    polish_workbook(wb)
    out2 = io.BytesIO()
    wb.save(out2)
    out2.seek(0)

    return out2, results, kpi_fp
