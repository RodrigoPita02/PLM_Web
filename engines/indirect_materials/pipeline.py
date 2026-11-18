"""Pipeline: Indirect Materials Report."""

import io
from datetime import datetime
import pandas as pd

from engines.common.excel_report import _style_sap_report
from engines.common.fast_loader import load_xlsx_fast
from engines.indirect_materials.constants import (
    ERROR_COLS,
    IM_MATERIAL_TYPES,
    _IM_FIELD_TO_VIEW,
    _IM_REQUIRED_VIEWS,
    _IM_VIEW_TO_DESC,
    _IM_VIEW_TO_FIELDS,
)
from engines.indirect_materials.helpers import (
    _filter_plant,
    _material_key,
    _material_type_map,
    _norm,
    _pick,
    _prepare_df,
    build_im_sap_like_df,
)
from engines.indirect_materials.rules import (
    im_validar_basic,
    im_validar_general_plant,
    im_validar_mara_wh_packaging,
    im_validar_mbew,
    im_validar_mrp_view1,
    im_validar_mrp_view2,
    im_validar_mrp_view3,
    im_validar_mvke,
    im_validar_plant_storage,
    im_validar_purchasing,
)

def _im_postprocess(
    errors_df: pd.DataFrame,
    sap_im_df: pd.DataFrame,
    mt_map: dict,
    plant_filter: str | None = None,
):
    """
    IM equivalent of postprocess_with_views_roles_and_classification.

    Reads the REAL 'Maintenance status' letters from MARC (via sap_im_df,
    which carries the original MARC columns) to determine which views are
    actually missing per (Material, Plant) — mirroring the SAP engine's
    logic exactly, instead of assuming every required view is missing.

    View 'K' (Basic Data) is always treated as present, even if absent
    from the Maintenance status string, and is therefore never reported
    as a missing view.

    Returns:
        errors_df, views_missing_output_df, contagem_final_df,
        coverage_df, view_tabs_dict
    """
    errors_df = errors_df.copy()

    # 1) Map Field → View
    if "Field" in errors_df.columns:
        errors_df["View"] = errors_df["Field"].map(_IM_FIELD_TO_VIEW).fillna("Not applicable")
    else:
        errors_df["View"] = "Not applicable"

    # Materials flagged by the MARA rule "Maintenance status must contain 'C'"
    # (ERSA + MT0020012). They are a Missing View 'C' (Classification).
    _mara_c_mats: set[str] = set()
    if not errors_df.empty and "Field" in errors_df.columns and "Material" in errors_df.columns:
        _mask_c = errors_df["Field"].astype(str).str.strip() == "Maintenance Status (MARA)"
        _mara_c_mats = {_material_key(m) for m in errors_df.loc[_mask_c, "Material"]}

    mat_col    = _pick(sap_im_df, ["Material", "MATNR"])
    plant_col  = _pick(sap_im_df, ["Plant", "WERKS"])
    mt_col     = _pick(sap_im_df, ["Material Type", "MTART"])
    mg_col     = _pick(sap_im_df, ["Material Group", "MATKL"])
    maint_col  = _pick(sap_im_df, ["Maintenance status", "Maintenance Status", "PSTAT", "Maintained Status"])

    # 2) Build Letters_Present per (Material, Plant), always including 'K'
    #    ('K' — Basic Data — is implicitly present and never reported missing)
    missing_lookup: dict[tuple, set] = {}
    missing_rows = []
    combos: list[tuple[str, str, str]] = []  # (material, plant, mt) for scope/coverage

    for _, row in sap_im_df.iterrows():
        mat   = _norm(row[mat_col])           if mat_col   else ""
        plant = _norm(row[plant_col]).upper() if plant_col else ""
        mt    = _norm(row[mt_col]).upper()    if mt_col    else mt_map.get(_material_key(mat), "")
        if mt not in IM_MATERIAL_TYPES:
            continue
        combos.append((mat, plant, mt))

        raw_status = _norm(row[maint_col]) if maint_col else ""
        present = set(raw_status.upper()) if raw_status else set()
        present.add("K")  # always present, regardless of Maintenance status content

        required = set(_IM_REQUIRED_VIEWS.get(mt, set()))
        # View 'S' (Warehouse Management) is only required for ERSA
        # materials belonging to Material Group MT0020012 — everyone
        # else is out of scope for this view entirely.
        mg = _norm(row[mg_col]).upper() if mg_col else ""
        if mt == "ERSA" and mg == "MT0020012":
            required.add("S")

        missing  = (required - present) - {"K"}  # K is never missing
        # MARA-level rule: ERSA MT0020012 without letter 'C' -> view C missing
        if _material_key(mat) in _mara_c_mats:
            missing = missing | {"C"}

        key = (mat, plant)
        missing_lookup[key] = missing

        if missing:
            missing_rows.append({
                "Material":         mat,
                "Plant":            plant,
                "Material Type":    mt,
                "Letters_Present":  ",".join(sorted(present)),
                "Letters_Missing":  ",".join(sorted(missing)),
            })

    views_missing_df = pd.DataFrame(missing_rows)

    # 3) Classify each error row: Missing View (justified) vs Field Error
    def _classify(row):
        # The MARA validation itself proves view 'C' is not maintained
        if str(row.get("Field", "")).strip() == "Maintenance Status (MARA)":
            return "Missing View"
        view = str(row.get("View", "")).strip()
        if view in ("", "Not applicable", "K"):
            return "Field Error"
        key = (str(row.get("Material", "")).strip(), str(row.get("Plant", "")).strip())
        if view in missing_lookup.get(key, set()):
            return "Missing View"
        return "Field Error"

    if not errors_df.empty:
        errors_df["Error_Type"]                = errors_df.apply(_classify, axis=1)
        errors_df["Justified_By_Missing_View"] = errors_df["Error_Type"] == "Missing View"
    else:
        errors_df["Error_Type"]                = pd.Series(dtype=str)
        errors_df["Justified_By_Missing_View"] = pd.Series(dtype=bool)

    # 4) Split — only actionable Field Errors go into "Field With Error" / Adjusted Summary
    views_missing_output_df = (
        errors_df[~errors_df["Justified_By_Missing_View"]].copy()
        if not errors_df.empty else errors_df.copy()
    )

    # 5) Adjusted Summary (contagem_final_df) — Field Errors only
    if not views_missing_output_df.empty and "Field" in views_missing_output_df.columns:
        contagem_final_df = (
            views_missing_output_df["Field"]
            .value_counts()
            .rename_axis("Field")
            .reset_index(name="Adjusted Nº of Errors")
        )
        contagem_final_df["View"] = contagem_final_df["Field"].map(_IM_FIELD_TO_VIEW).fillna("Not applicable")
        contagem_final_df["View_Description"] = contagem_final_df["View"].map(_IM_VIEW_TO_DESC).fillna("Not applicable")
        contagem_final_df = contagem_final_df.sort_values(
            "Adjusted Nº of Errors", ascending=False
        ).reset_index(drop=True)
    else:
        contagem_final_df = pd.DataFrame(
            columns=["Field", "Adjusted Nº of Errors", "View", "View_Description"]
        )

    # 6) Coverage per View — count of materials ACTUALLY missing each view
    view_counts: dict[str, int] = {}
    for missing in missing_lookup.values():
        for v in missing:
            view_counts[v] = view_counts.get(v, 0) + 1

    coverage_df = pd.DataFrame([
        {
            "View":                          v,
            "View_Description":              _IM_VIEW_TO_DESC.get(v, ""),
            "Materiais without Active View":  cnt,
        }
        for v, cnt in sorted(view_counts.items())
    ]) if view_counts else pd.DataFrame(columns=["View", "View_Description", "Materiais without Active View"])

    # 7) view_tabs_dict — Missing_<View> tabs: only materials ACTUALLY missing that view
    view_tabs_dict: dict[str, pd.DataFrame] = {}
    for mat, plant, mt in combos:
        missing = missing_lookup.get((mat, plant), set())
        for v in missing:
            fields = _IM_VIEW_TO_FIELDS.get(v, [])
            row = {
                "Material":         mat,
                "Plant":            plant,
                "Missing_View":     v,
                "View_Description": _IM_VIEW_TO_DESC.get(v, ""),
                "Fields_Expected":  ", ".join(fields),
            }
            view_tabs_dict.setdefault(v, []).append(row)

    view_tabs_dict = {
        v: pd.DataFrame(rows).drop_duplicates()
        for v, rows in view_tabs_dict.items()
    }

    return (
        errors_df,
        views_missing_output_df,
        contagem_final_df,
        coverage_df,
        view_tabs_dict,
    )


def _build_summary(errors_df: pd.DataFrame) -> pd.DataFrame:
    if errors_df.empty or "Field" not in errors_df.columns:
        return pd.DataFrame(columns=["Field", "Adjusted Nº of Errors"])
    return (
        errors_df["Field"]
        .value_counts()
        .rename_axis("Field")
        .reset_index(name="Adjusted Nº of Errors")
    )


def run_indirect_materials_validation(
    mara_file,
    marc_file,
    mvke_file,
    mbew_file,
    plant_filter: str | None = None,
    progress_callback=None,
):
    """
    Run Indirect Materials validation (ERSA / HIBE / VERP).

    Inputs are raw SAP table exports (no SAPUI5 Export):
      • mara_file — MARA export (.xlsx, sheet 0)
      • marc_file — MARC export (.xlsx, sheet 0)
      • mvke_file — MVKE export (.xlsx, sheet 0)
      • mbew_file — MBEW export (.xlsx, sheet 0)

    Returns: (excel_bytes, errors_df, summary_df)
    """
    def _cb(pct, msg):
        if progress_callback:
            progress_callback(pct, msg)

    _cb(5,  "Loading MARA...")
    mara_df = _prepare_df(load_xlsx_fast(mara_file, sheet_name=0))
    _cb(15, "Loading MARC...")
    marc_df = _prepare_df(load_xlsx_fast(marc_file, sheet_name=0))
    _cb(22, "Loading MVKE...")
    mvke_df = _prepare_df(load_xlsx_fast(mvke_file, sheet_name=0))
    _cb(28, "Loading MBEW...")
    mbew_df = _prepare_df(load_xlsx_fast(mbew_file, sheet_name=0))

    if mara_df is None or mara_df.empty:
        raise ValueError("MARA.xlsx not loaded or empty.")
    if marc_df is None or marc_df.empty:
        raise ValueError("MARC.xlsx not loaded or empty.")

    mt_map = _material_type_map(mara_df)
    if not mt_map:
        raise ValueError("MARA must contain Material and Material Type columns.")

    _cb(35, "Building sap_like_df from MARC + MARA...")
    sap_like_df = build_im_sap_like_df(marc_df, mara_df)

    # Filter to IM material types
    plant_col_sl = _pick(sap_like_df, ["Plant", "WERKS"])
    mat_col_sl   = _pick(sap_like_df, ["Material", "MATNR"])
    mt_col_sl    = _pick(sap_like_df, ["Material Type", "MTART"])

    if not mat_col_sl:
        raise ValueError("MARC/MARA merge is missing a Material column.")

    if mt_col_sl:
        sap_like_df["_IM_MT"] = sap_like_df[mt_col_sl].map(lambda x: _norm(x).upper())
    else:
        sap_like_df["_IM_MT"] = sap_like_df[mat_col_sl].map(
            lambda m: mt_map.get(_norm(m), "")
        )

    sap_im = sap_like_df[sap_like_df["_IM_MT"].isin(IM_MATERIAL_TYPES)].copy()

    if plant_filter and plant_col_sl:
        sap_im = _filter_plant(sap_im, plant_col_sl, plant_filter)

    if sap_im.empty:
        raise ValueError(
            "No Indirect Materials (ERSA/HIBE/VERP) found for the selected scope."
        )

    errors = []
    _cb(40, "Validating Basic / General Data (MARA + MAKT)...")
    im_validar_basic(sap_im, errors)
    _cb(50, "Validating MRP View 1 (MARC)...")
    im_validar_mrp_view1(sap_im, errors)
    _cb(56, "Validating MRP View 2 (MARC)...")
    im_validar_mrp_view2(sap_im, errors)
    _cb(61, "Validating MRP View 3 (MARC)...")
    im_validar_mrp_view3(sap_im, errors)
    _cb(66, "Validating Purchasing View (MARC)...")
    im_validar_purchasing(sap_im, errors)
    _cb(71, "Validating Plant Storage (MARC)...")
    im_validar_plant_storage(sap_im, errors)
    _cb(76, "Validating General Plant Data (MARC)...")
    im_validar_general_plant(sap_im, errors)
    _cb(81, "Validating MBEW...")
    im_validar_mbew(mbew_df, mt_map, errors, plant_filter=plant_filter)
    _cb(86, "Validating MARA packaging / WH fields...")
    im_validar_mara_wh_packaging(mara_df, mt_map, errors)
    _cb(91, "Validating MVKE...")
    im_validar_mvke(mvke_df, mt_map, errors)

    errors_df = (
        pd.DataFrame(errors, columns=ERROR_COLS)
        if errors
        else pd.DataFrame(columns=ERROR_COLS)
    )

    _cb(93, "Post-processing views and summary...")
    (
        errors_df,
        views_missing_output_df,
        contagem_final_df,
        coverage_df,
        view_tabs_dict,
    ) = _im_postprocess(errors_df, sap_im, mt_map, plant_filter=plant_filter)

    summary_df = _build_summary(errors_df)

    _cb(95, "Building Excel report...")
    out = io.BytesIO()

    if not contagem_final_df.empty and "Adjusted Nº of Errors" in contagem_final_df.columns:
        contagem_final_df = contagem_final_df.sort_values(
            "Adjusted Nº of Errors", ascending=False
        ).reset_index(drop=True)

    base_row = 1
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, sheet_name="Executive Summary", index=False)

        if not contagem_final_df.empty:
            contagem_final_df.to_excel(
                writer, sheet_name="Adjusted Summary", index=False, startrow=1
            )
            base_row = len(contagem_final_df) + 5
        if not coverage_df.empty:
            coverage_df.to_excel(
                writer, sheet_name="Adjusted Summary", index=False, startrow=base_row
            )

        views_missing_output_df.to_excel(
            writer, sheet_name="Field With Error", index=False
        )
        errors_df.to_excel(writer, sheet_name="All Errors", index=False)

        # Consolidated "View is Missing" tab (mirrors SAP engine)
        _cols_vm = ["Material", "Plant", "Missing_View", "View_Description", "Fields_Expected"]
        _frames_vm = [
            dfv[_cols_vm] for dfv in view_tabs_dict.values()
            if dfv is not None and not dfv.empty
        ]
        if _frames_vm:
            view_is_missing_df = pd.concat(_frames_vm, ignore_index=True).drop_duplicates()
            view_is_missing_df = view_is_missing_df.sort_values(
                ["Material", "Plant", "Missing_View"]
            )
        else:
            view_is_missing_df = pd.DataFrame(columns=_cols_vm)
        view_is_missing_df.to_excel(writer, sheet_name="View is Missing", index=False)

        for view_code, df_view in view_tabs_dict.items():
            df_view.to_excel(
                writer, sheet_name=f"Missing_{view_code}"[:31], index=False
            )

        sap_im.drop(columns=["_IM_MT"], errors="ignore").to_excel(
            writer, sheet_name="Original Data", index=False
        )

    out.seek(0)
    styled = _style_sap_report(
        out,
        errors_df=errors_df,
        views_missing_output_df=views_missing_output_df,
        contagem_final_df=contagem_final_df,
        coverage_df=coverage_df,
        view_tabs_dict=view_tabs_dict,
        sap_df=sap_im.drop(columns=["_IM_MT"], errors="ignore"),
        plant_filter=plant_filter,
    )

    _cb(100, "Done.")
    return styled, errors_df, summary_df


def build_im_log_payload(
    *,
    user,
    plant_filter,
    mara_file,
    marc_file,
    mvke_file,
    mbew_file,
    errors_df,
    summary_df,
    report_path,
):
    return {
        "run_type":     "INDIRECT_MATERIALS",
        "timestamp":    datetime.now().isoformat(timespec="seconds"),
        "user":         user,
        "plant_filter": plant_filter,
        "files": {
            "mara": getattr(mara_file, "name", "MARA.xlsx"),
            "marc": getattr(marc_file, "name", "MARC.xlsx"),
            "mvke": getattr(mvke_file, "name", "MVKE.xlsx"),
            "mbew": getattr(mbew_file, "name", "MBEW.xlsx"),
        },
        "error_count":    int(len(errors_df)),
        "material_count": int(errors_df["Material"].nunique()) if not errors_df.empty else 0,
        "field_count":    int(errors_df["Field"].nunique())    if not errors_df.empty else 0,
        "summary":        summary_df.head(20).to_dict(orient="records") if summary_df is not None else [],
        "report_path":    report_path,
    }
