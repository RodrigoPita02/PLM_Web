"""Código mantido só por compatibilidade (deprecated)."""

import io
import pandas as pd

from engines.common.fast_loader import load_xlsx_fast
from engines.common.helpers import _coerce_df_for_display
from engines.common.rolesviews import get_rolesviews_bytesio
from engines.material_rules.advanced_planning import validar_advanced_planning_fields
from engines.sap_material.postprocess import postprocess_with_views_roles_and_classification

def run_mrp_validation(mara_file, marc_file, zplm_marc_file=None,
                        *, plant_filter: str | None = None):
    """
    Deprecated standalone runner — Advanced Planning field validations are now
    integrated into run_full_validation() via validar_advanced_planning_fields().
    This function is kept only for backward compatibility.
    Returns: (excel_bytes, errors_df, summary_df)
    """
    import warnings
    warnings.warn(
        "run_mrp_validation() is deprecated. "
        "Advanced Planning validations now run automatically inside run_full_validation().",
        DeprecationWarning,
        stacklevel=2,
    )

    mara_df      = load_xlsx_fast(mara_file, sheet_name=0)
    marc_df      = load_xlsx_fast(marc_file, sheet_name=0)
    zplm_marc_df = load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None

    if mara_df is None or marc_df is None:
        raise ValueError("MARA.xlsx / MARC.xlsx not loaded.")

    for df in (mara_df, marc_df):
        df.columns = [str(c).strip() for c in df.columns]

    if plant_filter and 'Plant' in marc_df.columns:
        pf = str(plant_filter).strip().upper()
        marc_df = marc_df[marc_df['Plant'].astype(str).str.strip().str.upper() == pf].copy()

    resultados = []
    validar_advanced_planning_fields(marc_df, mara_df, resultados)
    errors_df = _coerce_df_for_display(pd.DataFrame(resultados))

    rolesviews_dict = load_xlsx_fast(get_rolesviews_bytesio(), sheet_name=None)
    roles_df = rolesviews_dict.get("Roles")
    views_df = rolesviews_dict.get("Views")

    (errors_df, _, _, contagem_final_df, coverage_df, view_tabs_dict) = \
        postprocess_with_views_roles_and_classification(
            errors_df=errors_df, mara_df=mara_df, marc_df=marc_df,
            roles_df=roles_df, views_df=views_df, plant_filter=plant_filter,
        )

    file_bytes = io.BytesIO()
    with pd.ExcelWriter(file_bytes, engine="openpyxl") as writer:
        errors_df.to_excel(writer, sheet_name="All Errors", index=False)
        if not contagem_final_df.empty:
            contagem_final_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=1)
        for view_code, df_view in view_tabs_dict.items():
            df_view.to_excel(writer, sheet_name=f"Missing_{view_code}"[:31], index=False)

    file_bytes.seek(0)
    summary_df = (
        errors_df['Field'].value_counts().rename_axis('Field').reset_index(name='Adjusted Nº of Errors')
        if 'Field' in errors_df.columns and not errors_df.empty
        else pd.DataFrame({"Field": [], "Adjusted Nº of Errors": []})
    )
    return file_bytes, errors_df, summary_df
