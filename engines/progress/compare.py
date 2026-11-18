"""Comparação de relatórios ajustados (SAP Validation Progress)."""

import io
import pandas as pd

from engines.common.fast_loader import load_xlsx_fast
from engines.common.helpers import _norm_str
from engines.common.rolesviews import get_rolesviews_bytesio

def _load_rolesviews_maps() -> tuple[dict, dict]:
    """Loads embedded RolesViews.xlsx and returns (field_to_view, view_to_desc)."""
    try:
        rv = load_xlsx_fast(get_rolesviews_bytesio(), sheet_name=None)
        roles_df = rv.get("Roles", pd.DataFrame())
        views_df = rv.get("Views", pd.DataFrame())

        field_to_view = {}
        if not roles_df.empty:
            roles_df['Field'] = roles_df['Field'].astype(str).str.strip()
            roles_df['View']  = roles_df['View'].astype(str).str.strip()
            field_to_view = roles_df.set_index('Field')['View'].to_dict()

        view_to_desc = {}
        if not views_df.empty:
            views_df['Codigo']    = views_df['Codigo'].astype(str).str.strip()
            views_df['Descricao'] = views_df['Descricao'].astype(str).str.strip()
            view_to_desc = views_df.set_index('Codigo')['Descricao'].to_dict()

        return field_to_view, view_to_desc
    except Exception:
        return {}, {}


def _extract_missing_views_from_workbook(book_dict: dict) -> dict:
    """
    Reads Missing_<View> sheets (or falls back to 'All Errors' column)
    and returns a map (Material, Plant) -> {view_codes}.
    """
    missing_map = {}
    missing_sheets = [name for name in book_dict if name.startswith("Missing_")]

    if missing_sheets:
        for sheet in missing_sheets:
            df = book_dict.get(sheet, pd.DataFrame())
            if df.empty:
                continue
            view_code = sheet.replace("Missing_", "").strip()
            for _, r in df.iterrows():
                mat = _norm_str(r.get('Material'))
                if not mat:
                    continue
                missing_map.setdefault((_norm_str(r.get('Material')), _norm_str(r.get('Plant'))), set()).add(view_code)
        return missing_map

    df_err = book_dict.get("All Errors", pd.DataFrame())
    if not df_err.empty and {'Material', 'Plant', 'View', 'Justified_By_Missing_View'}.issubset(df_err.columns):
        for _, r in df_err.iterrows():
            if bool(r.get('Justified_By_Missing_View', False)):
                mat, plant, view = _norm_str(r.get('Material')), _norm_str(r.get('Plant')), _norm_str(r.get('View'))
                if mat and view and view not in ('', 'Not applicable'):
                    missing_map.setdefault((mat, plant), set()).add(view)
    return missing_map


def compare_adjusted_reports(current_file, previous_file):
    """
    Compares two validation reports and returns:
      - 'New Errors':       errors in current but not in previous
      - 'Resolved Errors':  errors in previous but not in current
      - 'Changed Errors':   same (Material, Plant, Field) but different Error text
      - 'New Materials':    materials in current Original Data not in previous
      - 'Views Lost':       views newly missing in current vs previous
    Returns: (excel_bytes, diff_dfs_dict)
    """
    current_dict  = load_xlsx_fast(current_file,  sheet_name=None)
    previous_dict = load_xlsx_fast(previous_file, sheet_name=None)

    key_cols = ['Material', 'Plant', 'Field', 'Error']

    def load_errors(d):
        df = d.get("All Errors", d.get("Field With Error", pd.DataFrame()))
        if df.empty or not set(key_cols).issubset(df.columns):
            return pd.DataFrame(columns=key_cols)
        return df[key_cols].dropna(subset=['Material', 'Field']).copy()

    curr_err = load_errors(current_dict)
    prev_err = load_errors(previous_dict)

    for df in [curr_err, prev_err]:
        for c in key_cols:
            df[c] = df[c].astype(str).str.strip()

    curr_set  = set(curr_err.itertuples(index=False, name=None))
    prev_set  = set(prev_err.itertuples(index=False, name=None))

    new_errors      = pd.DataFrame(curr_set - prev_set, columns=key_cols)
    resolved_errors = pd.DataFrame(prev_set - curr_set, columns=key_cols)

    # Build (Material, Plant, Field) -> set[Error] maps via groupby — O(n) instead of O(n²)
    _grp_cols = ['Material', 'Plant', 'Field']
    curr_map = {
        key: set(grp['Error'])
        for key, grp in curr_err.groupby(_grp_cols)
    }
    prev_map = {
        key: set(grp['Error'])
        for key, grp in prev_err.groupby(_grp_cols)
    }
    changed_keys = set(curr_map) & set(prev_map)
    changed_rows = []
    for key in changed_keys:
        c_errs = curr_map[key]
        p_errs = prev_map[key]
        if c_errs != p_errs:
            mat, plant, field = key
            changed_rows.append({'Material': mat, 'Plant': plant, 'Field': field,
                                  'Previous Error': ' | '.join(sorted(p_errs)),
                                  'Current Error':  ' | '.join(sorted(c_errs))})
    changed_errors = pd.DataFrame(changed_rows)

    # New Materials
    def load_materials(d):
        df = d.get("Original Data", pd.DataFrame())
        return set(df['Material'].astype(str).str.strip()) if 'Material' in df.columns else set()

    new_mats = load_materials(current_dict) - load_materials(previous_dict)
    new_materials = pd.DataFrame([{"Material": m} for m in sorted(new_mats)])

    # Views Lost
    curr_views = _extract_missing_views_from_workbook(current_dict)
    prev_views = _extract_missing_views_from_workbook(previous_dict)

    views_lost_rows = []
    for key, curr_set_v in curr_views.items():
        prev_set_v = prev_views.get(key, set())
        newly_lost = curr_set_v - prev_set_v
        for v in sorted(newly_lost):
            views_lost_rows.append({'Material': key[0], 'Plant': key[1], 'View_Lost': v})
    views_lost = pd.DataFrame(views_lost_rows)

    diff_dfs = {
        "New Errors":      new_errors,
        "Resolved Errors": resolved_errors,
        "Changed Errors":  changed_errors,
        "New Materials":   new_materials,
        "Views Lost":      views_lost,
    }

    excel_bytes = io.BytesIO()
    with pd.ExcelWriter(excel_bytes, engine="openpyxl") as writer:
        for sheet, df in diff_dfs.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    excel_bytes.seek(0)

    return excel_bytes, diff_dfs
