"""
Indirect Materials (ERSA / HIBE / VERP) SAP validation engine.
Rules sourced from Indirect_Materials_SAP_Mapping.docx.

Input tables (same pattern as TN engine — no SAPUI5 Export):
  • MARA  — General Material Data
  • MARC  — Plant Data for Material
  • MVKE  — Sales Data for Material
  • MBEW  — Material Valuation
"""

import io
import re
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from sap_validation_engine import normalize_code
from sap_validation_engine import _style_sap_report
from utils.fast_loader import load_xlsx_fast

IM_MATERIAL_TYPES = frozenset({"ERSA", "HIBE", "VERP"})
IM_PLANTS = frozenset({"PT01", "PT02", "TN01", "TN02"})

ERROR_COLS = [
    "Material", "Plant", "Field", "Value", "Error", "Error_Type", "Table", "View",
]


# =============================================================================
# HELPERS
# =============================================================================

def _norm(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return str(x).strip()


def _is_blank(x):
    return _norm(x) == ""


def _is_x(x):
    v = _norm(x).upper()
    return v in {"X", "1", "YES", "TRUE"}


def _to_float(x):
    s = _norm(x)
    if not s:
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return None


def _pick(df, candidates):
    if df is None or df.empty:
        return None
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.lower()
        if key in cols:
            return cols[key]
    return None


def _prepare_df(df):
    if df is None:
        return None
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def _material_key(x):
    """
    Canonical Material key for cross-table joins/lookups (MARA <-> MARC).
    Strips leading zeros so '000000000000900022' and '900022' match,
    while '' stays '' (never becomes '0').
    """
    s = _norm(x)
    if not s:
        return ""
    # Drop a trailing '.0' from Excel-cast floats, then strip leading zeros
    if s.endswith(".0"):
        s = s[:-2]
    s = s.lstrip("0")
    return s if s else "0"


def _append_error(errors, *, material, plant, field, value, error, table=""):
    errors.append(
        {
            "Material": _norm(material),
            "Plant": _norm(plant),
            "Field": field,
            "Value": _norm(value),
            "Error": error,
            "Error_Type": "Field Error",
            "Table": table,
            "View": "",
        }
    )


def _normalize_columns(df, rename_map):
    """Rename DataFrame columns using a lowercase-keyed rename map."""
    df = df.copy()
    df.columns = [
        rename_map.get(str(col).strip().lower(), str(col).strip())
        for col in df.columns
    ]
    if df.columns.duplicated().any():
        df = df.T.groupby(level=0).first().T
    return df


def _normalize_df_columns(df, rename_map):
    if df is None:
        return df
    return _normalize_columns(df, rename_map)


# =============================================================================
# STEP 1 — Build sap_like_df from MARA + MARC (replaces SAPUI5 Export)
# =============================================================================

def build_im_sap_like_df(marc_df, mara_df):
    """
    Merges MARC + MARA into a single flat DataFrame for Indirect Materials
    validation, mirroring the column layout previously sourced from the
    SAPUI5 Export.

    Column origins:
      MARA: MATNR, MTART, MAKTX (via MAKT join in export), MEINS, SPART,
            MATKL, MTPOS, BRGEW, NTGEW, GEWEI, MAGRV, HUTYP,
            Packaging Mat. Type (field from MARA), WH Material Group
      MARC: WERKS, DISMM, DISPO, DISGR, MMSTA, DISLS, BSTMI, MABST,
            BESKZ, SOBSL, LGPRO, PERKZ, MTVFP, EKGRP, AUTPU, KORDB,
            IPRKZ (Period Ind.), MHDRZ (Expiration Date), PRCTR, XMCNG, LOGGR
    """
    marc = marc_df.copy()
    mara = mara_df.copy()
    marc.columns = [str(c).strip() for c in marc.columns]
    mara.columns = [str(c).strip() for c in mara.columns]

    # ── Normalise MARA column names ──────────────────────────────────────────
    mara_rename = {
        # MAKT (joined in export)
        'material description':               'Description',
        'maktx':                              'Description',
        # MARA fields
        'mtart':                              'Material Type',
        'material type':                      'Material Type',
        'matnr':                              'Material',
        'meins':                              'Base UOM',
        'base unit of measure':               'Base UOM',
        'base uom':                           'Base UOM',
        'spart':                              'Division',
        'division':                           'Division',
        'matkl':                              'Material Group',
        'material group':                     'Material Group',
        'mtpos':                              'General Item Category Group',
        'gen. item cat. grp':                 'General Item Category Group',
        'general item category group':        'General Item Category Group',
        'item category group':                'General Item Category Group',
        'brgew':                              'Gross Weight',
        'gross weight':                       'Gross Weight',
        'ntgew':                              'Net Weight',
        'net weight':                         'Net Weight',
        'gewei':                              'Unit of Weight (ISO Format)',
        'unit of weight':                     'Unit of Weight (ISO Format)',
        'unit of weight (iso format)':        'Unit of Weight (ISO Format)',
        'magrv':                              'Matl Grp Pack.Matls',
        'matl grp pack.matls':                'Matl Grp Pack.Matls',
        'material group: packaging materials':'Matl Grp Pack.Matls',
        'hutyp':                              'Handling Unit Type',
        'handling unit type':                 'Handling Unit Type',
        'packaging mat. type':                'Packaging Mat. Type',
        'packaging material type':            'Packaging Mat. Type',
        # WH Material Group lives in MARA (WRKST / custom field)
        'wh material group':                  'WH Material Group',
        'warehouse material group':           'WH Material Group',
    }
    mara = _normalize_columns(mara, mara_rename)

    # ── Normalise MARC column names ──────────────────────────────────────────
    marc_rename = {
        'matnr':                                    'Material',
        'werks':                                    'Plant',
        'plant':                                    'Plant',
        # MRP View 1
        'dismm':                                    'MRP Type',
        'mrp type':                                 'MRP Type',
        'dispo':                                    'MRP Controller',
        'mrp controller':                           'MRP Controller',
        'disgr':                                    'MRP Group',
        'mrp group':                                'MRP Group',
        'mmsta':                                    'Plant Specific Material Status',
        'plant-specific material status':           'Plant Specific Material Status',
        'plant specific material status':           'Plant Specific Material Status',
        'disls':                                    'Lot Size',
        'lot sizing procedure':                     'Lot Size',
        'lot size':                                 'Lot Size',
        'bstmi':                                    'Min Lot Size',
        'minimum lot size':                         'Min Lot Size',
        'min. lot size':                            'Min Lot Size',
        'min lot size':                             'Min Lot Size',
        'bstma':                                    'Maximum Lot Size',
        'maximum lot size':                         'Maximum Lot Size',
        'max. lot size':                            'Maximum Lot Size',
        'max lot size':                             'Maximum Lot Size',
        'mabst':                                    'Max Stock Level',
        'maximum stock level':                      'Max Stock Level',
        'max stock':                                'Max Stock Level',
        'max.stock level':                          'Max Stock Level',
        # MRP View 2
        'beskz':                                    'Procurement Type',
        'procurement type':                         'Procurement Type',
        'sobsl':                                    'Special Procurement',
        'special procurement':                      'Special Procurement',
        'special procurement type':                 'Special Procurement',
        'lgpro':                                    'Storage Location Ext Proc',
        'storage location for external procurement':'Storage Location Ext Proc',
        'storage location ext.proc.':               'Storage Location Ext Proc',
        'storage location ext proc':                'Storage Location Ext Proc',
        # MRP View 3
        'perkz':                                    'Period Indicator',
        'period indicator':                         'Period Indicator',
        'mtvfp':                                    'Availability Check',
        'availability check':                       'Availability Check',
        # Purchasing View
        'ekgrp':                                    'Purchasing Group',
        'purchasing group':                         'Purchasing Group',
        'autpu':                                    'Auto PO',
        'automatic purchase order allowed':         'Auto PO',
        'automatic po':                             'Auto PO',
        'kordb':                                    'Source List Required',
        'source list required':                     'Source List Required',
        'source list':                              'Source List Required',
        # Plant/Storage 1
        'iprkz':                                    'Period Ind.',
        'shelf life period ind.':                   'Period Ind.',
        'period ind.':                              'Period Ind.',
        'mhdrz':                                    'Expiration Date',
        'expiration date':                          'Expiration Date',
        'sled/bbd':                                 'Expiration Date',
        # General Plant View
        'prctr':                                    'Profit Center',
        'profit center':                            'Profit Center',
        'xmcng':                                    'Negative Stocks',
        'negative stocks allowed in plant':         'Negative Stocks',
        'negative stocks':                          'Negative Stocks',
        'loggr':                                    'Log. Handling Group',
        'log.handling group':                       'Log. Handling Group',
        'logistics handling group':                 'Log. Handling Group',
        'log. handling group':                      'Log. Handling Group',
    }
    marc = _normalize_columns(marc, marc_rename)

    # ── Select MARA columns present ──────────────────────────────────────────
    mara_cols = [c for c in [
        'Material', 'Material Type', 'Description', 'Base UOM',
        'Division', 'General Item Category Group',
        'Gross Weight', 'Net Weight', 'Unit of Weight (ISO Format)',
        'Material Group', 'Matl Grp Pack.Matls',
        'Handling Unit Type', 'Packaging Mat. Type', 'WH Material Group',
    ] if c in mara.columns]

    # ── Merge on canonical Material key (handles leading-zero mismatches,
    #    e.g. MARC '000000000000900022' vs MARA '900022') ───────────────────
    marc = marc.copy()
    mara_subset = mara[mara_cols].copy()
    marc['_MATKEY']        = marc['Material'].map(_material_key)
    mara_subset['_MATKEY'] = mara_subset['Material'].map(_material_key)
    mara_subset = mara_subset.drop(columns=['Material'])  # keep MARC's Material for display

    df = marc.merge(mara_subset, on='_MATKEY', how='left')
    df = df.drop(columns=['_MATKEY'])

    return df


# =============================================================================
# MATERIAL TYPE MAP & FILTERS
# =============================================================================

def _material_type_map(mara_df):
    """
    Material -> Material Type map, keyed by the CANONICAL material key
    (leading zeros stripped) so lookups work regardless of how each
    table formatted the Material number (e.g. '900022' vs
    '000000000000900022').
    """
    if mara_df is None or mara_df.empty:
        return {}
    mat_col = _pick(mara_df, ["Material", "MATNR"])
    mt_col  = _pick(mara_df, ["Material Type", "MTART"])
    if not mat_col or not mt_col:
        return {}
    return dict(
        zip(
            mara_df[mat_col].map(_material_key),
            mara_df[mt_col].map(lambda x: _norm(x).upper()),
        )
    )


def _filter_plant(df, plant_col, plant_filter):
    if df is None or df.empty or not plant_filter:
        return df
    pf = _norm(plant_filter).upper()
    return df[df[plant_col].map(lambda x: _norm(x).upper()) == pf].copy()


# =============================================================================
# VALIDATION FUNCTIONS
# (all operate on the merged sap_like_df — column names match the rename maps)
# =============================================================================

def im_validar_basic(sap_df, errors):
    """Basic / General Data fields (MARA + MAKT)."""
    mat_col      = _pick(sap_df, ["Material", "MATNR"])
    plant_col    = _pick(sap_df, ["Plant", "WERKS"])
    mt_col       = _pick(sap_df, ["Material Type", "MTART"])
    desc_col     = _pick(sap_df, ["Description", "Material Description", "MAKTX"])
    base_uom_col = _pick(sap_df, ["Base UOM", "Base Unit of Measure", "MEINS"])
    div_col      = _pick(sap_df, ["Division", "SPART"])
    mg_col       = _pick(sap_df, ["Material Group", "MATKL"])
    gicg_col     = _pick(sap_df, ["General Item Category Group", "Item Category Group", "MTPOS"])
    gross_col    = _pick(sap_df, ["Gross Weight", "BRGEW"])
    net_col      = _pick(sap_df, ["Net Weight", "NTGEW"])
    uow_col      = _pick(sap_df, ["Unit of Weight (ISO Format)", "Unit of Weight", "GEWEI"])
    pack_grp_col = _pick(sap_df, ["Matl Grp Pack.Matls", "Material Group: Packaging Materials", "MAGRV"])

    for idx, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = row[plant_col] if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # MATNR length — MARA
        if _is_blank(material):
            _append_error(
                errors, material=material, plant=plant, field="Material",
                value="", error="Mandatory field blank", table="MARA",
            )
        elif len(_norm(material)) > 40:
            _append_error(
                errors, material=material, plant=plant, field="Material",
                value=material, error="Material number exceeds 40 characters", table="MARA",
            )

        # MAKTX — MAKT
        if desc_col:
            desc = _norm(row[desc_col])
            if _is_blank(desc):
                _append_error(
                    errors, material=material, plant=plant, field="Description",
                    value="", error="Mandatory field blank", table="MAKT",
                )
            elif len(desc) > 40:
                _append_error(
                    errors, material=material, plant=plant, field="Description",
                    value=desc, error=f"Description exceeds 40 characters (len={len(desc)})",
                    table="MAKT",
                )

        # MEINS — MARA
        if base_uom_col and _is_blank(row[base_uom_col]):
            _append_error(
                errors, material=material, plant=plant, field="Base UOM",
                value="", error="Mandatory field blank", table="MARA",
            )

        # SPART — MARA
        if div_col:
            val = normalize_code(row[div_col], as_int=True, zero_pad=2)
            if val == "":
                _append_error(
                    errors, material=material, plant=plant, field="Division",
                    value="", error="Mandatory field blank", table="MARA",
                )
            elif val != "00":
                _append_error(
                    errors, material=material, plant=plant, field="Division",
                    value=row[div_col], error="Should be '00'", table="MARA",
                )

        # MATKL — MARA
        if mg_col:
            mg = _norm(row[mg_col]).upper()
            if _is_blank(mg):
                _append_error(
                    errors, material=material, plant=plant, field="Material Group",
                    value="", error="Mandatory field blank", table="MARA",
                )
            elif mt in {"HIBE", "ERSA"} and not re.fullmatch(r"MT\d{7}", mg):
                _append_error(
                    errors, material=material, plant=plant, field="Material Group",
                    value=mg, error="HIBE/ERSA must match pattern MTXXXXXXX (MT + 7 digits)",
                    table="MARA",
                )
            elif mt == "VERP" and not re.fullmatch(r"LF\d{7}", mg):
                _append_error(
                    errors, material=material, plant=plant, field="Material Group",
                    value=mg, error="VERP must match pattern LFXXXXXXX (LF + 7 digits)",
                    table="MARA",
                )

        # MTPOS — MARA  (only ERSA/HIBE)
        if gicg_col and mt in {"ERSA", "HIBE"}:
            val = _norm(row[gicg_col]).upper()
            if _is_blank(val):
                _append_error(
                    errors, material=material, plant=plant,
                    field="General Item Category Group",
                    value="", error="Mandatory field blank", table="MARA",
                )
            elif val != "NORM":
                _append_error(
                    errors, material=material, plant=plant,
                    field="General Item Category Group",
                    value=val, error="ERSA/HIBE must be 'NORM'", table="MARA",
                )

        # BRGEW — MARA (0 / 0.000 counts as blank)
        gross_val = _to_float(row[gross_col]) if gross_col else None
        if gross_col and (gross_val is None or gross_val == 0):
            _append_error(
                errors, material=material, plant=plant, field="Gross Weight",
                value=row[gross_col], error="Mandatory field blank or zero", table="MARA",
            )
        # NTGEW — MARA (0 / 0.000 counts as blank)
        net_val = _to_float(row[net_col]) if net_col else None
        if net_col and (net_val is None or net_val == 0):
            _append_error(
                errors, material=material, plant=plant, field="Net Weight",
                value=row[net_col], error="Mandatory field blank or zero", table="MARA",
            )
        # GEWEI — MARA (only mandatory if the material actually has a weight > 0)
        has_weight = (gross_val is not None and gross_val > 0) or (net_val is not None and net_val > 0)
        if uow_col and has_weight and _is_blank(row[uow_col]):
            _append_error(
                errors, material=material, plant=plant,
                field="Unit of Weight (ISO Format)",
                value="", error="Mandatory field blank when a weight is set", table="MARA",
            )

        # MAGRV — MARA  (only VERP)
        if pack_grp_col and mt == "VERP":
            val = _norm(row[pack_grp_col]).upper()
            if val != "ZSHP":
                _append_error(
                    errors, material=material, plant=plant,
                    field="Matl Grp Pack.Matls",
                    value=val, error="VERP must be 'ZSHP'", table="MARA",
                )


def im_validar_mrp_view1(sap_df, errors):
    """MRP View 1 — all fields from MARC."""
    mat_col      = _pick(sap_df, ["Material", "MATNR"])
    plant_col    = _pick(sap_df, ["Plant", "WERKS"])
    mt_col       = _pick(sap_df, ["Material Type", "MTART"])
    mrp_type_col = _pick(sap_df, ["MRP Type", "DISMM"])
    mrp_ctrl_col = _pick(sap_df, ["MRP Controller", "DISPO"])
    mrp_grp_col  = _pick(sap_df, ["MRP Group", "DISGR"])
    psms_col     = _pick(sap_df, ["Plant Specific Material Status", "MMSTA"])
    lot_col      = _pick(sap_df, ["Lot Size", "Lot Sizing Procedure", "DISLS"])
    min_lot_col  = _pick(sap_df, ["Min Lot Size", "Minimum Lot Size", "BSTMI"])
    max_lot_col  = _pick(sap_df, ["Maximum Lot Size", "Max Lot Size", "BSTMA"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # DISMM — MARC
        if mrp_type_col:
            val      = _norm(row[mrp_type_col]).upper()
            expected = "ND" if plant == "TN02" and mt == "VERP" else "PD"
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Type", value="",
                              error="Mandatory field blank", table="MARC")
            elif val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="MRP Type", value=val,
                              error=f"Should be '{expected}'", table="MARC")

        # DISPO — MARC
        if mrp_ctrl_col:
            val = _norm(row[mrp_ctrl_col]).upper()
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Controller", value="",
                              error="Mandatory field blank", table="MARC")
            elif mt in {"HIBE", "ERSA"} and not re.fullmatch(r"NP\d", val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Controller", value=val,
                              error="HIBE/ERSA must match NPX (NP + digit)", table="MARC")
            elif mt == "VERP" and not re.fullmatch(r"P\d{2}", val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Controller", value=val,
                              error="VERP must match PXX (P + 2 digits)", table="MARC")

        # DISGR — MARC
        if mrp_grp_col:
            val      = _norm(row[mrp_grp_col]).upper()
            expected = "ZSTO" if plant == "TN02" and mt == "VERP" else "ZIM"
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Group", value="",
                              error="Mandatory field blank", table="MARC")
            elif val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="MRP Group", value=val,
                              error=f"Should be '{expected}'", table="MARC")

        # MMSTA — MARC
        if psms_col:
            val = normalize_code(row[psms_col], as_int=True)
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="Plant Specific Material Status", value="",
                              error="Mandatory field blank", table="MARC")
            elif mt == "VERP" and val not in {"45", "70", "99"}:
                _append_error(errors, material=material, plant=plant,
                              field="Plant Specific Material Status", value=val,
                              error="VERP must be 45, 70 or 99", table="MARC")
            elif mt in {"ERSA", "HIBE"} and val not in {"80", "70", "99"}:
                _append_error(errors, material=material, plant=plant,
                              field="Plant Specific Material Status", value=val,
                              error="ERSA/HIBE must be 80, 70 or 99", table="MARC")

        # DISLS — MARC
        if lot_col:
            val = _norm(row[lot_col]).upper()
            if val != "H1":
                _append_error(errors, material=material, plant=plant,
                              field="Lot Sizing Procedure ", value=val,
                              error="Should be 'H1'", table="MARC")

        # BSTMI — MARC (mandatory for ERSA / HIBE / VERP)
        if min_lot_col and mt in {"ERSA", "HIBE", "VERP"} and _is_blank(row[min_lot_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Minimum Lot Size", value="",
                          error="Mandatory field blank for ERSA/HIBE/VERP", table="MARC")

        # BSTMA — MARC (mandatory for ERSA / HIBE / VERP)
        if max_lot_col and mt in {"ERSA", "HIBE", "VERP"} and _is_blank(row[max_lot_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Maximum Lot Size", value="",
                          error="Mandatory field blank for ERSA/HIBE/VERP", table="MARC")


def im_validar_mrp_view2(sap_df, errors):
    """MRP View 2 — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    mt_col    = _pick(sap_df, ["Material Type", "MTART"])
    mg_col    = _pick(sap_df, ["Material Group", "MATKL"])
    proc_col  = _pick(sap_df, ["Procurement Type", "BESKZ"])
    spec_col  = _pick(sap_df, ["Special Procurement", "SOBSL"])
    stor_col  = _pick(sap_df, ["Storage Location Ext Proc", "LGPRO"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""
        mg       = _norm(row[mg_col]).upper() if mg_col else ""

        # BESKZ — MARC
        if proc_col:
            val = _norm(row[proc_col]).upper()
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="Procurement Type", value="",
                              error="Mandatory field blank", table="MARC")
            elif val != "F":
                _append_error(errors, material=material, plant=plant,
                              field="Procurement Type", value=val,
                              error="Must be 'F' for ERSA/HIBE/VERP", table="MARC")

        # SOBSL — MARC  (only TN02)
        if spec_col and plant == "TN02":
            val = _norm(row[spec_col]).upper()
            if val != "Z3":
                _append_error(errors, material=material, plant=plant,
                              field="Special Procurement Type", value=val,
                              error="TN02 must be 'Z3'", table="MARC")

        # LGPRO — MARC  (Stor. Loc. for Ext. Proc.)
        # Expected values per Plant × Material Type (from mapping table):
        #   PT01: HIBE → P160 | VERP → P155 | ERSA → P160
        #   PT02: HIBE → P260 | VERP → P255 | ERSA → P260
        #   TN01: HIBE → T160 | VERP → T155 | ERSA → T160
        #   TN02: HIBE → blank | VERP → T255 | ERSA → blank
        # Special case: ERSA + Material Group MT0020012 may also use:
        #   PT01 → P100 | PT02 → P200 | TN01 → T100 | TN02 → blank
        _LGPRO_EXPECTED = {
            ("PT01", "HIBE"): "P160",
            ("PT01", "VERP"): "P155",
            ("PT01", "ERSA"): "P160",
            ("PT02", "HIBE"): "P260",
            ("PT02", "VERP"): "P255",
            ("PT02", "ERSA"): "P260",
            ("TN01", "HIBE"): "T160",
            ("TN01", "VERP"): "T155",
            ("TN01", "ERSA"): "T160",
            ("TN02", "HIBE"): "",
            ("TN02", "VERP"): "T255",
            ("TN02", "ERSA"): "",
        }
        _LGPRO_ALLOWED_FOR_MT0020012 = {
            "PT01": "P100",
            "PT02": "P200",
            "TN01": "T100",
            "TN02": "",
        }
        if stor_col and mt in {"HIBE", "ERSA", "VERP"}:
            expected_lgpro = _LGPRO_EXPECTED.get((plant, mt))
            val_lgpro = _norm(row[stor_col]).upper()
            if expected_lgpro is None:
                # Plant/MT combination not in mapping — skip
                pass
            else:
                allowed_values = {expected_lgpro} if expected_lgpro != "" else {""}
                if mt == "ERSA" and mg == "MT0020012":
                    special_value = _LGPRO_ALLOWED_FOR_MT0020012.get(plant)
                    if special_value is not None:
                        allowed_values.add(special_value)
                if "" in allowed_values:
                    if not _is_blank(val_lgpro):
                        if val_lgpro not in allowed_values:
                            allowed_values_str = ", ".join(sorted(v for v in allowed_values if v)) or "blank"
                            _append_error(errors, material=material, plant=plant,
                                          field="Storage Location for Ext. Procurement", value=val_lgpro,
                                          error=f"Must be blank or one of: {allowed_values_str}", table="MARC")
                else:
                    if val_lgpro not in allowed_values:
                        allowed_values_str = ", ".join(sorted(allowed_values))
                        _append_error(errors, material=material, plant=plant,
                                      field="Storage Location for Ext. Procurement", value=val_lgpro,
                                      error=f"Should be one of: {allowed_values_str}", table="MARC")


def im_validar_mrp_view3(sap_df, errors):
    """MRP View 3 — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    per_col   = _pick(sap_df, ["Period Indicator", "PERKZ"])
    avail_col = _pick(sap_df, ["Availability Check", "MTVFP"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""

        # PERKZ — MARC
        if per_col:
            val = _norm(row[per_col]).upper()
            if val != "M":
                _append_error(errors, material=material, plant=plant,
                              field="Period Indicator", value=val,
                              error="Should be 'M'", table="MARC")

        # MTVFP — MARC
        if avail_col:
            val = normalize_code(row[avail_col], as_int=True, zero_pad=2)
            if val != "01":
                _append_error(errors, material=material, plant=plant,
                              field="Availability Check", value=row[avail_col],
                              error="Should be '01'", table="MARC")


def im_validar_purchasing(sap_df, errors):
    """Purchasing View — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    mt_col    = _pick(sap_df, ["Material Type", "MTART"])
    pg_col    = _pick(sap_df, ["Purchasing Group", "EKGRP"])
    auto_col  = _pick(sap_df, ["Auto PO", "Automatic Purchase Order Allowed", "AUTPU"])
    src_col   = _pick(sap_df, ["Source List Required", "KORDB"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # EKGRP — MARC
        if pg_col:
            val     = _norm(row[pg_col]).upper()
            allowed = {"IMR", "IMT", "IPS"}
            if plant == "TN02" and mt == "VERP":
                if val != "STO":
                    _append_error(errors, material=material, plant=plant,
                                  field="Purchasing Group", value=val,
                                  error="TN02 VERP must be 'STO'", table="MARC")
            elif mt == "VERP":
                if val != "IFL":
                    _append_error(errors, material=material, plant=plant,
                                  field="Purchasing Group", value=val,
                                  error="VERP must be 'IFL'", table="MARC")
            elif mt in {"HIBE", "ERSA"}:
                if val not in allowed:
                    _append_error(errors, material=material, plant=plant,
                                  field="Purchasing Group", value=val,
                                  error="HIBE/ERSA must be IMR, IMT or IPS", table="MARC")

        # AUTPU — MARC
        if auto_col and not _is_x(row[auto_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Automatic Purchase Order Allowed", value=row[auto_col],
                          error="Must be 'X'", table="MARC")

        # KORDB — MARC
        if src_col and not _is_x(row[src_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Source List Required", value=row[src_col],
                          error="Must be 'X'", table="MARC")


def im_validar_plant_storage(sap_df, errors):
    """Plant Data / Storage 1 — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    mt_col    = _pick(sap_df, ["Material Type", "MTART"])
    per_col   = _pick(sap_df, ["Period Ind.", "IPRKZ"])
    exp_col   = _pick(sap_df, ["Expiration Date", "MHDRZ"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # IPRKZ — MARC (Period Ind. for SLED — mandatory, must be "D" for ERSA/HIBE/VERP)
        if per_col and mt in {"ERSA", "HIBE", "VERP"}:
            val = _norm(row[per_col]).upper() 
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="Period Ind. for SLED", value="",
                              error="Mandatory field blank for ERSA/HIBE/VERP", table="MARC")
            elif val != "D":
                _append_error(errors, material=material, plant=plant,
                              field="Period Ind. for SLED", value=val,
                              error="Should be 'D'", table="MARC")

        # MHDRZ — MARC (Expiration Date — must be blank for ERSA/HIBE/VERP)
        if exp_col and mt in {"ERSA", "HIBE", "VERP"} and not _is_blank(row[exp_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Expiration Date", value=row[exp_col],
                          error="Must be blank for ERSA/HIBE/VERP", table="MARC")


def im_validar_general_plant(sap_df, errors):
    """General Plant Data — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    mt_col    = _pick(sap_df, ["Material Type", "MTART"])
    pc_col    = _pick(sap_df, ["Profit Center", "PRCTR"])
    neg_col   = _pick(sap_df, ["Negative Stocks", "XMCNG"])
    log_col   = _pick(sap_df, ["Log. Handling Group", "LOGGR"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # PRCTR — MARC
        if pc_col:
            val      = _norm(row[pc_col]).upper()
            expected = {
                "PT01": "PT00351",
                "PT02": "PT00361",
                "TN01": "TN00413",
                "TN02": "TN00413",
            }.get(plant)
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="Profit Center", value="",
                              error="Mandatory field blank", table="MARC")
            elif expected and val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="Profit Center", value=val,
                              error=f"Should be '{expected}'", table="MARC")

        # XMCNG — MARC
        if neg_col:
            is_x = _is_x(row[neg_col])
            if mt == "VERP" and not is_x:
                _append_error(errors, material=material, plant=plant,
                              field="Neg. stocks in plant", value=row[neg_col],
                              error="VERP must be 'X'", table="MARC")
            elif mt != "VERP" and is_x:
                _append_error(errors, material=material, plant=plant,
                              field="Neg. stocks in plant", value=row[neg_col],
                              error="Only VERP may be 'X'", table="MARC")

        # LOGGR — MARC
        if log_col:
            val = _norm(row[log_col]).upper()
            if val != "NORM":
                _append_error(errors, material=material, plant=plant,
                              field="Log. Handling Group", value=val,
                              error="Should be 'NORM'", table="MARC")


def im_validar_mbew(mbew_df, mt_map, errors, plant_filter=None):
    """Accounting / Valuation — fields from MBEW."""
    if mbew_df is None or mbew_df.empty:
        return
    mat_col  = _pick(mbew_df, ["Material", "MATNR"])
    area_col = _pick(mbew_df, ["Plant", "Valuation Area", "BWKEY", "WERKS"])
    vc_col   = _pick(mbew_df, ["Valuation Class", "BKLAS"])
    pd_col   = _pick(mbew_df, ["Price Determination", "Material Price Determination: Control", "MLAST"])
    pc_col   = _pick(mbew_df, ["Price Control", "VPRSV"])
    pu_col   = _pick(mbew_df, ["Price Unit", "PEINH"])
    lifo_col = _pick(mbew_df, ["LIFO/FIFO-relevant", "XLIFO"])

    for _, row in mbew_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[area_col]).upper() if area_col else ""
        mt       = mt_map.get(_material_key(material), "")
        if mt not in IM_MATERIAL_TYPES:
            continue
        if plant_filter and plant != _norm(plant_filter).upper():
            continue

        # BKLAS — MBEW
        if vc_col:
            val      = normalize_code(row[vc_col], as_int=True)
            expected = {"HIBE": "3030", "ERSA": "3040"}.get(mt)
            if expected and val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="Valuation Class", value=val,
                              error=f"{mt} must be '{expected}'", table="MBEW")

        # MLAST — MBEW  (ERSA/HIBE only)
        if pd_col and mt in {"ERSA", "HIBE"}:
            val = normalize_code(row[pd_col], as_int=True)
            if val != "2":
                _append_error(errors, material=material, plant=plant,
                              field="Price Control Determination", value=row[pd_col],
                              error="ERSA/HIBE must be '2'", table="MBEW")

        # VPRSV — MBEW  (ERSA/HIBE only)
        if pc_col and mt in {"ERSA", "HIBE"}:
            val = _norm(row[pc_col]).upper()
            if val != "V":
                _append_error(errors, material=material, plant=plant,
                              field="Price Control", value=val,
                              error="ERSA/HIBE must be 'V'", table="MBEW")

        # PEINH — MBEW
        if pu_col:
            val = normalize_code(row[pu_col], as_int=True)
            if val != "10000":
                _append_error(errors, material=material, plant=plant,
                              field="Price Unit", value=row[pu_col],
                              error="Must be '10000'", table="MBEW")

        # XLIFO — MBEW
        if lifo_col and not _is_x(row[lifo_col]):
            _append_error(errors, material=material, plant=plant,
                          field="LIFO/FIFO-relevant", value=row[lifo_col],
                          error="Must be 'X'", table="MBEW")


def im_validar_mara_wh_packaging(mara_df, mt_map, errors):
    """WH & Packaging fields — all from MARA."""
    if mara_df is None or mara_df.empty:
        return
    mat_col = _pick(mara_df, ["Material", "MATNR"])
    mg_col  = _pick(mara_df, ["Material Group", "MATKL"])
    wh_col  = _pick(mara_df, ["WH Material Group", "Warehouse Material Group"])
    hut_col = _pick(mara_df, ["Handling Unit Type", "HUTYP"])
    pkg_col = _pick(mara_df, ["Packaging Mat. Type", "Packaging material type"])
    ms_col  = _pick(mara_df, ["Maintenance status", "Maintenance Status", "PSTAT"])

    for _, row in mara_df.iterrows():
        material = row[mat_col] if mat_col else ""
        mt       = mt_map.get(_material_key(material), "")
        if mt not in IM_MATERIAL_TYPES:
            continue
        mg = _norm(row[mg_col]).upper() if mg_col else ""

        # WH Material Group — MARA  (ERSA + specific MG only) — must be blank
        if wh_col and mt == "ERSA" and mg == "MT0020012":
            val = _norm(row[wh_col]).upper()
            if not _is_blank(val):
                _append_error(errors, material=material, plant="",
                              field="WH Material Group", value=val,
                              error="ERSA MT0020012 must be blank", table="MARA")

        # Maintenance status — MARA  (ERSA + MT0020012 must contain letter 'C')
        if ms_col and mt == "ERSA" and mg == "MT0020012":
            ms_val = _norm(row[ms_col]).upper()
            if "C" not in ms_val:
                _append_error(errors, material=material, plant="",
                              field="Maintenance Status (MARA)", value=ms_val,
                              error="ERSA MT0020012 must have letter 'C' in Maintenance status",
                              table="MARA")

        # HUTYP — MARA  (mandatory for VERP)
        if hut_col and mt == "VERP" and _is_blank(row[hut_col]):
            _append_error(errors, material=material, plant="",
                          field="Handling Unit Type", value="",
                          error="VERP cannot be blank", table="MARA")

        # Packaging Mat. Type — MARA  (mandatory for VERP)
        if pkg_col and mt == "VERP" and _is_blank(row[pkg_col]):
            _append_error(errors, material=material, plant="",
                          field="Packaging Mat. Type", value="",
                          error="Mandatory for VERP", table="MARA")


def im_validar_mvke(mvke_df, mt_map, errors):
    """Sales / Distribution fields — MVKE."""
    if mvke_df is None or mvke_df.empty:
        return
    mat_col = _pick(mvke_df, ["Material", "MATNR"])
    icg_col = _pick(mvke_df, ["Item Category Group", "MTPOS"])
    tax_cols = [
        c for c in mvke_df.columns
        if "tax" in str(c).lower() or "TAX" in str(c)
    ]

    for _, row in mvke_df.iterrows():
        material = row[mat_col] if mat_col else ""
        mt       = mt_map.get(_material_key(material), "")
        if mt != "VERP":
            continue

        # MTPOS — MVKE
        if icg_col:
            val = _norm(row[icg_col]).upper()
            if val != "NORM":
                _append_error(errors, material=material, plant="",
                              field="Item Category Group (MVKE)", value=val,
                              error="VERP must be 'NORM'", table="MVKE")

        # Tax data — MVKE
        if tax_cols:
            has_tax = any(not _is_blank(row[c]) for c in tax_cols)
            if not has_tax:
                _append_error(errors, material=material, plant="",
                              field="Tax Data (MVKE)", value="",
                              error="Tax data required for VERP", table="MVKE")


# =============================================================================
# SUMMARY & STYLING
# =============================================================================

# IM-specific Field → View mapping (aligned with RolesViews.xlsx)
# Fields not in the Roles sheet are mapped to "Not applicable"
_IM_FIELD_TO_VIEW = {
    # Basic Data (K)
    "Material":                             "K",
    "Description":                          "K",
    "Material Type":                        "K",
    "Base UOM":                             "K",
    "Division":                             "K",
    "Material Group":                       "K",
    "General Item Category Group":          "K",
    "Gross Weight":                         "K",
    "Net Weight":                           "K",
    "Unit of Weight (ISO Format)":          "K",
    "Matl Grp Pack.Matls":                  "K",
    # MRP / Purchasing (D / E)
    "MRP Type":                             "D",
    "MRP Controller":                       "D",
    "MRP Group":                            "D",
    "Plant Specific Material Status":       "D",
    "Lot Sizing Procedure ":                "D",
    "Minimum Lot Size":                     "D",
    "Maximum Lot Size":                     "D",
    "Procurement Type":                     "D",
    "Special Procurement Type":             "D",
    "Storage Location for Ext. Procurement":"D",
    "Period Indicator":                     "D",
    "Availability Check":                   "D",
    "Purchasing Group":                     "E",
    "Automatic Purchase Order Allowed":     "E",
    "Source List Required":                 "E",
    # Storage / General Plant (L)
    "Period Ind. for SLED":                 "L",
    "Expiration Date":                      "L",
    "Profit Center":                        "L",
    "Neg. stocks in plant":                 "L",
    "Log. Handling Group":                  "L",
    # Accounting (B)
    "Valuation Class":                      "B",
    "Price Control Determination":          "B",
    "Price Control":                        "B",
    "Price Unit":                           "B",
    "LIFO/FIFO-relevant":                   "B",
    # Sales (V)
    "Packaging Mat. Type":                  "V",
    "Item Category Group (MVKE)":           "V",
    "Tax Data (MVKE)":                      "V",
    # Warehouse Management (S) — only conditionally required (see _im_postprocess)
    "WH Material Group":                    "S",
    "Maintenance Status (MARA)":            "C",
    # No view
    "Handling Unit Type":                   "Not applicable",
}

_IM_VIEW_TO_DESC = {
    "K": "Basic Data",
    "D": "MRP",
    "E": "Purchasing",
    "L": "Storage",
    "B": "Accounting",
    "V": "Sales",
    "S": "Warehouse management",
    "C": "Classification",
}

# Required views per IM material type
# Note: "B" (Accounting) is NOT required for VERP — packaging materials
# are not accounting-relevant the same way, so it's excluded from scope.
_IM_REQUIRED_VIEWS = {
    "ERSA": {"K", "D", "E", "L", "B"},
    "HIBE": {"K", "D", "E", "L", "B"},
    "VERP": {"K", "D", "E", "L", "V"},
}

# All fields expected per view (for fields_expected column in view_tabs)
_IM_VIEW_TO_FIELDS: dict[str, list[str]] = {}
for _f, _v in _IM_FIELD_TO_VIEW.items():
    if _v != "Not applicable":
        _IM_VIEW_TO_FIELDS.setdefault(_v, []).append(_f)


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


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

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