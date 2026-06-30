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
    "Material", "Plant", "Field", "Value", "Error", "Error_Type", "Table",
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

    # ── Merge ────────────────────────────────────────────────────────────────
    df = marc.merge(mara[mara_cols], on='Material', how='left')

    return df


# =============================================================================
# MATERIAL TYPE MAP & FILTERS
# =============================================================================

def _material_type_map(mara_df):
    if mara_df is None or mara_df.empty:
        return {}
    mat_col = _pick(mara_df, ["Material", "MATNR"])
    mt_col  = _pick(mara_df, ["Material Type", "MTART"])
    if not mat_col or not mt_col:
        return {}
    return dict(
        zip(
            mara_df[mat_col].map(_norm),
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
                    errors, material=material, plant=plant, field="Description (MAKTX)",
                    value="", error="Mandatory field blank", table="MAKT",
                )
            elif len(desc) > 40:
                _append_error(
                    errors, material=material, plant=plant, field="Description (MAKTX)",
                    value=desc, error=f"Description exceeds 40 characters (len={len(desc)})",
                    table="MAKT",
                )

        # MEINS — MARA
        if base_uom_col and _is_blank(row[base_uom_col]):
            _append_error(
                errors, material=material, plant=plant, field="Base UOM (MEINS)",
                value="", error="Mandatory field blank", table="MARA",
            )

        # SPART — MARA
        if div_col:
            val = normalize_code(row[div_col], as_int=True, zero_pad=2)
            if val == "":
                _append_error(
                    errors, material=material, plant=plant, field="Division (SPART)",
                    value="", error="Mandatory field blank", table="MARA",
                )
            elif val != "00":
                _append_error(
                    errors, material=material, plant=plant, field="Division (SPART)",
                    value=row[div_col], error="Should be '00'", table="MARA",
                )

        # MATKL — MARA
        if mg_col:
            mg = _norm(row[mg_col]).upper()
            if _is_blank(mg):
                _append_error(
                    errors, material=material, plant=plant, field="Material Group (MATKL)",
                    value="", error="Mandatory field blank", table="MARA",
                )
            elif mt in {"HIBE", "ERSA"} and not re.fullmatch(r"MT\d{7}", mg):
                _append_error(
                    errors, material=material, plant=plant, field="Material Group (MATKL)",
                    value=mg, error="HIBE/ERSA must match pattern MTXXXXXXX (MT + 7 digits)",
                    table="MARA",
                )
            elif mt == "VERP" and not re.fullmatch(r"LF\d{7}", mg):
                _append_error(
                    errors, material=material, plant=plant, field="Material Group (MATKL)",
                    value=mg, error="VERP must match pattern LFXXXXXXX (LF + 7 digits)",
                    table="MARA",
                )

        # MTPOS — MARA  (only ERSA/HIBE)
        if gicg_col and mt in {"ERSA", "HIBE"}:
            val = _norm(row[gicg_col]).upper()
            if _is_blank(val):
                _append_error(
                    errors, material=material, plant=plant,
                    field="Gen. Item Cat. Grp (MTPOS)",
                    value="", error="Mandatory field blank", table="MARA",
                )
            elif val != "NORM":
                _append_error(
                    errors, material=material, plant=plant,
                    field="Gen. Item Cat. Grp (MTPOS)",
                    value=val, error="ERSA/HIBE must be 'NORM'", table="MARA",
                )

        # BRGEW — MARA
        if gross_col and _to_float(row[gross_col]) is None:
            _append_error(
                errors, material=material, plant=plant, field="Gross Weight (BRGEW)",
                value=row[gross_col], error="Mandatory field blank or zero", table="MARA",
            )
        # NTGEW — MARA
        if net_col and _to_float(row[net_col]) is None:
            _append_error(
                errors, material=material, plant=plant, field="Net Weight (NTGEW)",
                value=row[net_col], error="Mandatory field blank or zero", table="MARA",
            )
        # GEWEI — MARA
        if uow_col and _is_blank(row[uow_col]):
            _append_error(
                errors, material=material, plant=plant,
                field="Unit of Weight (GEWEI)",
                value="", error="Mandatory field blank", table="MARA",
            )

        # MAGRV — MARA  (only VERP)
        if pack_grp_col and mt == "VERP":
            val = _norm(row[pack_grp_col]).upper()
            if val != "ZSHP":
                _append_error(
                    errors, material=material, plant=plant,
                    field="Matl Grp Pack.Matls (MAGRV)",
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
                              field="MRP Type (DISMM)", value="",
                              error="Mandatory field blank", table="MARC")
            elif val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="MRP Type (DISMM)", value=val,
                              error=f"Should be '{expected}'", table="MARC")

        # DISPO — MARC
        if mrp_ctrl_col:
            val = _norm(row[mrp_ctrl_col]).upper()
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Controller (DISPO)", value="",
                              error="Mandatory field blank", table="MARC")
            elif mt in {"HIBE", "ERSA"} and not re.fullmatch(r"NP\d", val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Controller (DISPO)", value=val,
                              error="HIBE/ERSA must match NPX (NP + digit)", table="MARC")
            elif mt == "VERP" and not re.fullmatch(r"P\d{2}", val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Controller (DISPO)", value=val,
                              error="VERP must match PXX (P + 2 digits)", table="MARC")

        # DISGR — MARC
        if mrp_grp_col:
            val      = _norm(row[mrp_grp_col]).upper()
            expected = "ZSTO" if plant == "TN02" and mt == "VERP" else "ZIM"
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Group (DISGR)", value="",
                              error="Mandatory field blank", table="MARC")
            elif val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="MRP Group (DISGR)", value=val,
                              error=f"Should be '{expected}'", table="MARC")

        # MMSTA — MARC
        if psms_col:
            val = normalize_code(row[psms_col], as_int=True)
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="Plant Specific Material Status (MMSTA)", value="",
                              error="Mandatory field blank", table="MARC")
            elif mt == "VERP" and val not in {"45", "70", "99"}:
                _append_error(errors, material=material, plant=plant,
                              field="Plant Specific Material Status (MMSTA)", value=val,
                              error="VERP must be 45, 70 or 99", table="MARC")
            elif mt in {"ERSA", "HIBE"} and val != "80":
                _append_error(errors, material=material, plant=plant,
                              field="Plant Specific Material Status (MMSTA)", value=val,
                              error="ERSA/HIBE must be '80'", table="MARC")

        # DISLS — MARC
        if lot_col:
            val = _norm(row[lot_col]).upper()
            if val != "H1":
                _append_error(errors, material=material, plant=plant,
                              field="Lot Size (DISLS)", value=val,
                              error="Should be 'H1'", table="MARC")

        # BSTMI — MARC (mandatory for ERSA / HIBE / VERP)
        if min_lot_col and mt in {"ERSA", "HIBE", "VERP"} and _is_blank(row[min_lot_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Minimum Lot Size (BSTMI)", value="",
                          error="Mandatory field blank for ERSA/HIBE/VERP", table="MARC")

        # BSTMA — MARC (mandatory for ERSA / HIBE / VERP)
        if max_lot_col and mt in {"ERSA", "HIBE", "VERP"} and _is_blank(row[max_lot_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Maximum Lot Size (BSTMA)", value="",
                          error="Mandatory field blank for ERSA/HIBE/VERP", table="MARC")


def im_validar_mrp_view2(sap_df, errors):
    """MRP View 2 — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    mt_col    = _pick(sap_df, ["Material Type", "MTART"])
    proc_col  = _pick(sap_df, ["Procurement Type", "BESKZ"])
    spec_col  = _pick(sap_df, ["Special Procurement", "SOBSL"])
    stor_col  = _pick(sap_df, ["Storage Location Ext Proc", "LGPRO"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # BESKZ — MARC
        if proc_col:
            val = _norm(row[proc_col]).upper()
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="Procurement Type (BESKZ)", value="",
                              error="Mandatory field blank", table="MARC")
            elif val not in {"F", "X"}:
                _append_error(errors, material=material, plant=plant,
                              field="Procurement Type (BESKZ)", value=val,
                              error="Must be 'F' or 'X'", table="MARC")

        # SOBSL — MARC  (only TN02)
        if spec_col and plant == "TN02":
            val = _norm(row[spec_col]).upper()
            if val != "Z3":
                _append_error(errors, material=material, plant=plant,
                              field="Special Procurement (SOBSL)", value=val,
                              error="TN02 must be 'Z3'", table="MARC")

        # LGPRO — MARC  (Stor. Loc. for Ext. Proc.)
        # Expected values per Plant × Material Type (from mapping table):
        #   PT01: HIBE → P160 | VERP → P155 | ERSA → P160
        #   PT02: HIBE → P260 | VERP → P255 | ERSA → P260
        #   TN01: HIBE → T160 | VERP → T155 | ERSA → T160
        #   TN02: HIBE → blank | VERP → T255 | ERSA → blank
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
        if stor_col and mt in {"HIBE", "ERSA", "VERP"}:
            expected_lgpro = _LGPRO_EXPECTED.get((plant, mt))
            val_lgpro = _norm(row[stor_col]).upper()
            if expected_lgpro is None:
                # Plant/MT combination not in mapping — skip
                pass
            elif expected_lgpro == "":
                # Dash in mapping → field must be blank
                if not _is_blank(val_lgpro):
                    _append_error(errors, material=material, plant=plant,
                                  field="Storage Location Ext Proc (LGPRO)", value=val_lgpro,
                                  error="Must be blank for this Plant/Material Type", table="MARC")
            else:
                # Specific value expected
                if val_lgpro != expected_lgpro:
                    _append_error(errors, material=material, plant=plant,
                                  field="Storage Location Ext Proc (LGPRO)", value=val_lgpro,
                                  error=f"Should be '{expected_lgpro}'", table="MARC")


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
                              field="Period Indicator (PERKZ)", value=val,
                              error="Should be 'M'", table="MARC")

        # MTVFP — MARC
        if avail_col:
            val = normalize_code(row[avail_col], as_int=True, zero_pad=2)
            if val != "01":
                _append_error(errors, material=material, plant=plant,
                              field="Availability Check (MTVFP)", value=row[avail_col],
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
                                  field="Purchasing Group (EKGRP)", value=val,
                                  error="TN02 VERP must be 'STO'", table="MARC")
            elif mt == "VERP":
                if val != "IFL":
                    _append_error(errors, material=material, plant=plant,
                                  field="Purchasing Group (EKGRP)", value=val,
                                  error="VERP must be 'IFL'", table="MARC")
            elif mt in {"HIBE", "ERSA"}:
                if val not in allowed:
                    _append_error(errors, material=material, plant=plant,
                                  field="Purchasing Group (EKGRP)", value=val,
                                  error="HIBE/ERSA must be IMR, IMT or IPS", table="MARC")

        # AUTPU — MARC
        if auto_col and not _is_x(row[auto_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Auto PO (AUTPU)", value=row[auto_col],
                          error="Must be 'X'", table="MARC")

        # KORDB — MARC
        if src_col and not _is_x(row[src_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Source List Required (KORDB)", value=row[src_col],
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
                              field="Period Ind. for SLED (IPRKZ)", value="",
                              error="Mandatory field blank for ERSA/HIBE/VERP", table="MARC")
            elif val != "D":
                _append_error(errors, material=material, plant=plant,
                              field="Period Ind. for SLED (IPRKZ)", value=val,
                              error="Should be 'D'", table="MARC")

        # MHDRZ — MARC (Expiration Date — must be blank for ERSA/HIBE/VERP)
        if exp_col and mt in {"ERSA", "HIBE", "VERP"} and not _is_blank(row[exp_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Expiration Date (MHDRZ)", value=row[exp_col],
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
                              field="Profit Center (PRCTR)", value="",
                              error="Mandatory field blank", table="MARC")
            elif expected and val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="Profit Center (PRCTR)", value=val,
                              error=f"Should be '{expected}'", table="MARC")

        # XMCNG — MARC
        if neg_col:
            is_x = _is_x(row[neg_col])
            if mt == "VERP" and not is_x:
                _append_error(errors, material=material, plant=plant,
                              field="Negative Stocks (XMCNG)", value=row[neg_col],
                              error="VERP must be 'X'", table="MARC")
            elif mt != "VERP" and is_x:
                _append_error(errors, material=material, plant=plant,
                              field="Negative Stocks (XMCNG)", value=row[neg_col],
                              error="Only VERP may be 'X'", table="MARC")

        # LOGGR — MARC
        if log_col:
            val = _norm(row[log_col]).upper()
            if val != "NORM":
                _append_error(errors, material=material, plant=plant,
                              field="Log. Handling Group (LOGGR)", value=val,
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
        mt       = mt_map.get(_norm(material), "")
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
                              field="Valuation Class (BKLAS)", value=val,
                              error=f"{mt} must be '{expected}'", table="MBEW")

        # MLAST — MBEW  (ERSA/HIBE only)
        if pd_col and mt in {"ERSA", "HIBE"}:
            val = normalize_code(row[pd_col], as_int=True)
            if val != "2":
                _append_error(errors, material=material, plant=plant,
                              field="Price Determination (MLAST)", value=row[pd_col],
                              error="ERSA/HIBE must be '2'", table="MBEW")

        # VPRSV — MBEW  (ERSA/HIBE only)
        if pc_col and mt in {"ERSA", "HIBE"}:
            val = _norm(row[pc_col]).upper()
            if val != "V":
                _append_error(errors, material=material, plant=plant,
                              field="Price Control (VPRSV)", value=val,
                              error="ERSA/HIBE must be 'V'", table="MBEW")

        # PEINH — MBEW
        if pu_col:
            val = normalize_code(row[pu_col], as_int=True)
            if val != "10000":
                _append_error(errors, material=material, plant=plant,
                              field="Price Unit (PEINH)", value=row[pu_col],
                              error="Must be '10000'", table="MBEW")

        # XLIFO — MBEW
        if lifo_col and not _is_x(row[lifo_col]):
            _append_error(errors, material=material, plant=plant,
                          field="LIFO/FIFO-relevant (XLIFO)", value=row[lifo_col],
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

    for _, row in mara_df.iterrows():
        material = row[mat_col] if mat_col else ""
        mt       = mt_map.get(_norm(material), "")
        if mt not in IM_MATERIAL_TYPES:
            continue
        mg = _norm(row[mg_col]).upper() if mg_col else ""

        # WH Material Group — MARA  (ERSA + specific MG only)
        if wh_col and mt == "ERSA" and mg == "MT0020012":
            val = _norm(row[wh_col]).upper()
            if val != "Z006":
                _append_error(errors, material=material, plant="",
                              field="WH Material Group (MARA)", value=val,
                              error="ERSA MT0020012 must be 'Z006'", table="MARA")

        # HUTYP — MARA  (mandatory for VERP)
        if hut_col and mt == "VERP" and _is_blank(row[hut_col]):
            _append_error(errors, material=material, plant="",
                          field="Handling Unit Type (HUTYP)", value="",
                          error="VERP cannot be blank", table="MARA")

        # Packaging Mat. Type — MARA  (mandatory for VERP)
        if pkg_col and mt == "VERP" and _is_blank(row[pkg_col]):
            _append_error(errors, material=material, plant="",
                          field="Packaging Mat. Type (MARA)", value="",
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
        mt       = mt_map.get(_norm(material), "")
        if mt != "VERP":
            continue

        # MTPOS — MVKE
        if icg_col:
            val = _norm(row[icg_col]).upper()
            if val != "NORM":
                _append_error(errors, material=material, plant="",
                              field="Item Category Group (MTPOS / MVKE)", value=val,
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

def _build_summary(errors_df):
    if errors_df.empty or "Field" not in errors_df.columns:
        return pd.DataFrame(columns=["Field", "Adjusted Nº of Errors", "% of Total"])
    summary = (
        errors_df["Field"]
        .value_counts()
        .rename_axis("Field")
        .reset_index(name="Adjusted Nº of Errors")
    )
    total = summary["Adjusted Nº of Errors"].sum()
    if total:
        summary["% of Total"] = (
            summary["Adjusted Nº of Errors"] / total * 100
        ).round(1)
    else:
        summary["% of Total"] = 0.0
    return summary.sort_values("Adjusted Nº of Errors", ascending=False).reset_index(drop=True)


def _style_im_report(
    xlsx_bytes: io.BytesIO,
    *,
    errors_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    coverage_df: pd.DataFrame | None = None,
) -> io.BytesIO:

    return _style_sap_report(
        xlsx_bytes,
        errors_df=errors_df,
        contagem_final_df=summary_df,
        coverage_df=coverage_df,
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

    summary_df  = _build_summary(errors_df)
    coverage_df = pd.DataFrame(
        {
            "Metric": [
                "Materials in scope (MARC)",
                "Materials with errors",
                "Distinct fields with errors",
                "Total errors",
            ],
            "Value": [
                int(sap_im[mat_col_sl].nunique()),
                int(errors_df["Material"].nunique()) if not errors_df.empty else 0,
                int(errors_df["Field"].nunique())    if not errors_df.empty else 0,
                len(errors_df),
            ],
        }
    )

    _cb(95, "Building Excel report...")
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, sheet_name="Executive Summary", index=False)
        summary_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=1)
        if not coverage_df.empty:
            coverage_df.to_excel(
                writer, sheet_name="Adjusted Summary", index=False,
                startrow=len(summary_df) + 4,
            )
        errors_df.to_excel(writer, sheet_name="Field With Error", index=False)
        errors_df.to_excel(writer, sheet_name="All Errors",       index=False)

    out.seek(0)
    styled = _style_im_report(
        out,
        errors_df=errors_df,
        summary_df=summary_df,
        coverage_df=coverage_df,
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
        "run_type":    "INDIRECT_MATERIALS",
        "timestamp":   datetime.now().isoformat(timespec="seconds"),
        "user":        user,
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