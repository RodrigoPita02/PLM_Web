"""Helpers e construção do DataFrame SAP-like do Indirect Materials."""

import pandas as pd


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
