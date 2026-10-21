# =============================================================================
# SAP VALIDATION ENGINE — PART 1: IMPORTS, LOADERS & HELPERS
# =============================================================================

import io
import os
import json
import re
import math
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint

from utils.fast_loader import load_xlsx_fast
from rolesviews_data import get_rolesviews_bytesio as _get_rolesviews_bytesio_raw


def get_rolesviews_bytesio():
    """
    Defensive wrapper around rolesviews_data.get_rolesviews_bytesio().

    If the underlying implementation caches/reuses a single BytesIO object
    across calls, its read cursor may already be at EOF from an earlier
    read elsewhere in the same process (e.g. a UI preview, or another
    engine run in the same combined "All Sites" execution). openpyxl then
    fails with "File is not a zip file" because it reads zero bytes.
    Rewinding to position 0 here makes every call safe to read from
    scratch, regardless of what happened to the object before.
    """
    buf = _get_rolesviews_bytesio_raw()
    if hasattr(buf, "seek"):
        try:
            buf.seek(0)
        except Exception:
            pass
    return buf


# =============================================================================
# 1) UNIVERSAL FILE LOADER
# =============================================================================

def load_xlsx(file, *, sheet_name=None):
    """
    Loads an Excel file from multiple source types:
    - str (file path on disk)
    - Streamlit UploadedFile (uses .getvalue())
    - Generic file-like with .read()
    - io.BytesIO
    """
    if file is None:
        return None

    if isinstance(file, str):
        return pd.read_excel(file, engine="openpyxl", sheet_name=sheet_name)

    if hasattr(file, "getvalue"):
        bio = io.BytesIO(file.getvalue())
        return pd.read_excel(bio, engine="openpyxl", sheet_name=sheet_name)

    if hasattr(file, "read"):
        try:
            file.seek(0)
        except Exception:
            pass
        bio = io.BytesIO(file.read())
        return pd.read_excel(bio, engine="openpyxl", sheet_name=sheet_name)

    if isinstance(file, io.BytesIO):
        file.seek(0)
        return pd.read_excel(file, engine="openpyxl", sheet_name=sheet_name)

    raise ValueError(f"Unsupported file type: {type(file)}")

# =============================================================================
# RULE-SPECIFIC MATERIAL TYPE OVERRIDES
# Default = FERT
# Exceptions = HALB (only for listed rules)
# =============================================================================
RULE_MATERIAL_TYPE_OVERRIDES = {
    "VPRMKF-10849-LB": {
        "__default__": "FERT",
        "validar_enterproj_number": "HALB",
        "validar_material_group": "HALB",
        "validar_valuation_class": "HALB",
    }
}

# =============================================================================
# 2) GENERAL HELPERS
# =============================================================================

def normalize_code(value, *, as_int=False, zero_pad=None):
    """
    Normalizes a code value from Excel:
    - Strips whitespace.
    - Converts numeric strings (e.g. '2.0', '40.0') to integer strings.
    - Applies zero-padding if specified (e.g. zero_pad=4 -> '1' -> '0001').
    Returns a normalized string.
    """
    if pd.isna(value):
        return ''
    s = str(value).strip()

    try:
        num = float(s)
        if as_int or s.endswith('.0') or s.isdigit():
            s = str(int(num))
    except ValueError:
        pass

    if zero_pad:
        try:
            s = str(int(s)).zfill(zero_pad)
        except ValueError:
            if len(s) < zero_pad and s.isdigit():
                s = s.zfill(zero_pad)

    return s


def ensure_basic_data_k(marc_df):
    """
    Ensures 'Maintenance status' contains the letter 'K' (Basic Data) for all materials.
    - If blank/NaN, sets to 'K'.
    - If 'K' already present (even duplicated), keeps only one occurrence.
    - Appends 'K' at the end if not present.
    """
    col = 'Maintenance status'
    if col not in marc_df.columns:
        print(f"⚠️ MARC is missing column '{col}'.")
        return marc_df

    def add_k(s):
        s = '' if pd.isna(s) else str(s).strip()
        s = re.sub(r'K+', 'K', s)
        if 'K' in s:
            return s
        return 'K' if s == '' else (s + 'K')

    marc_df[col] = marc_df[col].apply(add_k)
    return marc_df


# =============================================================================
# SHARED HELPERS — sap_like_df construction (used by BOTH Palmela and Tunisia)
# =============================================================================
# These used to live only in tn_validation_engine.py, since Tunisia has no
# Item Browse Report / SAPUI5 Export and had to build an equivalent flat
# dataframe from MARC + MARA + ZPLM MARC instead. The Palmela engine now
# does the same (see run_full_validation below), so these helpers moved
# here to be the single shared source of truth for both engines.

def _normalize_columns(df, rename_map):
    """Case-insensitive column rename (keys in rename_map are matched lowercased)."""
    df = df.copy()
    df.columns = [
        rename_map.get(str(col).strip().lower(), str(col).strip())
        for col in df.columns
    ]
    if df.columns.duplicated().any():
        df = df.T.groupby(level=0).first().T
    return df


def _normalize_df_columns(df, rename_map):
    """Same as _normalize_columns but tolerates df=None (used on optional tables)."""
    if df is None:
        return df
    df = df.copy()
    df.columns = [
        rename_map.get(str(col).strip().lower(), str(col).strip())
        for col in df.columns
    ]
    if df.columns.duplicated().any():
        df = df.T.groupby(level=0).first().T
    return df


def _mt_map_from_mara(mara_df):
    return dict(zip(
        mara_df['Material'].astype(str).str.strip(),
        mara_df['Material Type'].astype(str).str.strip().str.upper()
    ))


def _pt_map_from_zplm(zplm_marc_df):
    if zplm_marc_df is None or zplm_marc_df.empty:
        return {}
    z = zplm_marc_df.copy()
    z.columns = [str(c).strip() for c in z.columns]
    for variant in ['Material Plant Type', 'MATERIAL PLANT TYPE']:
        if variant in z.columns and 'Plant Type' not in z.columns:
            z.rename(columns={variant: 'Plant Type'}, inplace=True)
    if not {'Material', 'Plant', 'Plant Type'}.issubset(z.columns):
        return {}
    return {
        (str(r['Material']).strip(), str(r['Plant']).strip().upper()):
        str(r['Plant Type']).strip().upper()
        for _, r in z.iterrows()
    }


def _filter_valid_material_types(df, mara_df):
    """
    Keep only materials with Material Type in {'FERT', 'HALB', 'ROH'}.
    Uses MARA as source of truth.
    """
    if df is None or df.empty:
        return df

    VALID_MT = {'FERT', 'HALB', 'ROH'}

    mt_map = dict(zip(
        mara_df['Material'].astype(str).str.strip(),
        mara_df['Material Type'].astype(str).str.strip().str.upper()
    ))

    df = df.copy()
    df['_MT'] = df['Material'].astype(str).str.strip().map(mt_map)

    # fallback (defensive)
    if 'Material Type' in df.columns:
        df['_MT'] = df['_MT'].fillna(
            df['Material Type'].astype(str).str.strip().str.upper()
        )

    df = df[df['_MT'].isin(VALID_MT)]

    return df.drop(columns=['_MT'], errors='ignore')


def _get_mps99_material_plant_pairs(marc_df):
    if marc_df is None or marc_df.empty:
        return set()

    df = marc_df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if 'Material' not in df.columns or 'Plant' not in df.columns:
        return set()

    candidates = [
        'Plant Specific Material Status',
        'Material Plant Status',
        'Plant Status',
        'MATERIAL PLANT STATUS',
        'MPS',
    ]
    status_col = next((c for c in candidates if c in df.columns), None)
    if status_col is None:
        return set()

    pairs = set()
    for _, row in df.iterrows():
        material = str(row['Material']).strip()
        plant = str(row['Plant']).strip().upper()
        status = normalize_code(row[status_col], as_int=True)
        if status == '99':
            pairs.add((material, plant))

    return pairs


def _exclude_material_plant_status_99(df, pairs, plant_col='Plant'):
    if df is None or df.empty or not pairs:
        return df
    if 'Material' not in df.columns or plant_col not in df.columns:
        return df

    materials = df['Material'].astype(str).str.strip()
    plants = df[plant_col].astype(str).str.strip().str.upper()
    keys = list(zip(materials, plants))
    mask = [key not in pairs for key in keys]
    return df.loc[mask].copy()


def build_sap_like_df(marc_df, mara_df, zplm_marc_df):
    """
    Merges MARC + MARA + ZPLM MARC into a single flat DataFrame that mirrors
    the column layout of the old SAPUI5 Export (Item Browse Report), so that
    every validar_* function that used to read the Item Browse export can
    read this instead. Used by BOTH the Palmela and Tunisia engines.
    """
    marc = marc_df.copy()
    mara = mara_df.copy()

    marc.columns = [str(c).strip() for c in marc.columns]
    mara.columns = [str(c).strip() for c in mara.columns]

    # ------------------------------------------------------------
    # ✅ DEDUPLICATE BEFORE ANY MERGE
    #
    # A left.merge(right, how='left') multiplies each left row once per
    # matching right row. If the source MARC/MARA/ZPLM MARC exports
    # contain duplicate rows for the same key (a common raw-export
    # artifact), every downstream validation ends up seeing — and
    # reporting — the exact same (Material, Plant, Field, Error) more
    # than once. Deduplicating each table on its natural key right here,
    # before any merge, prevents this at the source rather than trying
    # to catch it afterwards.
    # ------------------------------------------------------------
    if 'Material' in marc.columns and 'Plant' in marc.columns:
        marc = marc.drop_duplicates(subset=['Material', 'Plant'], keep='first')
    if 'Material' in mara.columns:
        mara = mara.drop_duplicates(subset=['Material'], keep='first')

    # ------------------------------------------------------------
    # ✅ NORMALIZE MARA COLUMN NAMES
    # ------------------------------------------------------------
    mara_rename_map = {
        'material description': 'Description',
        'gen. item cat. grp': 'General Item Category Group',
        'unit of weight': 'Unit of Weight (ISO Format)',
        'unit of weight (iso format)': 'Unit of Weight (ISO Format)',
        'material group': 'Material Group',
        'product hierarchy': 'Product Hierarchy',
        'old material number': 'Old material number',
        'base unit of measure': 'Base Unit of Measure',
        'export/import group': 'Export/Import Group',
        'intrastat group': 'Export/Import Group',
    }
    mara = _normalize_columns(mara, mara_rename_map)

    # ------------------------------------------------------------
    # ✅ NORMALIZE MARC COLUMN NAMES
    # ------------------------------------------------------------
    marc_rename_map = {
        'availability check': 'Availability Check',
        'consumption mode': 'Consumption Mode',
        'jit delivery schedules': 'JIT Delivery Sched.',
        'source list': 'Source list',
        'source list required': 'Source list required',
        'automatic po': 'Automatic Purchase Order Allowed',
        'automatic purchase order': 'Automatic Purchase Order Allowed',
        'automatic purchase order allowed': 'Automatic Purchase Order Allowed',
        'planning time fence': 'Planning Time Fence',
        'plant-specific material status': 'Plant Specific Material Status',
        'repetitive manufacturing': 'Indicator: Repetitive Manufa. Allow',
        'repetitive manufacturing profile': 'Repetitive Manufacturing Profile',
        'issue storage location': 'Issue Storage Location',
        'production storage location': 'Issue Storage Location',
        'planning strategy group': 'Planning Strategy Group',
        'planning calendar': 'Planning Calendar',
        'mrp type': 'MRP Type',
        'profit center': 'Profit Center',
        'purchasing group': 'Purchasing Group',
        'special procurement type': 'Special Procurement Type',
        'storage location for ext. procurement': 'Storage Location for Ext. Procurement',
        'individual/ collective requirements': 'Individual/ Collective Requirements',
        'individual collective requirements': 'Individual/ Collective Requirements',
        'production supervisor': 'Production Supervisor',
        'production scheduling profile': 'Production Scheduling Profile',
        'advanced planning': 'Advanced Planning',
        'heuristic': 'Heuristic',
        'pp plng procedure': 'PP Plng Procedure',
        'requirement strategy': 'Requirement Strategy',
        'backflush': 'Backflush',
        'material plant type': 'Plant Type',
        'period indicator': 'Period Type',
        'intrastat group': 'Export/Import Group',
    }
    marc = _normalize_columns(marc, marc_rename_map)

    # ------------------------------------------------------------
    # ✅ SELECT MARA COLUMNS
    # ------------------------------------------------------------
    mara_cols = [c for c in [
        'Material', 'Material Type', 'Description', 'Base Unit of Measure',
        'Division', 'Export/Import Group', 'General Item Category Group',
        'Gross Weight', 'Net Weight', 'Unit of Weight (ISO Format)',
        'Material Group', 'Product Hierarchy', 'Packaging Mat. Type',
        'Old material number', 'Batch Management',
    ] if c in mara.columns]

    # ------------------------------------------------------------
    # ✅ MERGE MARC + MARA
    # ------------------------------------------------------------
    df = marc.merge(mara[mara_cols], on='Material', how='left')
    df.rename(columns={'Base Unit of Measure': 'Base UOM'}, inplace=True)

    # ------------------------------------------------------------
    # ✅ MERGE ZPLM
    # ------------------------------------------------------------
    if zplm_marc_df is not None and not zplm_marc_df.empty:
        zplm = zplm_marc_df.copy()
        zplm.columns = [str(c).strip() for c in zplm.columns]

        zplm_rename_map = {
            'Material Plant Type': 'Plant Type',
            'enterproj number': 'Enterproj Number',
            'material transfer': 'Material Transfer',
            'eop date': 'EOP Date',
            'service liability date': 'Service Liability Date',
            'ltb quantity': 'LTB Quantity',
            'ltb start date': 'LTB Start Date',
        }
        zplm = _normalize_columns(zplm, zplm_rename_map)

        zplm_cols = ['Material', 'Plant']
        for col in [
            'Plant Type', 'Enterproj Number', 'Material Transfer',
            'EOP Date', 'Service Liability Date', 'LTB Quantity', 'LTB Start Date'
        ]:
            if col in zplm.columns:
                zplm_cols.append(col)

        zplm = zplm[zplm_cols].drop_duplicates(subset=['Material', 'Plant'], keep='first')
        df = df.merge(zplm, on=['Material', 'Plant'], how='left')
    else:
        for col in ['Plant Type', 'Enterproj Number']:
            if col not in df.columns:
                df[col] = ''

    # ------------------------------------------------------------
    # ✅ NORMALIZE PLANT TYPE
    # ------------------------------------------------------------
    if 'Plant Type' in df.columns:
        df['Plant Type'] = df['Plant Type'].fillna('').astype(str).str.strip().str.upper()

    # ------------------------------------------------------------
    # ✅ BATCH MANAGEMENT (CRITICAL)
    # ------------------------------------------------------------

    # 📌 Global (MARA)
    df['Batch Management'] = df.get('Batch Management', '')

    # 📌 Plant-level (MARC)
    batch_plant_candidates = [
        'Batch Management(Plant)',
        'Batch Management (Plant)',
        'Batch Mgmt Plant'
    ]

    batch_plant_col = next(
        (col for col in df.columns if col in batch_plant_candidates),
        None
    )

    if batch_plant_col:
        df['Batch Management(Plant)'] = df[batch_plant_col]
    else:
        df['Batch Management(Plant)'] = ''

    # 📌 SAPUI5 compatibility (for validar_batch_management)
    #
    # The old Item Browse Report represented this checkbox as '1' (active)
    # / '0' (inactive). Raw MARC uses SAP's native convention: 'X' (active)
    # / blank (inactive). validar_batch_management() still expects the old
    # '1'/'0' convention, so we translate it here rather than touching that
    # (shared, already-proven) validation function.
    def _x_to_1(series):
        s = series.astype(str).str.strip().str.upper()
        return s.map(lambda v: '1' if v == 'X' else ('0' if v in ('', 'NAN', 'NONE') else v))

    df['Batch Management Required Ind.'] = _x_to_1(df['Batch Management(Plant)'])

    # Same translation for JIT Delivery Sched. — validar_jit_delivery_sched()
    # expects '1.0' (active) / '' (inactive), matching how the old Item
    # Browse Report exported this checkbox as a numeric column.
    if 'JIT Delivery Sched.' in df.columns:
        _jit = df['JIT Delivery Sched.'].astype(str).str.strip().str.upper()
        df['JIT Delivery Sched.'] = _jit.map(
            lambda v: '1.0' if v == 'X' else ('' if v in ('', 'NAN', 'NONE') else v)
        )

    # ------------------------------------------------------------
    # ✅ ENSURE CRITICAL COLUMNS EXIST
    # ------------------------------------------------------------
    required_cols = [
        'Plant Type',
        'Batch Management Required Ind.',
        'MRP Type',
        'Availability Check'
    ]

    for col in required_cols:
        if col not in df.columns:
            df[col] = ''

    return df


def build_material_transfer_map(zplm_marc_df):
    """
    Creates a map (Material, Plant) -> bool indicating whether 'Material Transfer'
    is marked in ZPLM MARC. Considers values in {'X','1','TRUE','YES','Y'} (case-insensitive).
    """
    if zplm_marc_df is None or zplm_marc_df.empty:
        return {}

    df = zplm_marc_df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]

    candidates = ['MATERIAL TRANSFER', 'MAT_TRANSFER', 'Z_MATERIAL_TRANSFER']
    mt_col = next((c for c in candidates if c in df.columns), None)

    if mt_col is None or 'MATERIAL' not in df.columns or 'PLANT' not in df.columns:
        return {}

    def is_marked(x):
        s = '' if pd.isna(x) else str(x).strip().upper()
        return s in {'X', '1', 'TRUE', 'YES', 'Y'}

    df['_KEY'] = list(zip(
        df['MATERIAL'].astype(str).str.strip(),
        df['PLANT'].astype(str).str.strip()
    ))
    df['_MARKED'] = df[mt_col].apply(is_marked)

    return {key: grp['_MARKED'].any() for key, grp in df.groupby('_KEY')}


def build_csr_exempt_set(csr_item_df, csr_head_df) -> set:
    """
    Builds the set of (Material, Plant) pairs exempt from the FP -> MRP = M1
    rule because they have an active CSR (status I1, I2 or I4).

    Logic:
      1. Filter ZPLM_A_CSR_ITEM where CSR Status in {I1, I2, I4}  (I3/U3/H2 excluded).
      2. Take UUID from those rows.
      3. Join UUID -> ZPLM_A_CSR_HEAD to get Plant.
      4. Return {(Material, Plant)}.

    Column name detection covers SAP export variants:
      - Status column: "Change Status Request Status", "CSR Status", "Item Status", etc.
      - UUID column  : "UUID", "GUID", "Id", etc.
      - The function logs a warning (instead of printing) when columns are not found.
    """
    import logging
    _log = logging.getLogger(__name__)

    if csr_item_df is None or csr_item_df.empty:
        return set()
    if csr_head_df is None or csr_head_df.empty:
        return set()

    def _pick(df, candidates):
        return next((c for c in candidates if c in df.columns), None)

    i_mat = _pick(csr_item_df, [
        "Material", "MATNR", "material", "MATERIAL",
    ])
    i_uuid = _pick(csr_item_df, [
        "UUID", "GUID", "ZUUID", "uuid", "Id", "ID",
    ])
    i_status = _pick(csr_item_df, [
        "Change Status Request Status",  # nome real do export SAP
        "CSR Status", "CSR_STATUS", "Item Status", "Item_Status",
        "ZSTATUS", "Status", "STATUS",
    ])
    h_uuid = _pick(csr_head_df, [
        "UUID", "GUID", "ZUUID", "uuid", "Id", "ID",
    ])
    h_plant = _pick(csr_head_df, [
        "Plant", "WERKS", "plant", "PLANT",
    ])

    missing = [name for name, val in [
        ("i_mat",    i_mat),
        ("i_uuid",   i_uuid),
        ("i_status", i_status),
        ("h_uuid",   h_uuid),
        ("h_plant",  h_plant),
    ] if val is None]

    if missing:
        _log.warning(
            "build_csr_exempt_set: could not find columns %s. "
            "ITEM cols=%s | HEAD cols=%s",
            missing,
            list(csr_item_df.columns),
            list(csr_head_df.columns),
        )
        return set()

    item = csr_item_df.copy()
    item[i_mat]    = item[i_mat].astype(str).str.strip()
    item[i_uuid]   = item[i_uuid].astype(str).str.strip()
    item[i_status] = item[i_status].astype(str).str.strip().str.upper()

    # I1, I2, I4 = active CSR statuses that grant exemption
    item_filtered = item[item[i_status].isin({"I1", "I2", "I4"})]
    if item_filtered.empty:
        _log.debug(
            "build_csr_exempt_set: no rows with active CSR status (I1/I2/I4) found. "
            "Distinct values in '%s': %s",
            i_status,
            item[i_status].unique().tolist(),
        )
        return set()

    head = csr_head_df.copy()
    head[h_uuid]  = head[h_uuid].astype(str).str.strip()
    head[h_plant] = head[h_plant].astype(str).str.strip().str.upper()

    merged = item_filtered[[i_mat, i_uuid]].merge(
        head[[h_uuid, h_plant]],
        left_on=i_uuid,
        right_on=h_uuid,
        how="inner",
    )

    if merged.empty:
        _log.warning(
            "build_csr_exempt_set: join between ITEM and HEAD returned 0 rows. "
            "Check that UUID values match between the two files."
        )
        return set()

    exempt = set(zip(merged[i_mat].str.strip(), merged[h_plant]))
    _log.debug("build_csr_exempt_set: %d exempt (Material, Plant) pairs built.", len(exempt))
    return exempt


def get_effective_material_type(material, real_mt, rule_name):
    """
    Returns the effective Material Type for a given rule,
    allowing rule-specific overrides.
    Priority:
    1) Rule-specific override
    2) Material default override (__default__)
    3) Real Material Type from data
    """
    material = (material or "").strip()
    real_mt = (real_mt or "").strip()

    rules = RULE_MATERIAL_TYPE_OVERRIDES.get(material)
    if not rules:
        return real_mt

    # 1️⃣ Rule-specific override
    if rule_name in rules:
        return rules[rule_name]

    # 2️⃣ Default override
    if "__default__" in rules:
        return rules["__default__"]

    # 3️⃣ Fallback to real MT
    return real_mt

# =============================================================================
# 3) LOG PAYLOAD BUILDERS
# =============================================================================
# Note: build_fp_log_payload and build_dc_log_payload are defined in Part 4
# (end of file) where _safe_name is available. Forward references are resolved
# at runtime — Python loads the full module before any call is made.

# =============================================================================
# SAP VALIDATION ENGINE — PART 2: ALL VALIDATION FUNCTIONS
# =============================================================================
# Depends on: Part 1 (imports, loaders, helpers)
#
# SECTIONS:
#   A) SAPUI5 Export — field-level validations
#   B) Cross-file validations (SAPUI5 + MVKE, MARC, MARA, ZPLM)
#   C) Data Cleanup validations (FP55, LTB, EKPO, MARD, stock)
#   D) Cross-existence validations (MARA ↔ ZPLM MARA, MARC ↔ ZPLM MARC)
#   E) Post-processing helpers (Old material number, Enterproj enrichment, Neg. stocks)
# =============================================================================


# =============================================================================
# A) SAPUI5 EXPORT — FIELD-LEVEL VALIDATIONS
# =============================================================================

def validar_campo_obrigatorio(df, coluna, resultados):
    """
    Validates mandatory fields. Special case: 'Description' is not required
    when Plant Specific Material Status is 70 or 75.
    """
    if coluna == "Description" and "Plant Specific Material Status" in df.columns:
        df['_PSMS'] = df['Plant Specific Material Status'].apply(
            lambda x: int(float(str(x))) if str(x).strip() not in ['', 'nan'] else None
        )

    if coluna not in df.columns:
        print(f"⚠️ Column '{coluna}' was not found in the 'SAPUI5 Export' sheet.")
        return

    for idx in df[df[coluna].isnull()].index:
        if coluna == "Description" and "_PSMS" in df.columns:
            if df.at[idx, '_PSMS'] in (70, 75):
                continue
        resultados.append({
            'Line':     idx + 2,
            'Field':    coluna,
            'Error':    'Mandatory field blank',
            'Material': str(df.at[idx, 'Material']).strip() if 'Material' in df.columns else '',
            'Plant':    str(df.at[idx, 'Plant']).strip()    if 'Plant'    in df.columns else '',
        })


def validar_procurement_type(df, resultados):
    """
    Validates 'Procurement Type'. Plant Type takes precedence over Material Type.
    - Universal domain: E, F, X (mandatory).
    - Plant Type PP  -> F or X
    - Plant Type MP/FP -> E or X
    - Material Type FERT/HALB -> E or X  (when no Plant Type override)
    - Material Type ROH -> F or X         (when no Plant Type override)
    """
    needed = {'Procurement Type', 'Material Type', 'Plant Type'}
    if not needed.issubset(df.columns):
        return

    universal = {'E', 'F', 'X'}

    for idx, row in df.iterrows():
        erros = []
        material = str(row.get('Material', '')).strip() if 'Material' in df.columns else ''
        plant    = str(row.get('Plant', '')).strip()    if 'Plant'    in df.columns else ''

        raw_pt = row['Procurement Type']
        pt     = '' if pd.isna(raw_pt) else str(raw_pt).strip().upper()
        mt     = '' if pd.isna(row['Material Type']) else str(row['Material Type']).strip().upper()
        ptype  = '' if pd.isna(row['Plant Type'])    else str(row['Plant Type']).strip().upper()

        # 1) Mandatory & domain
        if pt == '':
            erros.append('Mandatory field blank')
        elif pt not in universal:
            erros.append(f"Invalid value: '{raw_pt}' (allowed: E, F, X)")

        # 2) Plant Type precedence
        apply_mt = True
        if pt in universal and ptype != '':
            if ptype == 'PP':
                if pt not in {'F', 'X'}:
                    erros.append(f"For Plant Type 'PP', Procurement Type should be 'F' or 'X' (current: '{pt}')")
                apply_mt = False
            elif ptype in {'MP', 'FP'}:
                if pt not in {'E', 'X'}:
                    erros.append(f"For Plant Type '{ptype}', Procurement Type should be 'E' or 'X' (current: '{pt}')")
                apply_mt = False

        # 3) Material Type rules (only when no Plant Type override)
        if pt in universal and apply_mt:
            if mt in {'FERT', 'HALB'} and pt not in {'E', 'X'}:
                erros.append(f"For Material Type '{mt}', Procurement Type should be 'E' or 'X' (current: '{pt}')")
            elif mt == 'ROH' and pt not in {'F', 'X'}:
                erros.append(f"For Material Type 'ROH', Procurement Type should be 'F' or 'X' (current: '{pt}')")

        if erros:
            resultados.append({
                'Line':     idx + 2,
                'Field':    'Procurement Type',
                'Error':    ' \n'.join(erros),
                'Material': material,
                'Plant':    plant,
            })


def validar_automatic_purchase_order(df, resultados):
    """PP -> 'X' (Yes); others -> blank (''), SAP checkbox logic."""

    if 'Plant Type' not in df.columns or 'Automatic Purchase Order Allowed' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'Automatic Purchase Order Allowed' not found.")
        return

    for idx, row in df.iterrows():
        erros = []

        plant_type = str(row.get('Plant Type', '')).strip().upper()
        raw_val = row.get('Automatic Purchase Order Allowed')

        val = str(raw_val).strip().upper() if pd.notna(raw_val) else ''

        # ✅ Normalizar checkbox
        if val in {'X', '1', 'YES'}:
            val_norm = 'X'
        else:
            val_norm = ''

        # ✅ Regra
        if val_norm == '':
            # só é obrigatório para PP
            if plant_type == 'PP':
                erros.append('Mandatory field blank')

        else:
            if plant_type == 'PP':
                if val_norm != 'X':
                    erros.append(
                        f"For Plant Type 'PP', should be 'X' (current: '{val}')"
                    )
            else:
                if val_norm != '':
                    erros.append(
                        f"For Plant Type '{plant_type}', should be blank (current: '{val}')"
                    )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Automatic Purchase Order Allowed',
                'Error': ' | '.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant': str(row.get('Plant', '')).strip(),
            })

def validar_batch_management(df, resultados):
    """
    Rules:
      - TN02: Batch Management Required Ind. must be empty.
      - Outside TN02:
          - PP / COI -> 'Yes' (1)
          - All other Plant Types -> 'No' (0)

    The TN02 rule takes precedence over the Plant Type rule.
    """
    col = 'Batch Management Required Ind.'

    required_columns = {
        'Plant',
        'Plant Type',
        col,
    }

    if not required_columns.issubset(df.columns):
        missing_columns = sorted(
            required_columns - set(df.columns)
        )

        print(
            "⚠️ Missing columns for Batch Management validation: "
            f"{missing_columns}"
        )
        return

    for idx, row in df.iterrows():
        erros = []

        plant = str(
            row.get('Plant', '')
        ).strip().upper()

        plant_type = str(
            row.get('Plant Type', '')
        ).strip().upper()

        raw_batch_ind = row.get(col)

        batch_ind = (
            ''
            if pd.isna(raw_batch_ind)
            else str(raw_batch_ind).strip().upper()
        )

        # ========================================================
        # RULE 1
        # TN02 must always be empty.
        #
        # This rule takes precedence over the Plant Type rule.
        # ========================================================
        if plant == 'TN02':
            if batch_ind != '':
                erros.append(
                    "For Plant 'TN02', should be blank "
                    f"(Current: '{batch_ind}')"
                )

        # ========================================================
        # RULE 2
        # Outside TN02, PP and COI must have Yes (1).
        #
        # X is accepted as equivalent to 1.
        # ========================================================
        elif plant_type in ('PP', 'COI'):
            if batch_ind == '':
                erros.append(
                    'Mandatory field blank'
                )

            elif batch_ind not in ('1', 'X'):
                erros.append(
                    f"For Plant Type '{plant_type}', "
                    "should be 'Yes' "
                    f"(Current: '{batch_ind}')"
                )

        # ========================================================
        # RULE 3
        # Outside TN02, all other Plant Types must have No (0).
        # ========================================================
        else:
            if batch_ind == '':
                erros.append(
                    'Mandatory field blank'
                )

            elif batch_ind != '0':
                erros.append(
                    f"For Plant Type '{plant_type}', "
                    "should be 'No' "
                    f"(Current: '{batch_ind}')"
                )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' | '.join(erros),
                'Material': str(
                    row.get('Material', '')
                ).strip(),
                'Plant': plant,
            })

def validar_consumption_mode(df, resultados):
    """FERT -> 0 or empty; all other Material Types -> 2 (mandatory)."""
    if 'Consumption Mode' not in df.columns:
        print("⚠️ Column 'Consumption Mode' not found.")
        return

    has_mt = 'Material Type' in df.columns

    for idx, row in df.iterrows():
        erros = []
        raw  = row['Consumption Mode']
        val  = normalize_code(raw, as_int=True)
        mt   = str(row['Material Type']).strip().upper() if has_mt else ''
        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()

        if mt == 'FERT':
            if val not in ('', '0'):
                erros.append(f"Invalid value: '{raw}' (Material Type FERT requires Consumption Mode = '0' or empty)")
        else:
            if val == '':
                erros.append('Mandatory field blank')
            elif val != '2':
                erros.append(f"Invalid value: '{raw}' (should be '2' for non-FERT materials)")

        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'Consumption Mode',
                'Error': '\n'.join(erros), 'Material': material, 'Plant': plant,
            })


def validar_division(df, resultados):
    """Division must always be '00'."""
    if 'Division' not in df.columns:
        print("⚠️ Column 'Division' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        raw = row['Division']
        val = normalize_code(raw, as_int=True, zero_pad=2)
        if val == '':
            erros.append('Mandatory field blank')
        elif val != '00':
            erros.append(f"Invalid value: '{raw}' (should be '00')")
        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'Division',
                'Error': '\n'.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_export_import_group(df, resultados):
    """Export/Import Group must always be 'ZVIS'."""
    if 'Export/Import Group' not in df.columns:
        print("⚠️ Column 'Export/Import Group' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        val = str(row['Export/Import Group']).strip()
        if pd.isna(row['Export/Import Group']) or val == '':
            erros.append('Mandatory field blank')
        elif val != 'ZVIS':
            erros.append(f"Invalid value: '{val}' (should be 'ZVIS')")
        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'Export/Import Group',
                'Error': '\n'.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_general_item_category_group(df, resultados):
    """General Item Category Group must always be 'NORM'."""
    if 'General Item Category Group' not in df.columns:
        print("⚠️ Column 'General Item Category Group' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        val = str(row['General Item Category Group']).strip()
        if pd.isna(row['General Item Category Group']) or val == '':
            erros.append('Mandatory field blank')
        elif val != 'NORM':
            erros.append(f"Invalid value: '{val}' (should be 'NORM')")
        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'General Item Category Group',
                'Error': '\n'.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_gross_weight(df, resultados):
    """FP only: Gross Weight mandatory > 0, equal to Net Weight, within unit limits."""
    required = {'Gross Weight', 'Net Weight', 'Unit of Weight (ISO Format)', 'Plant Type'}
    if not required.issubset(df.columns):
        print("⚠️ Required columns for Gross Weight validation not found.")
        return

    for idx, row in df.iterrows():
        plant_type = '' if pd.isna(row['Plant Type']) else str(row['Plant Type']).strip()
        if plant_type != 'FP':
            continue

        erros = []
        gross = row['Gross Weight']
        net   = row['Net Weight']
        unit  = str(row['Unit of Weight (ISO Format)']).strip().upper() if not pd.isna(row['Unit of Weight (ISO Format)']) else ''

        if pd.isna(gross) or pd.isna(net):
            erros.append('Mandatory fields blank')
        else:
            try:
                gv, nv = float(gross), float(net)
                if gv <= 0 or nv <= 0:
                    erros.append(f"Values must be greater than 0: Gross={gv}, Net={nv}")
                if gv != nv:
                    erros.append(f"Gross Weight ({gv}) not equal compared to Net Weight ({nv})")
                limite = {'KG': 5, 'G': 5000, 'MG': 5_000_000}.get(unit, 10)
                if gv > limite or nv > limite:
                    erros.append(f"Value too high for unit {unit}: Gross={gv}, Net={nv} (> {limite} {unit})")
            except ValueError:
                erros.append(f"Invalid value: Gross='{gross}', Net='{net}' (not numeric)")

        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'Gross Weight',
                'Error': ' \n '.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_net_weight(df, resultados):
    """FP only: Net Weight mandatory > 0, equal to Gross Weight, within unit limits."""
    required = {'Gross Weight', 'Net Weight', 'Unit of Weight (ISO Format)', 'Plant Type'}
    if not required.issubset(df.columns):
        print("⚠️ Required columns for Net Weight validation not found.")
        return

    for idx, row in df.iterrows():
        plant_type = '' if pd.isna(row['Plant Type']) else str(row['Plant Type']).strip()
        if plant_type != 'FP':
            continue

        erros = []
        gross = row['Gross Weight']
        net   = row['Net Weight']
        unit  = str(row['Unit of Weight (ISO Format)']).strip().upper() if not pd.isna(row['Unit of Weight (ISO Format)']) else ''

        if pd.isna(gross) or pd.isna(net):
            erros.append('Mandatory fields blank')
        else:
            try:
                gv, nv = float(gross), float(net)
                if nv <= 0 or gv <= 0:
                    erros.append(f"Values must be greater than 0: Gross={gv}, Net={nv}")
                if gv != nv:
                    erros.append(f"Gross Weight ({gv}) not equal compared to Net Weight ({nv})")
                limite = {'KG': 5, 'G': 5000, 'MG': 5_000_000}.get(unit, 10)
                if gv > limite or nv > limite:
                    erros.append(f"Value too high for unit {unit}: Gross={gv}, Net={nv} (> {limite} {unit})")
            except ValueError:
                erros.append(f"Invalid value: Gross='{gross}', Net='{net}' (not numeric)")

        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'Net Weight',
                'Error': ' \n '.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_jit_delivery_sched(df, resultados):
    """PP -> '1' (Yes); other Plant Types -> must be empty."""
    if 'Plant Type' not in df.columns or 'JIT Delivery Sched.' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'JIT Delivery Sched.' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        plant_type = str(row['Plant Type']).strip()
        jit_sched  = str(row['JIT Delivery Sched.']).strip()

        if plant_type == 'PP':
            if pd.isna(row['JIT Delivery Sched.']) or jit_sched == '':
                erros.append("Mandatory field blank For Plant Type 'PP'")
            elif jit_sched != '1.0':
                erros.append(f"For Plant Type 'PP', should be 'Yes' (Current: '{jit_sched}')")
        else:
            if jit_sched != '' and not pd.isna(row['JIT Delivery Sched.']):
                erros.append(f"For Plant Type '{plant_type}', field must be empty (Current: '{jit_sched}')")

        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'JIT Delivery Sched.',
                'Error': ' \n '.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_loading_group(df, resultados):
    """Loading Group must always be '0001'."""
    if 'Loading Group' not in df.columns:
        return

    for idx, row in df.iterrows():
        erros = []
        raw = row['Loading Group']
        val = normalize_code(raw, as_int=True, zero_pad=4)
        if val == '':
            erros.append('Mandatory field blank')
        elif val != '0001':
            erros.append(f"Invalid value: '{raw}' (should be '0001')")
        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'Loading Group',
                'Error': '\n'.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_lot_sizing_procedure(df, resultados):
    """
    Lot Sizing Procedure validation.

    Rules and precedence:

    1. Any material in TN02:
         Lot Sizing Procedure must be 'EX'.

    2. Outside TN02:
         - FP or MP: must be 'EX'.
         - PP: must be 'ZK'.

    3. Exception outside TN02:
         - MP material containing the token '14K450':
           allows either 'EX' or 'ZK'.

    Notes:
      - The TN02 rule has the highest precedence.
      - Therefore, a TN02 PP material must have 'EX', not 'ZK'.
      - The field is mandatory for all validated rows.
    """

    if df is None or df.empty:
        return

    required_columns = {
        'Material',
        'Plant',
        'Plant Type',
        'Lot Sizing Procedure',
    }

    if not required_columns.issubset(df.columns):
        missing_columns = sorted(
            required_columns - set(df.columns)
        )

        print(
            "Missing columns for Lot Sizing Procedure validation: "
            f"{missing_columns}"
        )
        return

    token_pattern = re.compile(
        r'(?i)(?:^|[^A-Z0-9])14K450(?:[^A-Z0-9]|$)'
    )

    def has_14k450(material):
        return bool(
            token_pattern.search(
                str(material).strip()
            )
        )

    for idx, row in df.iterrows():
        erros = []

        material = _norm_str(
            row.get('Material')
        )

        plant = _norm_str(
            row.get('Plant')
        ).upper()

        plant_type = _norm_str(
            row.get('Plant Type')
        ).upper()

        lot_sizing = _norm_str(
            row.get('Lot Sizing Procedure')
        ).upper()

        # ========================================================
        # Mandatory field
        # ========================================================
        if lot_sizing == '':
            if plant == 'TN02':
                erros.append(
                    "Mandatory field blank "
                    "(all TN02 materials require "
                    "Lot Sizing Procedure = 'EX')"
                )
            elif plant_type in {'FP', 'MP'}:
                erros.append(
                    f"Mandatory field blank "
                    f"(Plant Type '{plant_type}' requires "
                    "Lot Sizing Procedure = 'EX')"
                )
            elif plant_type == 'PP':
                erros.append(
                    "Mandatory field blank "
                    "(Plant Type 'PP' requires "
                    "Lot Sizing Procedure = 'ZK')"
                )
            else:
                erros.append(
                    "Mandatory field blank"
                )

        else:
            # ====================================================
            # RULE 1
            # All TN02 materials must have EX.
            #
            # This rule takes precedence over:
            #   - Plant Type rules
            #   - MP + 14K450 exception
            # ====================================================
            if plant == 'TN02':
                if lot_sizing != 'EX':
                    erros.append(
                        "All materials in TN02 require "
                        "Lot Sizing Procedure = 'EX' "
                        f"(current: '{lot_sizing}')"
                    )

            # ====================================================
            # RULE 2
            # MP + 14K450 exception outside TN02
            # ====================================================
            elif (
                plant_type == 'MP'
                and has_14k450(material)
            ):
                if lot_sizing not in {'EX', 'ZK'}:
                    erros.append(
                        f"Invalid value: '{lot_sizing}' "
                        "for Plant Type 'MP' with material "
                        "containing '14K450'. "
                        "Allowed values: 'EX' or 'ZK'"
                    )

            # ====================================================
            # RULE 3
            # FP and MP outside TN02 must have EX.
            # ====================================================
            elif plant_type in {'FP', 'MP'}:
                if lot_sizing != 'EX':
                    erros.append(
                        f"Plant Type '{plant_type}' requires "
                        "Lot Sizing Procedure = 'EX' "
                        f"(current: '{lot_sizing}')"
                    )

            # ====================================================
            # RULE 4
            # PP outside TN02 must have ZK.
            # ====================================================
            elif plant_type == 'PP':
                if lot_sizing != 'ZK':
                    erros.append(
                        "Plant Type 'PP' requires "
                        "Lot Sizing Procedure = 'ZK' "
                        f"(current: '{lot_sizing}')"
                    )

            # ====================================================
            # Unknown or unsupported Plant Type
            # ====================================================
            elif lot_sizing not in {'EX', 'ZK'}:
                erros.append(
                    f"Invalid Lot Sizing Procedure: "
                    f"'{lot_sizing}'"
                )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Lot Sizing Procedure',
                'Error': ' \n'.join(erros),
                'Material': material,
                'Plant': plant,
            })


def validar_per_lot_sz_plng_cal(
    marc_df,
    resultados,
    mara_df=None,
    matplsta_df=None
):
    """
    Validate Per.Lot Sz.Plng Cal. by Plant Type, Planning Status,
    Plant-Specific Material Status and Advanced Planning.

    FP / MP:
    - The field is mandatory only when:
        - Advanced Planning is active;
        - Planning Status is not 9;
        - PSMS is not 25 or 75.

    PP:
    - Planning Status 1:
        - The field must be empty.
    - Planning Status 9 or blank:
        - The field is not validated.
    - Any other Planning Status:
        - The field is not validated until a specific rule is defined.

    mara_df is kept only for backward compatibility.
    matplsta_df is used to obtain Planning Status by Material and Plant.
    """

    field = 'Per.Lot Sz.Plng Cal.'

    if marc_df is None or marc_df.empty or field not in marc_df.columns:
        return

    def _pick(candidates):
        return next(
            (column for column in candidates if column in marc_df.columns),
            None
        )

    def _s(value):
        return '' if pd.isna(value) else str(value).strip()

    def is_true(value):
        return _s(value).upper() in {
            'TRUE',
            'X',
            '1',
            'YES',
            'Y',
        }

    # -------------------------------------------------------------------------
    # Identify the required MARC columns
    # -------------------------------------------------------------------------
    pt_col = _pick([
        'Plant Type',
        'Material Plant Type',
    ])

    advanced_col = _pick([
        'Advanced Planning',
    ])

    psms_col = _pick([
        'Plant-Specific Material Status',
        'Plant Specific Material Status',
    ])

    if pt_col is None or advanced_col is None or psms_col is None:
        return

    # -------------------------------------------------------------------------
    # Build Material + Plant -> Planning Status mapping
    # -------------------------------------------------------------------------
    planning_status_map = {}

    if matplsta_df is not None and not matplsta_df.empty:
        mps = matplsta_df.copy()
        mps.columns = mps.columns.astype(str).str.strip()

        required_matplsta_columns = {
            'Product Number',
            'Location',
            'Planning Status',
        }

        if required_matplsta_columns.issubset(mps.columns):
            for _, source_row in mps.iterrows():
                material_key = _s(
                    source_row.get('Product Number')
                )

                plant_key = _s(
                    source_row.get('Location')
                ).upper()

                planning_status_map[(material_key, plant_key)] = (
                    normalize_code(
                        source_row.get('Planning Status'),
                        as_int=True
                    )
                )

    # -------------------------------------------------------------------------
    # Validate each MARC row
    # -------------------------------------------------------------------------
    for idx, row in marc_df.iterrows():
        material = _s(
            row.get('Material')
        )

        plant = _s(
            row.get('Plant')
        ).upper()

        plant_type = _s(
            row.get(pt_col)
        ).upper()

        raw_value = row.get(field)
        value = _s(raw_value)

        advanced = row.get(advanced_col)

        psms = normalize_code(
            row.get(psms_col),
            as_int=True
        )

        planning_status = planning_status_map.get(
            (material, plant),
            ''
        )

        # =====================================================================
        # PLANT TYPE PP
        # =====================================================================
        if plant_type == 'PP':

            # PP with Planning Status 9 or blank:
            # Per.Lot Sz.Plng Cal. is outside the validation scope.
            if planning_status in {'', '9'}:
                continue

            # PP with Planning Status 1:
            # the field must be empty.
            if planning_status == '1' and value != '':
                resultados.append({
                    'Line': idx + 2,
                    'Field': field,
                    'Error': (
                        "Field must be empty for Plant Type 'PP' "
                        "with Planning Status 1 "
                        f"(current: '{value}')"
                    ),
                    'Material': material,
                    'Plant': plant,
                })

            # PP does not enter the FP/MP validation rules.
            continue

        # =====================================================================
        # PLANT TYPES FP AND MP
        # =====================================================================
        if plant_type not in {'FP', 'MP'}:
            continue

        # ---------------------------------------------------------------------
        # Exception 1:
        # Planning Status 9 does not require Per.Lot Sz.Plng Cal.
        # ---------------------------------------------------------------------
        if planning_status == '9':
            continue

        # ---------------------------------------------------------------------
        # Exception 2:
        # PSMS 25 or 75 does not require Per.Lot Sz.Plng Cal.
        # ---------------------------------------------------------------------
        if psms in {'25', '75'}:
            continue

        # ---------------------------------------------------------------------
        # Exception 3:
        # Per.Lot Sz.Plng Cal. is not mandatory when
        # Advanced Planning is inactive.
        # ---------------------------------------------------------------------
        if not is_true(advanced):
            continue

        # ---------------------------------------------------------------------
        # FP/MP:
        # The field is mandatory when none of the exceptions above applies.
        # ---------------------------------------------------------------------
        if value == '':
            resultados.append({
                'Line': idx + 2,
                'Field': field,
                'Error': (
                    f"Mandatory field blank for Plant Type '{plant_type}' "
                    "when Advanced Planning is active, "
                    "Planning Status is not 9, "
                    "and PSMS is not 25 or 75"
                ),
                'Material': material,
                'Plant': plant,
            })

def validar_material_group(df, resultados):
    """Material Group is mandatory. Certain values are enforced per Material Type."""
    if 'Material Type' not in df.columns or 'Material Group' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Material Group' not found.")
        return

    for idx, row in df.iterrows():
        erros = []

        material = str(row.get('Material', '')).strip()
        real_mt  = str(row.get('Material Type', '')).strip()

        # ✅ Mantém o teu setup (com override se quiseres)
        mt = get_effective_material_type(
            material,
            real_mt,
            'validar_material_group',
        )

        raw_mg = row.get('Material Group')
        mg = '' if pd.isna(raw_mg) else str(raw_mg).strip()

        if mg == '':
            erros.append('Mandatory field blank')
        else:
            # ✅ NOVA LÓGICA (valores OBRIGATÓRIOS)
            if mt == 'FERT' and mg != 'L004':
                erros.append(f"For FERT, must be 'L004' (Current: '{mg}')")

            elif mt == 'HALB' and mg != 'L003':
                erros.append(f"For HALB, must be 'L003' (Current: '{mg}')")

            # ✅ ROH fica como estava
            elif mt == 'ROH' and mg in ('L003', 'L004'):
                erros.append(f"For ROH, cannot be '{mg}'")

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Material Group',
                'Error': ' \n'.join(erros),
                'Material': material,
                'Plant': str(row.get('Plant', '')).strip(),
            })

def validar_material_group_appoint(df, resultados):
    """
    Informational count: materials (FERT, HALB, ROH) with Material Group = 'APPOINT'.
    Not an error per se — surfaced in Field With Error and Adjusted Summary
    so the team is aware of how many materials still carry this value.
    """
    if 'Material Group' not in df.columns or 'Material Type' not in df.columns:
        return

    valid_types = {'FERT', 'HALB', 'ROH'}

    for idx, row in df.iterrows():
        mt  = str(row.get('Material Type', '')).strip().upper()
        mg  = str(row.get('Material Group', '')).strip().upper()
        if mt not in valid_types:
            continue
        if mg == 'APPOINT':
            resultados.append({
                'Line':     idx + 2,
                'Field':    'Material Group (APPOINT)',
                'Error':    f"Material Group = 'APPOINT' (Material Type: {mt})",
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_purchasing_group_dxx(df, resultados):
    """
    Informational count: materials (FERT, HALB, ROH) with Purchasing Group = 'DXX'.
    Not an error per se — surfaced in Field With Error and Adjusted Summary
    so the team is aware of how many materials still carry this value.
    """
    if 'Purchasing Group' not in df.columns or 'Material Type' not in df.columns:
        return

    valid_types = {'FERT', 'HALB', 'ROH'}

    for idx, row in df.iterrows():
        mt  = str(row.get('Material Type', '')).strip().upper()
        pg  = str(row.get('Purchasing Group', '')).strip().upper()
        if mt not in valid_types:
            continue
        if pg == 'DXX':
            resultados.append({
                'Line':     idx + 2,
                'Field':    'Purchasing Group (DXX)',
                'Error':    f"Purchasing Group = 'DXX' (Material Type: {mt})",
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_period_type(df, resultados):
    """Period Type must be one of: T, P, M, W."""
    if 'Period Type' not in df.columns:
        print("⚠️ Column 'Period Type' not found.")
        return

    valid = {'T', 'P', 'M', 'W'}

    for idx, row in df.iterrows():
        erros = []
        val = str(row['Period Type']).strip()
        if pd.isna(row['Period Type']) or val == '':
            erros.append('Mandatory field blank')
        elif val not in valid:
            erros.append(f"Invalid value: '{val}' (allowed: T, P, M, W)")
        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'Period Type',
                'Error': ' \n'.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_planning_calendar(df, resultados):
    """
    Planning Calendar validation (PSMS ∈ [10, 70] only):
    - PT01 + FP/MP -> 'PT1'
    - PT02 + FP/MP -> 'PT2'
    - PP -> not empty, not 'PT2' (PT01) or not empty (PT02)
    - Exception: MP + '14K450' in material name -> cannot be empty or 'PT2'.
    """
    required = {'Material', 'Plant', 'Plant Type', 'Planning Calendar', 'Plant Specific Material Status'}
    if not required.issubset(df.columns):
        missing = [c for c in required if c not in df.columns]
        print(f"⚠️ Missing columns for Planning Calendar validation: {missing}")
        return

    def parse_status(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return int(float(str(val).strip()))
        except Exception:
            return None

    for idx, row in df.iterrows():
        erros = []
        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()
        mt       = str(row.get('Plant Type', '')).strip().upper()
        pc_raw   = row.get('Planning Calendar', None)
        pc       = '' if pd.isna(pc_raw) else str(pc_raw).strip().upper()
        psms     = parse_status(row.get('Plant Specific Material Status', None))
        has14k   = '14K450' in material.upper()

        if psms is None or not (10 <= psms <= 70):
            continue

        # Exception: MP + 14K450 override
        if has14k and mt == 'MP':
            if pc == '':
                erros.append("For Plant Type 'MP' with '14K450', Planning Calendar must not be empty")
            elif pc == 'PT2':
                erros.append("For Plant Type 'MP' with '14K450', Planning Calendar must not be 'PT2'")
            if erros:
                resultados.append({'Line': idx + 2, 'Field': 'Planning Calendar',
                                   'Error': ' \n'.join(erros), 'Material': material, 'Plant': plant})
            continue

        # Standard rules
        if plant == 'PT01':
            if mt in ('FP', 'MP') and pc != 'PT1':
                erros.append(f"For Plant PT01 and Plant Type '{mt}', Planning Calendar must be 'PT1' (Current: '{pc_raw}')")
            elif mt == 'PP':
                if pc == '':
                    erros.append("For Plant PT01 and Plant Type 'PP', Planning Calendar must not be empty")
                elif pc == 'PT2':
                    erros.append("For Plant PT01 and Plant Type 'PP', Planning Calendar must not be 'PT2'")
        elif plant == 'PT02':
            if mt in ('FP', 'MP') and pc != 'PT2':
                erros.append(f"For Plant PT02 and Plant Type '{mt}', Planning Calendar must be 'PT2' (Current: '{pc_raw}')")
            elif mt == 'PP' and pc == '':
                erros.append("For Plant PT02 and Plant Type 'PP', Planning Calendar must not be empty")

        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Planning Calendar',
                               'Error': ' \n'.join(erros), 'Material': material, 'Plant': plant})


def validar_planning_time_fence(df, resultados):
    """
    Planning Time Fence rules (PSMS = 10 is skipped):
    - PD -> empty or 0
    - P1 / M1 -> must be > 0
    """
    required = {'Material', 'Plant', 'MRP Type', 'Planning Time Fence'}
    if not required.issubset(df.columns):
        print("⚠️ Required columns for Planning Time Fence not found.")
        return

    def parse_tf(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return float(str(val).strip())
        except Exception:
            return 'INVALID'

    def parse_status(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return int(float(str(val).strip()))
        except Exception:
            return None

    for idx, row in df.iterrows():
        erros = []
        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()
        mt       = str(row.get('MRP Type', '')).strip().upper()
        tf       = parse_tf(row.get('Planning Time Fence', None))
        psms     = parse_status(row.get('Plant Specific Material Status', None))

        if psms == 10:
            continue

        if tf == 'INVALID':
            erros.append("Planning Time Fence contains non-numeric value.")
        else:
            if mt == 'PD' and tf not in (None, 0):
                erros.append(f"For {mt}, Planning Time Fence should be empty or 0 (Current: '{tf}')")
            elif mt in ('P1', 'M1') and (tf is None or tf == 0):
                erros.append(f"For {mt}, Planning Time Fence should be > 0 (Current: '{tf}')")

        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Planning Time Fence',
                               'Error': ' \n'.join(erros), 'Material': material, 'Plant': plant})


def validar_mrp_type(df, resultados, zplm_marc_df=None,
                     csr_item_df=None, csr_head_df=None):
    """
    MRP Type validation with PSMS precedence, FP/MP transfer logic,
    and CSR exemption for FP materials.

    EXTRA RULE (ONLY change):
    - For FP WITHOUT transfer AND with active CSR:
      If CSR ITEM (ZPLM CSR ITEM) has Material Status = 70 or 99
      in the SECOND 'Material Status' column (column D),
      THEN this validation is IGNORED for that material/plant.
    """

    needed = {'MRP Type', 'Plant Type', 'Planning Time Fence', 'Plant'}
    if not needed.issubset(df.columns):
        return

    has_psms = 'Plant Specific Material Status' in df.columns
    has_pc   = 'Planning Calendar' in df.columns
    valores_permitidos = {'M1', 'PD', 'P1', 'ND', 'V1'}

    # ------------------------------------------------------------
    # Transfer map
    # ------------------------------------------------------------
    transfer_map = {}
    if zplm_marc_df is not None:
        try:
            transfer_map = build_material_transfer_map(zplm_marc_df)
        except Exception:
            transfer_map = {}

    # ------------------------------------------------------------
    # Active CSR (I1 / I2 / I4)
    # ------------------------------------------------------------
    csr_exempt = build_csr_exempt_set(csr_item_df, csr_head_df)

    # ------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------
    def parse_tf(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return float(str(val).strip())
        except Exception:
            return 'INVALID'

    def parse_status(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return int(float(str(val).strip()))
        except Exception:
            return None

    # ------------------------------------------------------------
    # Main loop
    # ------------------------------------------------------------
    for idx, row in df.iterrows():
        erros = []

        mrp   = '' if pd.isna(row['MRP Type']) else str(row['MRP Type']).strip().upper()
        pt    = '' if pd.isna(row['Plant Type']) else str(row['Plant Type']).strip().upper()
        plant = '' if pd.isna(row['Plant']) else str(row['Plant']).strip()
        tf_raw = row.get('Planning Time Fence', None)
        tf_val = parse_tf(tf_raw)
        psms   = parse_status(row.get('Plant Specific Material Status')) if has_psms else None
        pc     = '' if (not has_pc or pd.isna(row.get('Planning Calendar'))) else str(row['Planning Calendar']).strip().upper()
        material = str(row.get('Material', '')).strip()

        is_transfer = transfer_map.get((material, plant), False)
        has_csr     = (material, plant) in csr_exempt

        # --------------------------------------------------------
        # ✅ NEW: ignore validation for FP without transfer
        # ✅ when CSR ITEM Material Status (column D) is 70 or 99
        # --------------------------------------------------------
        if pt == 'FP' and not is_transfer and has_csr:
            if csr_item_df is not None and not csr_item_df.empty:
                try:
                    csr_rows = csr_item_df[
                        (csr_item_df['Material'].astype(str).str.strip() == material) &
                        (csr_item_df['Plant'].astype(str).str.strip() == plant)
                    ]
                    if not csr_rows.empty:
                        # coluna D = índice 3 (segunda "Material Status")
                        csr_status_d = parse_status(csr_rows.iloc[0, 3])
                        if csr_status_d in {70, 99}:
                            continue  # 🔥 IGNORA TODA a validação
                except Exception:
                    pass

        # --------------------------------------------------------
        # PSMS precedence (INALTERADO)
        # --------------------------------------------------------
        if psms == 10:
            continue

        if psms == 25:
            if mrp != 'ND':
                erros.append(
                    "Mandatory field blank (MRP Type)" if mrp == '' else
                    f"For Part Status = 25, MRP Type must be 'ND' (Current: '{mrp}')"
                )
            if tf_val == 'INVALID':
                erros.append("Planning Time Fence invalid for Part Status = 25 (must be 0 or empty)")
            elif tf_val not in (None, 0):
                erros.append(
                    f"For Part Status = 25, Planning Time Fence must be 0 or empty (Current: '{tf_raw}')"
                )
            if has_pc:
                if plant == 'PT01' and pc != 'PT1':
                    erros.append(
                        f"For Part Status = 25 & PT01, Planning Calendar must be 'PT1' (Current: '{pc}')"
                    )
                elif plant == 'PT02' and pc != 'PT2':
                    erros.append(
                        f"For Part Status = 25 & PT02, Planning Calendar must be 'PT2' (Current: '{pc}')"
                    )
            if erros:
                resultados.append({
                    'Line': idx + 2,
                    'Field': 'MRP Type',
                    'Error': ' \n'.join(erros),
                    'Material': material,
                    'Plant': plant
                })
            continue

        if psms in (70, 75, 99):
            if mrp != 'ND':
                erros.append(
                    "Mandatory field blank (MRP Type)" if mrp == '' else
                    f"For Part Status = {psms}, MRP Type must be 'ND' (Current: '{mrp}')"
                )
            if erros:
                resultados.append({
                    'Line': idx + 2,
                    'Field': 'MRP Type',
                    'Error': ' \n'.join(erros),
                    'Material': material,
                    'Plant': plant
                })
            continue

        if psms == 65 and mrp == 'ND':
            continue

        # --------------------------------------------------------
        # Base rules (INALTERADO)
        # --------------------------------------------------------
        if mrp == '':
            erros.append("Mandatory field blank (MRP Type)")
        elif mrp not in valores_permitidos:
            erros.append(f"Invalid MRP Type: '{mrp}'")

        if tf_val == 'INVALID':
            erros.append("Planning Time Fence invalid (must be numeric)")
        else:
            has_tf = (tf_val is not None) and (tf_val > 0)
            no_tf  = (tf_val is None) or (tf_val == 0)

            if pt == 'FP' and not is_transfer:
                if has_csr and mrp == 'ND':
                    pass
                else:
                    if mrp != 'M1':
                        erros.append(f"For FP, MRP must be 'M1' (Current: '{mrp}')")
                    if not has_tf:
                        erros.append(
                            f"For FP, Time Fence must be > 0 (Current: '{tf_raw}')"
                        )

            elif pt in {'FP', 'MP'} and is_transfer:
                if mrp not in {'M1', 'P1', 'PD'}:
                    erros.append(
                        f"For {pt} with transfer, allowed MRP Types: M1, P1, PD (Current: '{mrp}')"
                    )

            elif pt == 'MP' and not is_transfer:
                if mrp != 'PD':
                    erros.append(
                        f"For MP without transfer, MRP must be 'PD' (Current: '{mrp}')"
                    )
                if not no_tf:
                    erros.append(
                        f"For MP without transfer, Time Fence must be 0 or empty (Current: '{tf_raw}')"
                    )

            elif pt == 'PP':
                if has_tf and mrp != 'P1':
                    erros.append(
                        f"For PP with Time Fence > 0, MRP must be 'P1' (Current: '{mrp}')"
                    )
                elif no_tf and mrp != 'PD':
                    erros.append(
                        f"For PP with Time Fence = 0/empty, MRP must be 'PD' (Current: '{mrp}')"
                    )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'MRP Type',
                'Error': ' \n'.join(erros),
                'Material': material,
                'Plant': plant
            })

def validar_price_control(df, resultados):
    """Price Control must always be 'S'."""
    if 'Price Control' not in df.columns:
        print("⚠️ Column 'Price Control' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        val = str(row['Price Control']).strip()
        if pd.isna(row['Price Control']) or val == '':
            erros.append('Mandatory field blank')
        elif val != 'S':
            erros.append(f"Invalid value: '{val}' (should be 'S')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Price Control',
                               'Error': ' | '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def validar_price_unit(df, resultados):
    """Price Unit must always be '10000'."""
    if 'Price Unit' not in df.columns:
        print("⚠️ Column 'Price Unit' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        val = str(row['Price Unit']).strip()
        if pd.isna(row['Price Unit']) or val == '':
            erros.append('Mandatory field blank')
        elif val != '10000':
            erros.append(f"Invalid value: '{val}' (should be '10000')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Price Unit',
                               'Error': ' | '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def _prepare_mvke(mvke_df, material_list):
    """
    Internal helper: normalises MVKE and filters to the requested materials.
    Returns (mvke_normalised, col_dict), or (None, None) if mandatory columns
    are missing.
    """
    mvke = mvke_df.copy()
    mvke.columns = [c.strip() for c in mvke.columns]

    def pick(candidates):
        return next((c for c in candidates if c in mvke.columns), None)

    cols = {
        'mat':   pick(['Material', 'MATNR']),
        'sorg':  pick(['Sales Organization', 'Sales Org', 'VKORG']),
        'dchan': pick(['Distribution Channel', 'Distr. Channel', 'VTWEG']),
        'aag':   pick(['Acct Assmt Grp Mat.', 'Account Assignment Group', 'KTGRM']),
        'icg':   pick(['Item Category Group', 'Item Cat Group', 'MVGR1', 'KDMAT']),
        'prodh': pick(['Product Hierarchy', 'PRODH', 'PRODH5']),
    }

    if any(cols[k] is None for k in ('mat', 'sorg', 'dchan', 'aag')):
        return None, None

    def norm_dc(x):
        try:
            return normalize_code(x, as_int=True)
        except Exception:
            return '' if pd.isna(x) else str(x).strip()

    mvke['_mat']   = mvke[cols['mat']].astype(str).str.strip()
    mvke['_sorg']  = mvke[cols['sorg']].astype(str).str.strip().str.upper()
    mvke['_dchan'] = mvke[cols['dchan']].apply(norm_dc)
    mvke['_aag']   = mvke[cols['aag']].apply(lambda x: '' if pd.isna(x) else str(x).strip())
    mvke['_icg']   = (mvke[cols['icg']].astype(str).str.strip().str.upper()
                      if cols['icg'] else pd.Series('', index=mvke.index))
    mvke['_prodh'] = (mvke[cols['prodh']].apply(lambda x: '' if pd.isna(x) else str(x).strip())
                      if cols['prodh'] else pd.Series('', index=mvke.index))

    mvke = mvke[mvke['_mat'].isin(material_list)].copy()
    return mvke, cols


def validar_product_hierarchy(df, resultados, mvke_df=None, mara_df=None):
    """
    Validates the Product Hierarchy field.

    SAPUI5 Export rules:
      - Mandatory for FERT only.
      - Position 13 (1-based) must be 'P'.
    """
    if 'Material Type' not in df.columns or 'Product Hierarchy' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Product Hierarchy' not found.")
        return

    # -- SAPUI5 mandatory / format checks (FERT only) -------------------------
    for idx, row in df.iterrows():
        if str(row['Material Type']).strip() != 'FERT':
            continue

        ph    = row['Product Hierarchy']
        ph_s  = '' if pd.isna(ph) else str(ph).strip()
        mat   = str(row.get('Material', '')).strip()
        plant = str(row.get('Plant',    '')).strip()

        if ph_s == '':
            resultados.append({
                'Line': idx + 2, 'Field': 'Product Hierarchy',
                'Error': "Mandatory field blank for Material Type 'FERT'",
                'Material': mat, 'Plant': plant,
            })
            continue

        if len(ph_s) >= 13 and ph_s[12] != 'P':
            resultados.append({
                'Line': idx + 2, 'Field': 'Product Hierarchy',
                'Error': f"Position 13 must be 'P' (Current: '{ph_s[12]}')",
                'Material': mat, 'Plant': plant,
            })


def validar_product_hierarchy_mvke(df, resultados, mvke_df):
    """
    Validates the Product Hierarchy field as carried in MVKE — a separate
    field from the MARA-level Product Hierarchy already validated by
    validar_product_hierarchy() above.

    Rule (FERT only, all plants):
      - MVKE carries one row per Sales Organization + Distribution Channel
        combination for a material, so Product Hierarchy may legitimately
        be populated on only one of a material's combinations.
      - Only flagged as an error if EVERY MVKE combination for that FERT
        material has Product Hierarchy blank (i.e. the field isn't
        populated on any combination for that material), or if the
        material has no MVKE entries at all.
      - Whichever combination(s) DO have a value must match the Product
        Hierarchy on the MARA record for the same material (the
        'Product Hierarchy' column already merged into `df` from MARA).
        Combinations with no value are not compared and do not count as
        errors.

    NOTE: `df` (sap_like_df) — not the raw global mara_df — is used to pick
    the FERT material list and the MARA-side comparison value, the same way
    validar_item_category_group()/validar_account_assignment_group() do.
    This keeps the check correctly scoped to whichever plant set the
    calling engine (Palmela or Tunisia) is running for: the raw mara_df is
    intentionally left unfiltered elsewhere in the codebase (for unrelated
    lookups), so using it here would flag every Palmela-only FERT material
    as "missing from MVKE" when this function runs inside the Tunisia
    engine (whose mvke_df is filtered to TN combinations only), and
    vice-versa.
    """
    for col in ['Material', 'Material Type', 'Product Hierarchy']:
        if col not in df.columns:
            return

    if mvke_df is None or mvke_df.empty:
        return

    sap_subset = df[df['Material Type'].astype(str).str.strip() == 'FERT'].copy()
    if sap_subset.empty:
        return

    sap_subset['Material'] = sap_subset['Material'].astype(str).str.strip()
    mara_ph_map = dict(zip(
        sap_subset['Material'],
        sap_subset['Product Hierarchy'].apply(lambda x: '' if pd.isna(x) else str(x).strip())
    ))
    fert_materials = set(mara_ph_map)

    mvke_norm, _ = _prepare_mvke(mvke_df, fert_materials)

    for material in sorted(fert_materials):
        mara_ph = mara_ph_map.get(material, '')
        grp = (mvke_norm[mvke_norm['_mat'] == material]
               if mvke_norm is not None and not mvke_norm.empty else None)

        if grp is None or grp.empty:
            resultados.append({
                'Line': '', 'Field': 'Product Hierarchy (MVKE)',
                'Error': ("Mandatory field blank in MVKE — no Sales Org/Distribution "
                          "Channel combination has this field populated (FERT material)"),
                'Material': material, 'Plant': '',
            })
            continue

        populated = grp[grp['_prodh'] != '']

        if populated.empty:
            resultados.append({
                'Line': '', 'Field': 'Product Hierarchy (MVKE)',
                'Error': ("Mandatory field blank in MVKE — no Sales Org/Distribution "
                          "Channel combination has this field populated (FERT material)"),
                'Material': material, 'Plant': '',
            })
            continue

        for _, row in populated.iterrows():
            combo   = f"SalesOrg='{row['_sorg']}' / DC='{row['_dchan']}'"
            mvke_ph = row['_prodh']
            if mvke_ph != mara_ph:
                resultados.append({
                    'Line': '', 'Field': 'Product Hierarchy (MVKE)',
                    'Error': (f"Product Hierarchy in MVKE ({combo}) = '{mvke_ph}' does not "
                              f"match Product Hierarchy in MARA = '{mara_ph}'"),
                    'Material': material, 'Plant': '',
                })


def validar_production_scheduling_profile(df, resultados):
    """FERT and HALB must have Production Scheduling Profile = 'ZREM01'."""
    if 'Material Type' not in df.columns or 'Production Scheduling Profile' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Production Scheduling Profile' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        mt  = str(row['Material Type']).strip()
        val = str(row['Production Scheduling Profile']).strip()
        if mt in ['FERT', 'HALB']:
            if pd.isna(row['Production Scheduling Profile']) or val == '':
                erros.append("Mandatory field blank For Material Type FERT or HALB")
            elif val != 'ZREM01':
                erros.append(f"Invalid value: '{val}' (should be 'ZREM01' For FERT or HALB)")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Production Scheduling Profile',
                               'Error': ' | '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def validar_profit_center(df, resultados):
    """PT01 -> 'PT00351'; PT02 -> 'PT00361'. Mandatory for all."""
    if 'Plant' not in df.columns or 'Profit Center' not in df.columns:
        print("⚠️ Columns 'Plant' or 'Profit Center' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        plant = str(row['Plant']).strip()
        val   = str(row['Profit Center']).strip()

        if pd.isna(row['Profit Center']) or val == '':
            erros.append('Mandatory field blank')
        else:
            if plant == 'PT01' and val != 'PT00351':
                erros.append(f"For Plant PT01, should be 'PT00351' (Current: '{val}')")
            elif plant == 'PT02' and val != 'PT00361':
                erros.append(f"For Plant PT02, should be 'PT00361' (Current: '{val}')")

        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Profit Center',
                               'Error': ' | '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant': plant})


def validar_purchasing_group(df, resultados):
    """Purchasing Group is mandatory only for Plant Type = PP.
    MP and FP should not have a Purchasing Group
    """
    if 'Plant Type' not in df.columns or 'Purchasing Group' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'Purchasing Group' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        if str(row['Plant Type']).strip() == 'PP':
            val = str(row['Purchasing Group']).strip()
            if pd.isna(row['Purchasing Group']) or val == '':
                erros.append("Mandatory field blank For Plant Type 'PP'")
        elif str(row['Plant Type']).strip() in ['MP', 'FP']:
            val = str(row['Purchasing Group']).strip()
            if not pd.isna(row['Purchasing Group']) and val != '':
                erros.append("Field should be blank For Plant Type 'MP' or 'FP'")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Purchasing Group',
                               'Error': ' | '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def validar_repetitive_manufacturing(marc_df, zplm_marc_df, mara_df, resultados):
    """
    Repetitive Manufacturing rules by Material Plant Type (from ZPLM MARC),
    only for materials with Material Type in MARA = FERT/HALB/ROH:
    - PP    -> field must be empty
    - MP/FP -> field must be 'X'
    """
    required_marc = {'Material', 'Plant', 'Repetitive Manufacturing'}
    required_mara = {'Material', 'Material Type'}

    if not required_marc.issubset(marc_df.columns):
        print("⚠️ Required columns not found in MARC.")
        return

    if mara_df is None or mara_df.empty or not required_mara.issubset(mara_df.columns):
        print("⚠️ Required columns not found in MARA.")
        return

    if zplm_marc_df is None or zplm_marc_df.empty:
        print("⚠️ ZPLM MARC dataframe not provided or empty.")
        return

    # Mapa Material -> Material Type (MARA)
    mt_map = (
        mara_df[['Material', 'Material Type']]
        .dropna(subset=['Material'])
        .assign(
            Material=lambda d: d['Material'].astype(str).str.strip(),
            Material_Type=lambda d: d['Material Type'].astype(str).str.strip().str.upper()
        )
        .set_index('Material')['Material_Type']
        .to_dict()
    )

    zplm = zplm_marc_df.copy()
    zplm.columns = [str(c).strip() for c in zplm.columns]

    material_col = next((c for c in ['Material', 'Material Number'] if c in zplm.columns), None)
    plant_col    = next((c for c in ['Plant', 'Plant Code'] if c in zplm.columns), None)
    mpt_col      = next((c for c in ['Material Plant Type'] if c in zplm.columns), None)

    if not material_col or not plant_col or not mpt_col:
        print("⚠️ Required columns not found in ZPLM MARC.")
        return

    # mapa (Material, Plant) -> Material Plant Type
    zplm['__MAT__'] = zplm[material_col].fillna('').astype(str).str.strip()
    zplm['__PLANT__'] = zplm[plant_col].fillna('').astype(str).str.strip()
    zplm['__MPT__'] = zplm[mpt_col].fillna('').astype(str).str.strip().str.upper()

    mpt_map = (
        zplm.drop_duplicates(subset=['__MAT__', '__PLANT__'], keep='last')
            .set_index(['__MAT__', '__PLANT__'])['__MPT__']
            .to_dict()
    )

    for idx, row in marc_df.iterrows():
        erros = []

        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()

        # só valida FERT/HALB/ROH
        mt = mt_map.get(material, '')
        if mt not in {'FERT', 'HALB', 'ROH'}:
            continue

        mpt = mpt_map.get((material, plant), '')
        repet = '' if pd.isna(row.get('Repetitive Manufacturing')) else str(row.get('Repetitive Manufacturing')).strip()

        if not mpt:
            erros.append(f"Material Plant Type not found in ZPLM MARC for Material '{material}' / Plant '{plant}'")
        else:
            if mpt == 'PP':
                if repet != '':
                    erros.append(
                        f"For Material Plant Type 'PP', field 'Repetitive Manufacturing' must be empty "
                        f"(Current: '{repet}')"
                    )
            elif mpt in ['MP', 'FP']:
                if repet != 'X':
                    erros.append(
                        f"For Material Plant Type '{mpt}', field 'Repetitive Manufacturing' should be 'X' "
                        f"(Current: '{repet}')"
                    )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Repetitive Manufacturing',
                'Error': ' \n '.join(erros),
                'Material': material,
                'Plant': plant
            })

def validar_rem_profile(marc_df, zplm_marc_df, mara_df, resultados):
    """
    REM Profile rules by Material Plant Type (from ZPLM MARC),
    only for materials with Material Type in MARA = FERT/HALB/ROH:
    - PP    -> field must be empty
    - MP/FP -> field must be 'Z002'
    """
    required_marc = {'Material', 'Plant', 'REM Profile'}
    required_mara = {'Material', 'Material Type'}

    if not required_marc.issubset(marc_df.columns):
        print("⚠️ Required columns not found in MARC.")
        return

    if mara_df is None or mara_df.empty or not required_mara.issubset(mara_df.columns):
        print("⚠️ Required columns not found in MARA.")
        return

    if zplm_marc_df is None or zplm_marc_df.empty:
        print("⚠️ ZPLM MARC dataframe not provided or empty.")
        return

    # Mapa Material -> Material Type (MARA)
    mt_map = (
        mara_df[['Material', 'Material Type']]
        .dropna(subset=['Material'])
        .assign(
            Material=lambda d: d['Material'].astype(str).str.strip(),
            Material_Type=lambda d: d['Material Type'].astype(str).str.strip().str.upper()
        )
        .set_index('Material')['Material_Type']
        .to_dict()
    )

    zplm = zplm_marc_df.copy()
    zplm.columns = [str(c).strip() for c in zplm.columns]

    material_col = next((c for c in ['Material', 'Material Number'] if c in zplm.columns), None)
    plant_col    = next((c for c in ['Plant', 'Plant Code'] if c in zplm.columns), None)
    mpt_col      = next((c for c in ['Material Plant Type'] if c in zplm.columns), None)

    if not material_col or not plant_col or not mpt_col:
        print("⚠️ Required columns not found in ZPLM MARC.")
        return

    # mapa (Material, Plant) -> Material Plant Type
    zplm['__MAT__'] = zplm[material_col].fillna('').astype(str).str.strip()
    zplm['__PLANT__'] = zplm[plant_col].fillna('').astype(str).str.strip()
    zplm['__MPT__'] = zplm[mpt_col].fillna('').astype(str).str.strip().str.upper()

    mpt_map = (
        zplm.drop_duplicates(subset=['__MAT__', '__PLANT__'], keep='last')
            .set_index(['__MAT__', '__PLANT__'])['__MPT__']
            .to_dict()
    )

    for idx, row in marc_df.iterrows():
        erros = []

        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()

        # só valida FERT/HALB/ROH
        mt = mt_map.get(material, '')
        if mt not in {'FERT', 'HALB', 'ROH'}:
            continue

        mpt = mpt_map.get((material, plant), '')
        rem = '' if pd.isna(row.get('REM Profile')) else str(row.get('REM Profile')).strip()

        if not mpt:
            erros.append(f"Material Plant Type not found in ZPLM MARC for Material '{material}' / Plant '{plant}'")
        else:
            if mpt == 'PP':
                if rem != '':
                    erros.append(
                        f"For Material Plant Type 'PP', field 'REM Profile' must be empty "
                        f"(Current: '{rem}')"
                    )
            elif mpt in ['MP', 'FP']:
                if rem != 'Z002':
                    erros.append(
                        f"For Material Plant Type '{mpt}', field 'REM Profile' should be 'Z002' "
                        f"(Current: '{rem}')"
                    )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'REM Profile',
                'Error': ' \n '.join(erros),
                'Material': material,
                'Plant': plant
            })

def validar_source_list_required_marc(marc_df, mara_df, resultados):
    """
    For materials FERT, HALB, ROH:
    MARC['Source list'] must NOT be 'X'
    """

    if marc_df is None or mara_df is None:
        print("⚠️ MARC or MARA not provided.")
        return

    if 'Material' not in marc_df.columns or 'Source list' not in marc_df.columns:
        print("⚠️ Columns 'Material' or 'Source list' not found in MARC.")
        return

    if 'Material' not in mara_df.columns or 'Material Type' not in mara_df.columns:
        print("⚠️ Columns 'Material' or 'Material Type' not found in MARA.")
        return

    # Map Material → Material Type
    mt_map = dict(zip(
        mara_df['Material'].astype(str).str.strip(),
        mara_df['Material Type'].astype(str).str.strip().str.upper()
    ))

    for idx, row in marc_df.iterrows():
        material = str(row.get('Material', '')).strip()
        plant = str(row.get('Plant', '')).strip()
        val = str(row.get('Source list', '')).strip().upper()

        mt = mt_map.get(material, '')

        # Only validate FERT / HALB / ROH
        if mt not in {'FERT', 'HALB', 'ROH'}:
            continue

        if val == 'X':
            resultados.append({
                'Line': idx + 2,
                'Field': 'Source list',
                'Error': "Must NOT be 'X' for FERT/HALB/ROH",
                'Material': material,
                'Plant': plant
            })

def validar_storage_location(df, resultados):
    """
    Issue Storage Location validation for Palmela.

    PT01, Material Types FERT/HALB/ROH:
      - Plant Type FP     -> must be 'P120'
      - Plant Type MP/PP  -> must be 'P105'

    PT02, Material Types FERT/HALB/ROH and Plant Types FP/MP/PP:
      - Default           -> must be 'P205'
      - Exception         -> 'P220' is also allowed when PSMS = 25
    """
    required = {
        'Plant',
        'Plant Type',
        'Issue Storage Location',
        'Material Type',
    }
    if not required.issubset(df.columns):
        missing = sorted(required - set(df.columns))
        print(f"⚠️ Missing columns for Issue Storage Location validation: {missing}")
        return

    valid_material_types = {'FERT', 'HALB', 'ROH'}
    valid_plant_types = {'FP', 'MP', 'PP'}

    for idx, row in df.iterrows():
        erros = []
        plant = str(row.get('Plant', '')).strip().upper()
        ptype = str(row.get('Plant Type', '')).strip().upper()
        mt = str(row.get('Material Type', '')).strip().upper()
        raw = row.get('Issue Storage Location')
        val = '' if pd.isna(raw) else str(raw).strip().upper()
        material = str(row.get('Material', '')).strip()

        # This rule only applies to the explicitly defined material/plant types.
        if mt not in valid_material_types or ptype not in valid_plant_types:
            continue

        if plant == 'PT01':
            expected = 'P120' if ptype == 'FP' else 'P105'
            if val == '':
                erros.append(
                    f"Mandatory field blank (PT01 + Plant Type '{ptype}' requires '{expected}')"
                )
            elif val != expected:
                erros.append(
                    f"For PT01 + Material Type '{mt}' + Plant Type '{ptype}', "
                    f"Issue Storage Location must be '{expected}' (Current: '{val}')"
                )

        elif plant == 'PT02':
            psms_raw = row.get('Plant Specific Material Status', '')
            psms = '' if pd.isna(psms_raw) else str(psms_raw).strip()
            try:
                psms = str(int(float(psms)))
            except Exception:
                pass

            if val == '':
                erros.append("Mandatory field blank (PT02 requires 'P205', or 'P220' when PSMS = 25)")
            elif val == 'P220':
                if psms not in ['25', '75']:
                    erros.append(
                        "Value 'P220' is only allowed for PT02 when Plant Specific Material Status "
                        f"is 25 or 75(Current PSMS: '{psms}')"
                    )
            elif val != 'P205':
                erros.append(
                    "For PT02, Issue Storage Location must be 'P205', or 'P220' "
                    f"when PSMS = 25 (Current: '{val}')"
                )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Issue Storage Location',
                'Error': ' | '.join(erros),
                'Material': material,
                'Plant': plant,
            })


def validar_valuation_class(df, resultados):
    """
    Valuation Class must match Material Type:
    - '3000' -> ROH only
    - '7900' -> HALB only
    - '7920' -> FERT only
    """
    if 'Material Type' not in df.columns or 'Valuation Class' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Valuation Class' not found.")
        return

    for idx, row in df.iterrows():
        erros = []

        material = str(row.get('Material', '')).strip()
        real_mt  = str(row.get('Material Type', '')).strip()

        # ✅ Material Type efetivo por regra
        mt = get_effective_material_type(
            material,
            real_mt,
            'validar_valuation_class',
        )

        val = normalize_code(row.get('Valuation Class'), as_int=True)

        if val == '':
            erros.append('Mandatory field blank')
        else:
            if val == '3000' and mt != 'ROH':
                erros.append(
                    f"Value '3000' only allowed for Material Type = ROH (Current: '{mt}')"
                )
            elif val == '7900' and mt != 'HALB':
                erros.append(
                    f"Value '7900' only allowed for Material Type = HALB (Current: '{mt}')"
                )
            elif val == '7920' and mt != 'FERT':
                erros.append(
                    f"Value '7920' only allowed for Material Type = FERT (Current: '{mt}')"
                )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Valuation Class',
                'Error': ' | '.join(erros),
                'Material': material,
                'Plant': str(row.get('Plant', '')).strip(),
            })

def validar_with_qty_structure(df, resultados):
    """With Qty Structure must always be 'X' (SAP checkbox logic)."""

    if 'With Qty Structure' not in df.columns:
        print("⚠️ Column 'With Qty Structure' not found.")
        return

    for idx, row in df.iterrows():
        erros = []

        raw_val = row.get('With Qty Structure')
        val = str(raw_val).strip().upper() if pd.notna(raw_val) else ''

        # ✅ Normalizar checkbox
        if val in {'X', '1', 'YES'}:
            val_norm = 'X'
        else:
            val_norm = ''

        # ✅ Regra
        if val_norm == '':
            erros.append('Mandatory field blank')
        elif val_norm != 'X':
            erros.append(f"Invalid value: '{val}' (should be 'X')")

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'With Qty Structure',
                'Error': ' | '.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant': str(row.get('Plant', '')).strip()
            })

def validar_backflush(df, resultados):
    """Backflush is mandatory."""
    if 'Backflush' not in df.columns:
        return
    for idx, row in df.iterrows():
        if pd.isna(row['Backflush']) or str(row['Backflush']).strip() == '':
            resultados.append({'Line': idx + 2, 'Field': 'Backflush',
                               'Error': 'Mandatory field blank',
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def validar_strategy_group(df, resultados):
    """Strategy Group must always be '40'."""
    if 'Strategy Group' not in df.columns:
        return
    for idx, row in df.iterrows():
        erros = []
        raw = row['Strategy Group']
        val = normalize_code(raw, as_int=True)
        if val == '':
            erros.append('Mandatory field blank')
        elif val != '40':
            erros.append(f"Invalid value: '{raw}' (should be '40')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Strategy Group',
                               'Error': ' \n'.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def validar_product_related_origin(df, resultados):
    """Product-Related Origin is mandatory."""
    if 'Product-Related Origin' not in df.columns:
        return
    for idx, row in df.iterrows():
        if pd.isna(row['Product-Related Origin']) or str(row['Product-Related Origin']).strip() == '':
            resultados.append({'Line': idx + 2, 'Field': 'Product-Related Origin',
                               'Error': 'Mandatory field blank',
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def validar_availability_check(df, resultados):
    """Availability Check is mandatory."""
    if 'Availability Check' not in df.columns:
        return
    for idx, row in df.iterrows():
        if pd.isna(row['Availability Check']) or str(row['Availability Check']).strip() == '':
            resultados.append({'Line': idx + 2, 'Field': 'Availability Check',
                               'Error': 'Mandatory field blank',
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def _build_mard_stock_lookup(mard_df):
    """
    Builds a {(Material, Plant): has_stock (bool)} lookup from MARD.

    Replicates the exact same stock-column detection and aggregation logic
    used by the Data Cleanups rule `validar_mard_total_stock`: a
    Material+Plant combination is considered to have stock when the sum
    across Unrestricted, Stock in transfer, Quality Inspection, Blocked and
    Returns is > 0.

    Returns None when MARD is missing/empty or lacks usable
    Material/Plant/stock columns, so callers can tell "no stock information
    available" apart from "no stock" (an empty-but-usable lookup, where
    every combination legitimately defaults to no stock, is returned as {}).
    """
    if mard_df is None or mard_df.empty:
        return None
    if 'Material' not in mard_df.columns or 'Plant' not in mard_df.columns:
        return None

    import re as _re

    def _norm(s):
        return _re.sub(r'[^a-z0-9]+', '', str(s).lower())

    def _pick_first_cols(df, wanted_labels):
        out = {k: None for k in wanted_labels}
        norm_cols = [_norm(c) for c in df.columns]
        seen = set()
        for key, patterns in wanted_labels.items():
            for i, col in enumerate(df.columns):
                if col in seen:
                    continue
                for pat in patterns:
                    if norm_cols[i].startswith(_norm(pat)):
                        out[key] = col
                        seen.add(col)
                        break
                if out[key] is not None:
                    break
        return out

    mard = mard_df.copy()
    mard['Material'] = mard['Material'].astype(str).str.strip()
    mard['Plant'] = mard['Plant'].astype(str).str.strip()

    wanted = {
        "unrestricted": ["Unrestricted"],
        "stock_transfer": ["Stock in transfer"],
        "quality": ["Quality Inspection"],
        "blocked": ["Blocked"],
        "returns": ["Returns"],
    }
    colmap = _pick_first_cols(mard, wanted)
    stock_cols = [c for c in colmap.values() if c is not None]
    if not stock_cols:
        return None

    mard_use = mard[['Material', 'Plant'] + stock_cols].copy()
    for c in stock_cols:
        mard_use[c] = pd.to_numeric(mard_use[c], errors='coerce').fillna(0)

    row_has_stock = None
    for c in stock_cols:
        row_has_stock = (mard_use[c] > 0) if row_has_stock is None else (row_has_stock | (mard_use[c] > 0))
    mard_use['_row_has_stock'] = row_has_stock.fillna(False)

    agg = mard_use.groupby(['Material', 'Plant'])['_row_has_stock'].max()
    return {idx: bool(val) for idx, val in agg.items()}


def validar_mrp_controller(df, resultados, mard_df=None):
    """
    MRP Controller validation rules:

    - MRP Controller is mandatory.
    - PSMS 70 -> MRP Controller must be 'OB1' when the Material+Plant has
      NO stock in MARD. When stock exists, this rule does not enforce a
      specific value (the mandatory-field check above still applies).
    - PSMS 75 -> same logic, but requires 'OB2' when there is no stock.
    - PSMS 99       -> MRP Controller must be 'DS1' (not stock-dependent).
    - MRP Controller 'Z01' is allowed only when PSMS is 00 or 10.

    Stock is determined via `_build_mard_stock_lookup`, which mirrors the
    Data Cleanups "MARD Total Stock" rule: Unrestricted / Stock in transfer
    / Quality Inspection / Blocked / Returns, summed per (Material, Plant).

    `mard_df` is optional for backward compatibility: when it is not
    supplied (e.g. MARD wasn't uploaded), the legacy behaviour is used —
    PSMS 70 always requires 'OB1' and PSMS 75 always requires 'OB2',
    regardless of stock.
    """
    if 'MRP Controller' not in df.columns:
        return

    has_psms = 'Plant Specific Material Status' in df.columns
    # None => no MARD info available (legacy behaviour); {} or populated
    # dict => usable stock info, missing combinations default to no stock.
    stock_lookup = _build_mard_stock_lookup(mard_df)

    def parse_status(val):
        """
        Normalizes SAP/Excel status values:
        '00'   -> 0
        '10'   -> 10
        70.0   -> 70
        blank  -> None
        invalid value -> None
        """
        if pd.isna(val) or str(val).strip() == '':
            return None

        try:
            return int(float(str(val).strip()))
        except (ValueError, TypeError):
            return None

    for idx, row in df.iterrows():
        erros = []

        mrp_ctrl = (
            ''
            if pd.isna(row['MRP Controller'])
            else str(row['MRP Controller']).strip().upper()
        )

        psms = (
            parse_status(row.get('Plant Specific Material Status'))
            if has_psms
            else None
        )

        material = str(row.get('Material', '')).strip()
        plant = str(row.get('Plant', '')).strip()

        # Rule 1: Mandatory field
        if mrp_ctrl == '':
            erros.append('Mandatory field blank')

        # Rule 2: PSMS 70 requires OB1 — only when the material has NO stock
        # (per MARD). When stock exists, no specific value is enforced here.
        # If MARD wasn't supplied, fall back to the legacy behaviour (always
        # requires OB1).
        elif psms == 70:
            if stock_lookup is not None:
                has_stock = stock_lookup.get((material, plant), False)
                if not has_stock and mrp_ctrl != 'OB1':
                    erros.append(
                        f"Invalid MRP Controller '{mrp_ctrl}': "
                        f"Material Status 70 with no stock requires 'OB1'"
                    )
            elif mrp_ctrl != 'OB1':
                erros.append(
                    f"Invalid MRP Controller '{mrp_ctrl}': "
                    f"Material Status {psms} requires 'OB1'"
                )

        # Rule 3: PSMS 75 requires OB2 — only when the material has NO stock
        # (per MARD). Same fallback as PSMS 70 when MARD isn't supplied.
        elif psms == 75:
            if stock_lookup is not None:
                has_stock = stock_lookup.get((material, plant), False)
                if not has_stock and mrp_ctrl != 'OB2':
                    erros.append(
                        f"Invalid MRP Controller '{mrp_ctrl}': "
                        f"Material Status 75 with no stock requires 'OB2'"
                    )
            elif mrp_ctrl != 'OB2':
                erros.append(
                    f"Invalid MRP Controller '{mrp_ctrl}': "
                    f"Material Status {psms} requires 'OB2'"
                )

        # Rule 4: PSMS 99 requires DS1 (not stock-dependent)
        elif psms == 99 and mrp_ctrl != 'DS1':
            erros.append(
                f"Invalid MRP Controller '{mrp_ctrl}': "
                "Material Status 99 requires 'DS1'"
            )

        # Rule 5: Z01 is allowed only for PSMS 00 or 10
        elif mrp_ctrl == 'Z01' and psms not in {0, 10}:
            erros.append(
                "MRP Controller 'Z01' is allowed only when "
                "Material Status is 00 or 10"
            )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'MRP Controller',
                'Error': ' \n'.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant': str(row.get('Plant', '')).strip(),
            })

def validar_plant_specific_material_status(df, resultados):
    """Plant Specific Material Status is mandatory."""
    if 'Plant Specific Material Status' not in df.columns:
        return
    for idx, row in df.iterrows():
        if pd.isna(row['Plant Specific Material Status']) or str(row['Plant Specific Material Status']).strip() == '':
            resultados.append({'Line': idx + 2, 'Field': 'Plant Specific Material Status',
                               'Error': 'Mandatory field blank',
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


# =============================================================================
# B) CROSS-FILE VALIDATIONS (SAPUI5 + MVKE / MARC / MARA)
# =============================================================================

def validar_account_assignment_group(df, resultados, mvke_df):
    """
    For FERT, ROH and HALB validates two things in one pass over MVKE:

    1. Combination structure:
       - FERT must have EXACTLY PT00/10, PT00/20 and US00/30.
         Missing or unexpected combinations are each reported individually.
         If no MVKE entries exist at all for a FERT, all 3 are reported missing.
       - ROH/HALB: if US00/30 is present, PT00/20 must also be present.

    2. Account Assignment Group value:
       For every existing MVKE entry the AAG must be 'Z1'.
    """
    for col in ['Material', 'Material Type']:
        if col not in df.columns:
            return

    valid_types = {'FERT', 'HALB', 'ROH'}
    sap_subset  = df[df['Material Type'].astype(str).str.strip().isin(valid_types)].copy()
    sap_subset['Material'] = sap_subset['Material'].astype(str).str.strip()

    if sap_subset.empty or mvke_df is None or mvke_df.empty:
        return

    mt_map       = dict(zip(sap_subset['Material'],
                            sap_subset['Material Type'].astype(str).str.strip()))
    material_set = set(mt_map)

    mvke_norm, _ = _prepare_mvke(mvke_df, material_set)
    materials_in_mvke = (set(mvke_norm['_mat'].unique())
                         if mvke_norm is not None and not mvke_norm.empty else set())

    FERT_REQUIRED = {('PT00', '10'), ('PT00', '20'), ('US00', '30')}
    # Tunisia FERT materials additionally carry TN00/10 + TN00/20 MVKE entries
    # (already validated separately and correctly by tn_validar_account_assignment_group).
    # These are legitimate combinations from Palmela's point of view too — not
    # required here, but must not be flagged as "unexpected" when present.
    FERT_ALLOWED_EXTRA = {('TN00', '10'), ('TN00', '20')}

    def _err(material, msg):
        resultados.append({
            'Line': '', 'Field': 'Account Assignment Group',
            'Error': msg, 'Material': material, 'Plant': '',
        })

    # --- 1. Combination structure ---
    for material, mt in mt_map.items():
        if material not in materials_in_mvke:
            if mt == 'FERT':
                _err(material,
                     "FERT: no entries found in MVKE — "
                     "required combinations PT00/10, PT00/20, US00/30 are all missing")
            continue

        grp    = mvke_norm[mvke_norm['_mat'] == material]
        combos = set(zip(grp['_sorg'], grp['_dchan']))

        if mt == 'FERT':
            for sorg, dchan in sorted(FERT_REQUIRED - combos):
                _err(material,
                     f"FERT: missing required MVKE combination "
                     f"SalesOrg='{sorg}' / DC='{dchan}'")
            for sorg, dchan in sorted(combos - FERT_REQUIRED - FERT_ALLOWED_EXTRA):
                _err(material,
                     f"FERT: unexpected MVKE combination "
                     f"SalesOrg='{sorg}' / DC='{dchan}' "
                     f"(only PT00/10, PT00/20, US00/30 are allowed)")
        elif mt in ('ROH', 'HALB'):
            if ('US00', '30') in combos and ('PT00', '20') not in combos:
                _err(material,
                     f"{mt}: has MVKE combination US00/30 but is missing PT00/20")

    # --- 2. AAG value per entry ---
    if mvke_norm is None or mvke_norm.empty:
        return

    for _, row in mvke_norm.iterrows():
        combo = f"SalesOrg='{row['_sorg']}' / DC='{row['_dchan']}'"
        aag   = row['_aag']

        if aag == '':
            _err(row['_mat'], f"Mandatory field blank in MVKE ({combo})")
        elif aag.upper() != 'Z1':
            _err(row['_mat'],
                 f"Invalid value '{aag}' in MVKE ({combo}) — must be 'Z1'")


def validar_item_category_group(df, resultados, mvke_df):
    """
    For every MVKE entry of FERT, ROH and HALB:
    Item Category Group must be 'NORM'.
    """
    for col in ['Material', 'Material Type']:
        if col not in df.columns:
            return

    valid_types = {'FERT', 'HALB', 'ROH'}
    sap_subset  = df[df['Material Type'].astype(str).str.strip().isin(valid_types)].copy()
    sap_subset['Material'] = sap_subset['Material'].astype(str).str.strip()

    if sap_subset.empty or mvke_df is None or mvke_df.empty:
        return

    material_set = set(sap_subset['Material'])
    mvke_norm, _ = _prepare_mvke(mvke_df, material_set)
    if mvke_norm is None or mvke_norm.empty:
        return

    for _, row in mvke_norm.iterrows():
        combo = f"SalesOrg='{row['_sorg']}' / DC='{row['_dchan']}'"
        icg   = row['_icg']

        if icg == '' or icg.upper() in ('NAN', 'NONE'):
            resultados.append({
                'Line': '', 'Field': 'Item Category Group',
                'Error': f"Mandatory field blank in MVKE ({combo})",
                'Material': row['_mat'], 'Plant': '',
            })
        elif icg != 'NORM':
            resultados.append({
                'Line': '', 'Field': 'Item Category Group',
                'Error': f"Invalid value '{icg}' in MVKE ({combo}) — must be 'NORM'",
                'Material': row['_mat'], 'Plant': '',
            })


def validar_consumption_per(marc_df, mara_df, resultados):
    """
    MARC: Bwd consumption per. and Fwd consumption per. must be 0
    for FERT, HALB, ROH materials.
    """
    needed = {'Material', 'Plant', 'Bwd consumption per.', 'Fwd consumption per.'}
    if not needed.issubset(marc_df.columns):
        print("⚠️ MARC is missing required columns for Bwd/Fwd consumption per. validation.")
        return

    mt_map = dict(zip(
        mara_df['Material'].astype(str).str.strip(),
        mara_df['Material Type'].astype(str).str.strip().str.upper()
    ))

    def norm_val(val):
        try:
            return None if pd.isna(val) else float(val)
        except Exception:
            return None

    for idx, row in marc_df.iterrows():
        material = str(row['Material']).strip()
        plant    = str(row['Plant']).strip()
        mt       = mt_map.get(material, '')

        if mt not in {'FERT', 'HALB', 'ROH'}:
            continue

        bwd = norm_val(row.get('Bwd consumption per.'))
        fwd = norm_val(row.get('Fwd consumption per.'))

        if bwd not in (0, None):
            resultados.append({'Line': idx + 2, 'Field': 'Bwd consumption per.',
                               'Error': f"Invalid value: '{row.get('Bwd consumption per.')}' (must be 0 for {mt})",
                               'Material': material, 'Plant': plant})
        if fwd not in (0, None):
            resultados.append({'Line': idx + 2, 'Field': 'Fwd consumption per.',
                               'Error': f"Invalid value: '{row.get('Fwd consumption per.')}' (must be 0 for {mt})",
                               'Material': material, 'Plant': plant})


def validar_enterproj_number(df, resultados):
    """FERT -> Enterproj Number required; ROH / HALB -> must be empty."""
    if 'Material Type' not in df.columns or 'Enterproj Number' not in df.columns:
        return

    for idx, row in df.iterrows():
        erros = []

        material = str(row.get('Material', '')).strip()
        real_mt  = str(row.get('Material Type', '')).strip()

        # ✅ Material Type efetivo por regra
        mt = get_effective_material_type(
            material,
            real_mt,
            'validar_enterproj_number',
        )

        enterproj = '' if pd.isna(row['Enterproj Number']) else str(row['Enterproj Number']).strip()

        if mt == 'FERT':
            if enterproj == '':
                erros.append(
                    "For Material Type 'FERT', the field is required and cannot be empty."
                )

        elif mt in ('ROH', 'HALB'):
            if enterproj != '':
                display = (
                    str(int(float(enterproj)))
                    if enterproj.replace('.', '', 1).isdigit()
                    else enterproj
                )
                erros.append(
                    f"For Material Type '{mt}', the field must be empty (current: '{display}')"
                )

        if erros and not any(
            r['Material'] == material and r['Field'] == 'Enterproj Number'
            for r in resultados
        ):
            resultados.append({
                'Line': idx + 2,
                'Field': 'Enterproj Number',
                'Error': ' | '.join(erros),
                'Material': material,
                'Plant': '',
            })


def validar_product_hierarchy_vs_enterproj(mara_df, zplm_mara_df, resultados):
    """
    For FERT: validates that the Product Hierarchy suffix (after position 13/'P')
    matches the Enterproj Number from ZPLM MARA.
    """
    needed_mara = {'Material', 'Material Type', 'Product Hierarchy'}
    if not needed_mara.issubset(mara_df.columns):
        return
    if zplm_mara_df is None or zplm_mara_df.empty:
        return

    zplm = zplm_mara_df.copy()
    zplm.columns = [c.strip().upper() for c in zplm.columns]
    if not {'MATERIAL', 'ENTERPROJ NUMBER'}.issubset(zplm.columns):
        return

    zplm_map = dict(zip(zplm['MATERIAL'].astype(str).str.strip(), zplm['ENTERPROJ NUMBER']))

    def is_nan_like(v):
        return v is None or (isinstance(v, float) and math.isnan(v))

    def norm_side(value, target_len=None):
        if is_nan_like(value):
            return ("0" * target_len) if (target_len and target_len > 0) else "0"
        s = str(value).strip()
        if s.lower() in ("nan", "none", ""):
            return ("0" * target_len) if (target_len and target_len > 0) else "0"
        try:
            s = str(int(float(s)))
        except Exception:
            pass
        if target_len and target_len > 0:
            s = s.zfill(target_len)
        return s

    for idx, row in mara_df.iterrows():
        material = str(row['Material']).strip()
        if str(row['Material Type']).strip().upper() != 'FERT':
            continue
        ph_raw = str(row['Product Hierarchy']).strip()
        if pd.isna(ph_raw) or ph_raw == '' or len(ph_raw) < 13:
            continue
        if ph_raw[12] != 'P':
            resultados.append({'Line': '', 'Field': 'Product Hierarchy', 'Material': material, 'Plant': '',
                               'Error': "Invalid Product Hierarchy: missing 'P' at 13th position (FERT rule)"})
            continue
        suffix = ph_raw[13:].strip()
        ph_norm = norm_side(suffix, target_len=len(suffix))
        ep_norm = norm_side(zplm_map.get(material, None), target_len=len(suffix))
        if ph_norm != ep_norm:
            resultados.append({'Line': '', 'Field': 'Product Hierarchy', 'Material': material, 'Plant': '',
                               'Error': f"PH mismatch: ProductHierarchy→'{ph_norm}' ≠ Enterproj→'{ep_norm}'"})


def validar_marm_pc_ea_roh_halb(marm_df, mara_df, resultados):
    """
    MARM validation (material-level, single error per material)

    Applies ONLY when:
      - Material Type ∈ {ROH, HALB}
      - Base Unit of Measure (MARA) = 'EA'

    Rule:
      The material MUST have at least ONE Alternative Unit of Measure = 'PC'
      with Numerator = 1 and Denominator = 1.

    Behaviour:
      - One single error per material
      - Error message includes:
          * Material Type
          * Alternative UoM(s) found in MARM
      - No duplicated errors
    """

    if marm_df is None or marm_df.empty or mara_df is None or mara_df.empty:
        return

    required_marm = {
        'Material',
        'Alternative Unit of Measure',
        'Numerator',
        'Denominator'
    }

    required_mara = {
        'Material',
        'Material Type',
        'Base Unit of Measure'
    }

    if not required_marm.issubset(marm_df.columns):
        print("⚠️ MARM missing required columns for PC validation.")
        return

    if not required_mara.issubset(mara_df.columns):
        print("⚠️ MARA missing required columns for PC validation.")
        return

    # Normalise
    marm = marm_df.copy()
    mara = mara_df.copy()

    marm.columns = marm.columns.astype(str).str.strip()
    mara.columns = mara.columns.astype(str).str.strip()

    # Filter MARA to relevant materials
    mara_relevant = mara[
    mara['Material Type'].astype(str).str.strip().str.upper().isin({'ROH', 'HALB'}) & (mara['Base Unit of Measure'].astype(str).str.strip().str.upper() == 'EA')
    ].copy()

    if mara_relevant.empty:
        return

    # Group MARM by Material
    marm_groups = marm.groupby(marm['Material'].astype(str).str.strip())

    for _, mara_row in mara_relevant.iterrows():
        material = str(mara_row['Material']).strip()
        material_type = str(mara_row['Material Type']).strip().upper()

        # Material not present in MARM
        if material not in marm_groups.groups:
            error_msg = (
                f"Material Type: {material_type} | "
                f"Alternative UoM found: none | "
                "Expected at least one Alternative Unit of Measure 'PC' "
                "with Numerator = 1 and Denominator = 1"
            )

            resultados.append({
                'Line': '',
                'Field': 'Alternative UoM',
                'Error': error_msg,
                'Material': material,
                'Plant': ''
            })
            continue

        df_m = marm_groups.get_group(material)

        # List all Alternative UoM found
        alt_uoms = (
            df_m['Alternative Unit of Measure']
            .astype(str)
            .str.strip()
            .str.upper()
            .dropna()
            .unique()
            .tolist()
        )

        alt_uoms_str = ", ".join(sorted(alt_uoms)) if alt_uoms else "none"
        base_msg = f"Material Type: {material_type} | Alternative UoM found: {alt_uoms_str}"

        # Filter PC rows
        pc_rows = df_m[
            df_m['Alternative Unit of Measure']
            .astype(str).str.strip().str.upper() == 'PC'
        ]

        if pc_rows.empty:
            error_msg = (
                f"{base_msg} | "
                "Expected at least one Alternative Unit of Measure 'PC' "
                "with Numerator = 1 and Denominator = 1"
            )

            resultados.append({
                'Line': '',
                'Field': 'Alternative UoM',
                'Error': error_msg,
                'Material': material,
                'Plant': ''
            })
            continue

        # Check if ANY PC row is valid (1 / 1)
        valid_pc_found = False
        num_invalid = False
        den_invalid = False

        def _auom_to_float(x):
            """
            Safely converts Numerator/Denominator values to float.
            Handles European decimal-comma exports (e.g. '1,000' -> 1.0)
            and blank/NaN values. Returns None if truly unparseable.
            """
            if x is None or (isinstance(x, float) and pd.isna(x)):
                return None
            s = str(x).strip()
            if s == '':
                return None
            s = s.replace(',', '.')
            try:
                return float(s)
            except (ValueError, TypeError):
                return None

        for _, r in pc_rows.iterrows():
            num = _auom_to_float(r.get('Numerator'))
            den = _auom_to_float(r.get('Denominator'))

            if (
                num is not None and den is not None
                and abs(num - 1) < 1e-6 and abs(den - 1) < 1e-6
            ):
                valid_pc_found = True
                break

            if num is None or abs(num - 1) >= 1e-6:
                num_invalid = True
            if den is None or abs(den - 1) >= 1e-6:
                den_invalid = True

        if valid_pc_found:
            continue

        # Build aggregated error message
        if num_invalid and den_invalid:
            reason = "Alternative Unit of Measure 'PC' exists but Numerator and Denominator must be 1"
        elif num_invalid:
            reason = "Alternative Unit of Measure 'PC' exists but Numerator must be 1"
        else:
            reason = "Alternative Unit of Measure 'PC' exists but Denominator must be 1"

        error_msg = f"{base_msg} | {reason}"

        resultados.append({
            'Line': '',
            'Field': 'Alternative UoM',
            'Error': error_msg,
            'Material': material,
            'Plant': ''
        })

def validar_lifo_fifo_relevant_mbew(mbew_df, mara_df, resultados):
    """
    SAP Material Validation:
    MBEW - LIFO/FIFO-relevant must be 'X' for FERT, HALB and ROH
    per (Material, Valuation Area = Plant).
    """

    if mbew_df is None or mbew_df.empty:
        return

    required_mbew = {'Material', 'Valuation Area', 'LIFO/FIFO-relevant'}
    required_mara = {'Material', 'Material Type'}

    if (not required_mbew.issubset(mbew_df.columns) or
        not required_mara.issubset(mara_df.columns)):
        return

    mbew = mbew_df.copy()
    mara = mara_df[['Material', 'Material Type']].copy()

    # Join Material Type
    df = mbew.merge(mara, on='Material', how='left')

    scope = df['Material Type'].isin(['FERT', 'HALB', 'ROH'])

    for idx, row in df[scope].iterrows():
        val = str(row.get('LIFO/FIFO-relevant', '')).strip().upper()
        plant = str(row.get('Valuation Area', '')).strip()

        if val != 'X':
            resultados.append({
                'Line': idx + 2,
                'Field': 'LIFO/FIFO-relevant',
                'Error': "Must be 'X' for material types FERT / HALB / ROH",
                'Material': str(row.get('Material', '')).strip(),
                'Plant': plant
            })

def validar_country_region_origin_fp(
    marc_df,
    zplm_marc_df,
    mara_df,
    resultados
):
    """
    MARC - Country/Region of Origin is mandatory for:
    - Plant Type = FP (from ZPLM MARC)
    - PSMS in {25,30,35,40,45,50,55,65}
    - Exception: FP with PSMS = 10
    - ONLY for Material Types: FERT, HALB, ROH
    """

    required_marc = {
        'Material',
        'Plant',
        'Plant Specific Material Status',
        'Cntry/Reg of Origin',
    }

    required_zplm = {
        'Material',
        'Plant',
        'Plant Type',
    }

    required_mara = {
        'Material',
        'Material Type',
    }

    if (
        marc_df is None or zplm_marc_df is None or mara_df is None
        or marc_df.empty or zplm_marc_df.empty or mara_df.empty
        or not required_marc.issubset(marc_df.columns)
        or not required_zplm.issubset(zplm_marc_df.columns)
        or not required_mara.issubset(mara_df.columns)
    ):
        return

    valid_status = {25, 30, 35, 40, 45, 50, 55, 65}

    df = marc_df.merge(
        zplm_marc_df[['Material', 'Plant', 'Plant Type']],
        on=['Material', 'Plant'],
        how='left'
    ).merge(
        mara_df[['Material', 'Material Type']],
        on='Material',
        how='left'
    )

    for idx, row in df.iterrows():

        mt = str(row.get('Material Type', '')).strip().upper()

        # Apenas FERT / HALB / ROH
        if mt not in {'FERT', 'HALB', 'ROH'}:
            continue

        plant_type = str(
            row.get('Plant Type', '')
        ).strip().upper()

        psms = normalize_code(
            row.get('Plant Specific Material Status', ''),
            as_int=True
        )

        origin = str(
            row.get('Cntry/Reg of Origin', '')
        ).strip()

        # Excluir FP com PSMS 10
        if psms == 10:
            continue

        if (
            plant_type == 'FP'
            and psms in valid_status
            and origin == ''
        ):
            resultados.append({
                'Line': idx + 2,
                'Field': 'Cntry/Reg of Origin',
                'Error': (
                    'Mandatory for FP materials with '
                    'PSMS 25/30/35/40/45/50/55/65'
                ),
                'Material': str(row.get('Material', '')).strip(),
                'Plant': str(row.get('Plant', '')).strip(),
            })

def validar_commodity_code_fp(
    marc_df,
    zplm_marc_df,
    mara_df,
    resultados
):
    """  
    MARC - Commodity Code is mandatory for:
    - Plant Type = FP
    - Except PSMS = 10
    - ONLY for Material Types: FERT, HALB, ROH
    """

    required_marc = {
        'Material',
        'Plant',
        'Plant Specific Material Status',
        'Commodity Code',
    }

    required_zplm = {
        'Material',
        'Plant',
        'Plant Type',
    }

    required_mara = {
        'Material',
        'Material Type',
    }

    if (
        marc_df is None or zplm_marc_df is None or mara_df is None
        or marc_df.empty or zplm_marc_df.empty or mara_df.empty
        or not required_marc.issubset(marc_df.columns)
        or not required_zplm.issubset(zplm_marc_df.columns)
        or not required_mara.issubset(mara_df.columns)
    ):
        return

    df = marc_df.merge(
        zplm_marc_df[['Material', 'Plant', 'Plant Type']],
        on=['Material', 'Plant'],
        how='left'
    ).merge(
        mara_df[['Material', 'Material Type']],
        on='Material',
        how='left'
    )

    for idx, row in df.iterrows():

        mt = str(row.get('Material Type', '')).strip().upper()

        if mt not in {'FERT', 'HALB', 'ROH'}:
            continue

        plant_type = str(row.get('Plant Type', '')).strip().upper()

        psms = normalize_code(
            row.get('Plant Specific Material Status', ''),
            as_int=True
        )

        commodity_code = str(
            row.get('Commodity Code', '')
        ).strip()

        if (
            plant_type == 'FP'
            and psms != 10
            and commodity_code == ''
        ):
            resultados.append({
                'Line': idx + 2,
                'Field': 'Commodity Code',
                'Error': (
                    'Mandatory for FP materials '
                    '(except PSMS 10)'
                ),
                'Material': str(row.get('Material', '')).strip(),
                'Plant': str(row.get('Plant', '')).strip(),
            })
            
def validar_stock_determination_group_marc(marc_df, mara_df, resultados):
    """
    MARC:
    Stock Determination Group must be '0001' for Material Types:
    - FERT
    - HALB
    - ROH
    Validation is per (Material, Plant).
    """
    required_marc = {"Material", "Plant", "Stock Determination Group"}
    required_mara = {"Material", "Material Type"}

    if (
        marc_df is None or mara_df is None
        or not required_marc.issubset(marc_df.columns)
        or not required_mara.issubset(mara_df.columns)
    ):
        return

    df = marc_df.merge(
        mara_df[["Material", "Material Type"]],
        on="Material",
        how="left"
    )

    for idx, row in df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant = str(row.get("Plant", "")).strip()
        mt = str(row.get("Material Type", "")).strip()

        # ✅ NORMALIZAÇÃO CORRETA
        sdg = normalize_code(
            row.get("Stock Determination Group", ""),
            as_int=True,
            zero_pad=4
        )

        if mt in {"FERT", "HALB", "ROH"}:
            if sdg == "":
                error = "Mandatory field blank (should be '0001')"
            elif sdg != "0001":
                error = f"Invalid value: '{sdg}' (should be '0001')"
            else:
                continue

            resultados.append({
                "Line": idx + 2,
                "Field": "Stock Determination Group",
                "Error": error,
                "Material": material,
                "Plant": plant,
            })

def validar_storage_location_ext_procurement(marc_df, zplm_marc_df, mara_df, resultados):
    """
    Validate MARC 'Storage Location for Ext. Procurement' for Palmela.

    Applies to Material Types FERT/HALB/ROH and Plant Types FP/MP/PP.
    Expected values are defined by Plant and Plant Type:
      PT01: FP -> P120, MP -> P115, PP -> P100
      PT02: FP -> P220, MP -> P215, PP -> P200

    The same Plant/Plant Type mapping applies to FERT, HALB and ROH.
    """
    col = 'Storage Location for Ext. Procurement'
    required_marc = {'Material', 'Plant', col}
    required_mara = {'Material', 'Material Type'}

    if (
        marc_df is None or mara_df is None
        or marc_df.empty or mara_df.empty
        or not required_marc.issubset(marc_df.columns)
        or not required_mara.issubset(mara_df.columns)
    ):
        return

    pt_map = _pt_map_from_zplm(zplm_marc_df)
    mt_map = _mt_map_from_mara(mara_df)

    marc = marc_df.copy()
    marc.columns = [str(c).strip() for c in marc.columns]
    marc['Material'] = marc['Material'].astype(str).str.strip()
    marc['Plant'] = marc['Plant'].astype(str).str.strip().str.upper()
    marc = marc.drop_duplicates(subset=['Material', 'Plant'], keep='first')

    expected_map = {
        ('PT01', 'FP'): 'P120',
        ('PT01', 'MP'): 'P115',
        ('PT01', 'PP'): 'P100',
        ('PT02', 'FP'): 'P220',
        ('PT02', 'MP'): 'P215',
        ('PT02', 'PP'): 'P200',
    }
    valid_material_types = {'FERT', 'HALB', 'ROH'}

    for idx, row in marc.iterrows():
        material = str(row.get('Material', '')).strip()
        plant = str(row.get('Plant', '')).strip().upper()
        plant_type = pt_map.get((material, plant), '')
        mt = mt_map.get(material, '')
        raw = row.get(col)
        value = '' if pd.isna(raw) else str(raw).strip().upper()

        if mt not in valid_material_types:
            continue

        expected = expected_map.get((plant, plant_type))
        if expected is None:
            continue

        if value == '':
            error = (
                f"Mandatory field blank (expected '{expected}' for {plant} / "
                f"Material Type '{mt}' / Plant Type '{plant_type}')"
            )
        elif value != expected:
            error = (
                f"Invalid value: '{value}' (expected '{expected}' for {plant} / "
                f"Material Type '{mt}' / Plant Type '{plant_type}')"
            )
        else:
            continue

        resultados.append({
            'Line': idx + 2,
            'Field': col,
            'Error': error,
            'Material': material,
            'Plant': plant,
        })

def validar_individual_collective_requirements_marc(marc_df, mara_df, resultados):
    """
    MARC:
    Individual/Collective Requirements must be '1'
    for Material Types:
    - FERT
    - HALB
    - ROH
    Validation is per (Material, Plant).
    """
    required_marc = {
        "Material",
        "Plant",
        "Individual/ Collective Requirements",
    }
    required_mara = {
        "Material",
        "Material Type",
    }

    if (
        marc_df is None or mara_df is None
        or not required_marc.issubset(marc_df.columns)
        or not required_mara.issubset(mara_df.columns)
    ):
        return

    # Join MARC + MARA to get Material Type
    df = marc_df.merge(
        mara_df[["Material", "Material Type"]],
        on="Material",
        how="left",
    )

    for idx, row in df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant = str(row.get("Plant", "")).strip()
        mt = str(row.get("Material Type", "")).strip()

        # ✅ Normalização correta do campo SAP numérico
        icr = normalize_code(
            row.get("Individual/ Collective Requirements", ""),
            as_int=True,
        )

        if mt in {"FERT", "HALB", "ROH"}:
            if icr == "":
                error = "Mandatory field blank (should be '1')"
            elif icr != "1":
                error = f"Invalid value: '{icr}' (should be '1')"
            else:
                continue  # válido

            resultados.append({
                "Line": idx + 2,
                "Field": "Individual/ Collective Requirements",
                "Error": error,
                "Material": material,
                "Plant": plant,
            })

def validar_wh_material_group_mara(mara_df, resultados):
    """
    MARA: WH Material Group must be EMPTY for Material Types FERT, HALB and ROH.
    Validation is material-level because the field belongs to MARA.
    """
    required_cols = {"Material", "Material Type", "WH Material Group"}
    if (
        mara_df is None
        or mara_df.empty
        or not required_cols.issubset(mara_df.columns)
    ):
        return

    valid_material_types = {"FERT", "HALB", "ROH"}
    mara = mara_df.drop_duplicates(subset=["Material"], keep="first")

    for idx, row in mara.iterrows():
        material = _norm_str(row.get("Material"))
        material_type = _norm_str(row.get("Material Type")).upper()
        wh_material_group = _norm_str(row.get("WH Material Group"))

        if material_type not in valid_material_types:
            continue

        if wh_material_group != "":
            resultados.append({
                "Line": idx + 2,
                "Field": "WH Material Group",
                "Error": (
                    "Field must be empty for Material Types FERT, HALB and ROH "
                    f"(current: '{wh_material_group}')"
                ),
                "Material": material,
                "Plant": "",
            })


def validar_packaging_material_type_mara(mara_df, resultados):
    """
    MARA:
    Packaging Mat. Type must be EMPTY for Material Types:
    - FERT
    - HALB
    - ROH
    """
    required_cols = {
        "Material",
        "Material Type",
        "Packaging Mat. Type",
    }

    if mara_df is None or not required_cols.issubset(mara_df.columns):
        return

    for idx, row in mara_df.iterrows():
        material = str(row.get("Material", "")).strip()
        mt = str(row.get("Material Type", "")).strip()

        # ✅ TRATAMENTO CORRETO DE VALORES VAZIOS
        raw_pkg = row.get("Packaging Mat. Type", "")
        pkg_type = "" if pd.isna(raw_pkg) else str(raw_pkg).strip()

        if mt in {"FERT", "HALB", "ROH"} and pkg_type != "":
            resultados.append({
                "Line": idx + 2,
                "Field": "Packaging Mat. Type",
                "Error": "Field must be empty for FERT, HALB and ROH materials",
                "Material": material,
                "Plant": "",  # MARA → sem plant
            })

def validar_production_supervisor_marc(marc_df, mara_df, resultados):
    """
    MARC:
    Production Supervisor must be 'Z01'
    for Material Types:
    - FERT
    - HALB
    Validation is per (Material, Plant).
    """
    required_marc = {
        "Material",
        "Plant",
        "Production Supervisor",
    }
    required_mara = {
        "Material",
        "Material Type",
    }

    if (
        marc_df is None or mara_df is None
        or not required_marc.issubset(marc_df.columns)
        or not required_mara.issubset(mara_df.columns)
    ):
        return

    # Join MARC + MARA to get Material Type
    df = marc_df.merge(
        mara_df[["Material", "Material Type"]],
        on="Material",
        how="left",
    )

    for idx, row in df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant = str(row.get("Plant", "")).strip()
        mt = str(row.get("Material Type", "")).strip()

        # ✅ normalização segura (NaN → '')
        psuper = _norm_str(row.get("Production Supervisor")).upper()

        if mt in {"FERT", "HALB"}:
            if psuper == "":
                error = "Mandatory field blank (should be 'Z01')"
            elif psuper != "Z01":
                error = f"Invalid value: '{psuper}' (should be 'Z01')"
            else:
                continue  # válido

            resultados.append({
                "Line": idx + 2,
                "Field": "Production Supervisor",
                "Error": error,
                "Material": material,
                "Plant": plant,
            })

def validar_batch_management_plant_marc(
    marc_df,
    zplm_marc_df,
    mara_df,
    resultados
):
    """
    Validate MARC field 'Batch Management(Plant)'.

    Scope:
      - Material Types FERT, HALB and ROH only.
      - Validation is performed per Material + Plant.

    Rules and precedence:

    1. Any material in TN02:
         Batch Management(Plant) must be empty.

    2. Outside TN02:
         - Plant Type PP or COI:
             Batch Management(Plant) must be active.
             Accepted values: '1' or 'X'.

         - Other Plant Types:
             Batch Management(Plant) must not be active.
             Therefore, it must not contain '1' or 'X'.

    3. Dependency rule outside TN02:
         If Batch Management(Plant) is active,
         then MARA Batch Management must also be active.

         Values '1' and 'X' are considered equivalent.

    Notes:
      - The TN02 rule has the highest precedence.
      - Therefore, TN02 materials with Plant Type PP or COI
        must still have Batch Management(Plant) empty.
      - The dependency rule is not applied to TN02 because
        Batch Management(Plant) must always be empty there.
    """

    marc_field = 'Batch Management(Plant)'
    mara_field = 'Batch Management'

    required_marc_columns = {
        'Material',
        'Plant',
        marc_field,
    }

    required_zplm_columns = {
        'Material',
        'Plant',
        'Plant Type',
    }

    required_mara_columns = {
        'Material',
        'Material Type',
        mara_field,
    }

    if (
        marc_df is None
        or zplm_marc_df is None
        or mara_df is None
        or marc_df.empty
        or zplm_marc_df.empty
        or mara_df.empty
    ):
        return

    if not required_marc_columns.issubset(marc_df.columns):
        missing_columns = sorted(
            required_marc_columns - set(marc_df.columns)
        )

        print(
            "Missing MARC columns for Batch Management(Plant) "
            f"validation: {missing_columns}"
        )
        return

    if not required_zplm_columns.issubset(zplm_marc_df.columns):
        missing_columns = sorted(
            required_zplm_columns - set(zplm_marc_df.columns)
        )

        print(
            "Missing ZPLM MARC columns for Batch Management(Plant) "
            f"validation: {missing_columns}"
        )
        return

    if not required_mara_columns.issubset(mara_df.columns):
        missing_columns = sorted(
            required_mara_columns - set(mara_df.columns)
        )

        print(
            "Missing MARA columns for Batch Management(Plant) "
            f"validation: {missing_columns}"
        )
        return

    # ------------------------------------------------------------
    # Create local copies
    # ------------------------------------------------------------
    marc = marc_df.copy()
    zplm_marc = zplm_marc_df.copy()
    mara = mara_df.copy()

    marc.columns = marc.columns.astype(str).str.strip()
    zplm_marc.columns = (
        zplm_marc.columns
        .astype(str)
        .str.strip()
    )
    mara.columns = mara.columns.astype(str).str.strip()

    # ------------------------------------------------------------
    # Normalize key columns
    # ------------------------------------------------------------
    marc['Material'] = (
        marc['Material']
        .fillna('')
        .astype(str)
        .str.strip()
    )

    marc['Plant'] = (
        marc['Plant']
        .fillna('')
        .astype(str)
        .str.strip()
        .str.upper()
    )

    zplm_marc['Material'] = (
        zplm_marc['Material']
        .fillna('')
        .astype(str)
        .str.strip()
    )

    zplm_marc['Plant'] = (
        zplm_marc['Plant']
        .fillna('')
        .astype(str)
        .str.strip()
        .str.upper()
    )

    mara['Material'] = (
        mara['Material']
        .fillna('')
        .astype(str)
        .str.strip()
    )

    # ------------------------------------------------------------
    # Avoid duplicate validations caused by duplicate source rows
    # ------------------------------------------------------------
    marc = marc.drop_duplicates(
        subset=['Material', 'Plant'],
        keep='first'
    )

    zplm_marc = zplm_marc.drop_duplicates(
        subset=['Material', 'Plant'],
        keep='first'
    )

    mara = mara.drop_duplicates(
        subset=['Material'],
        keep='first'
    )

    # ------------------------------------------------------------
    # Merge:
    #   - Plant Type from ZPLM MARC
    #   - Material Type and Batch Management from MARA
    # ------------------------------------------------------------
    df = (
        marc.merge(
            zplm_marc[
                [
                    'Material',
                    'Plant',
                    'Plant Type',
                ]
            ],
            on=['Material', 'Plant'],
            how='left',
        )
        .merge(
            mara[
                [
                    'Material',
                    'Material Type',
                    mara_field,
                ]
            ],
            on='Material',
            how='left',
        )
    )

    valid_material_types = {
        'FERT',
        'HALB',
        'ROH',
    }

    # ------------------------------------------------------------
    # Validate each Material + Plant
    # ------------------------------------------------------------
    for idx, row in df.iterrows():
        erros = []

        material = _norm_str(
            row.get('Material')
        )

        plant = _norm_str(
            row.get('Plant')
        ).upper()

        plant_type = _norm_str(
            row.get('Plant Type')
        ).upper()

        material_type = _norm_str(
            row.get('Material Type')
        ).upper()

        batch_management_plant = _norm_str(
            row.get(marc_field)
        ).upper()

        batch_management_mara = _norm_str(
            row.get(mara_field)
        ).upper()

        # Treat values read numerically from Excel as equivalent
        if batch_management_plant == '1.0':
            batch_management_plant = '1'

        if batch_management_mara == '1.0':
            batch_management_mara = '1'

        # X and 1 represent the same active indicator
        active_values = {'1', 'X'}

        # Only validate FERT, HALB and ROH
        if material_type not in valid_material_types:
            continue

        # ========================================================
        # RULE 1
        # Any material in TN02 must have
        # Batch Management(Plant) empty.
        #
        # This takes precedence over the Plant Type rules.
        # ========================================================
        if plant == 'TN02':
            if batch_management_plant != '':
                erros.append(
                    "Batch Management(Plant) must be empty "
                    "for all TN02 materials "
                    f"(current: '{batch_management_plant}')"
                )

        # ========================================================
        # RULE 2
        # Outside TN02, PP and COI must have an active value.
        #
        # Both 1 and X are accepted.
        # ========================================================
        elif plant_type in {'PP', 'COI'}:
            if batch_management_plant == '':
                erros.append(
                    "Mandatory field blank "
                    f"(Plant Type '{plant_type}' requires "
                    "Batch Management(Plant) = '1' or 'X')"
                )

            elif batch_management_plant not in active_values:
                erros.append(
                    f"For Plant Type '{plant_type}', "
                    "Batch Management(Plant) should be 'Yes' "
                    f"(Current: '{batch_management_plant}')"
                )

        # ========================================================
        # RULE 3
        # Outside TN02, other Plant Types must not have an
        # active value.
        #
        # Both 1 and X are considered active.
        # ========================================================
        else:
            if batch_management_plant in active_values:
                erros.append(
                    "Batch Management(Plant) should be 'No' "
                    f"for Plant Type '{plant_type}' "
                    f"(Current: '{batch_management_plant}')"
                )

        # ========================================================
        # DEPENDENCY RULE
        #
        # Outside TN02, if Batch Management(Plant) is active,
        # MARA Batch Management must also be active.
        #
        # Both 1 and X are accepted in both fields.
        # ========================================================
        if (
            plant != 'TN02'
            and batch_management_plant in active_values
            and batch_management_mara not in active_values
        ):
            erros.append(
                "Batch Management(Plant) cannot be '1' or 'X' "
                "when MARA Batch Management is not '1' or 'X' "
                f"(current MARA value: "
                f"'{batch_management_mara or 'blank'}')"
            )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': marc_field,
                'Error': ' | '.join(erros),
                'Material': material,
                'Plant': plant,
            })

def validar_version_indicator_marc(
    marc_df,
    zplm_marc_df,
    mara_df,
    resultados,
):
    """
    MARC:
    Version Indicator rules:
    - Plant Type FP or MP  → must be 'X'
    - Any other Plant Type → must be EMPTY
    - ONLY for Material Types: FERT, HALB, ROH
    - EXCLUDES: Planning Status (PSMS) ∈ {10, 70, 75}
    - EXCLUDES: Plant TN02
    Validation per (Material, Plant).
    """

    required_marc = {
        "Material",
        "Plant",
        "Version Indicator",
        "Plant Specific Material Status",
    }

    required_zplm = {
        "Material",
        "Plant",
        "Plant Type",
    }

    required_mara = {
        "Material",
        "Material Type",
    }

    if (
        marc_df is None
        or zplm_marc_df is None
        or mara_df is None
        or marc_df.empty
        or zplm_marc_df.empty
        or mara_df.empty
        or not required_marc.issubset(marc_df.columns)
        or not required_zplm.issubset(zplm_marc_df.columns)
        or not required_mara.issubset(mara_df.columns)
    ):
        return

    df = marc_df.merge(
        zplm_marc_df[["Material", "Plant", "Plant Type"]],
        on=["Material", "Plant"],
        how="left",
    ).merge(
        mara_df[["Material", "Material Type"]],
        on="Material",
        how="left",
    )

    for idx, row in df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant = str(row.get("Plant", "")).strip().upper()

        # ✅ Skip TN02
        if plant == "TN02":
            continue

        plant_type = str(row.get("Plant Type", "")).strip().upper()
        mt = str(row.get("Material Type", "")).strip().upper()

        if mt not in {"FERT", "HALB", "ROH"}:
            continue

        psms_raw = row.get("Plant Specific Material Status", "")

        try:
            psms = (
                int(float(str(psms_raw).strip()))
                if str(psms_raw).strip()
                else None
            )
        except (ValueError, TypeError):
            psms = None

        if psms in {10, 70, 75}:
            continue

        version_ind = _norm_str(
            row.get("Version Indicator")
        ).upper()

        erros = []

        if plant_type in {"FP", "MP"}:
            if version_ind != "X":
                erros.append(
                    "Mandatory for Plant Type FP/MP (should be 'X')"
                )
        else:
            if version_ind == "X":
                erros.append(
                    f"Not allowed for Plant Type '{plant_type}'"
                )

        if erros:
            resultados.append({
                "Line": idx + 2,
                "Field": "Version Indicator",
                "Error": " | ".join(erros),
                "Material": material,
                "Plant": plant,
            })

def validar_costing_lot_size(marc_df, mara_df, resultados):
    """
    MARC:
    Costing Lot Size must be 10000 for Material Types:
    - FERT
    - HALB
    - ROH
    Validation per (Material, Plant).
    """

    required_marc = {"Material", "Plant", "Costing Lot Size"}
    required_mara = {"Material", "Material Type"}

    if (
        marc_df is None
        or mara_df is None
        or marc_df.empty
        or mara_df.empty
        or not required_marc.issubset(marc_df.columns)
        or not required_mara.issubset(mara_df.columns)
    ):
        return

    # Map Material -> Material Type
    mt_map = dict(
        zip(
            mara_df["Material"].astype(str).str.strip(),
            mara_df["Material Type"].astype(str).str.strip().str.upper()
        )
    )

    def parse_val(x):
        if pd.isna(x) or str(x).strip() == '':
            return None
        try:
            return int(float(str(x).strip()))
        except Exception:
            return 'INVALID'

    for idx, row in marc_df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant    = str(row.get("Plant", "")).strip()
        mt       = mt_map.get(material, '')

        # ✅ Só FERT / HALB / ROH
        if mt not in {"FERT", "HALB", "ROH"}:
            continue

        raw_val = row.get("Costing Lot Size")
        val     = parse_val(raw_val)

        if val is None:
            resultados.append({
                "Line": idx + 2,
                "Field": "Costing Lot Size",
                "Error": "Mandatory field blank",
                "Material": material,
                "Plant": plant,
            })

        elif val == 'INVALID':
            resultados.append({
                "Line": idx + 2,
                "Field": "Costing Lot Size",
                "Error": f"Invalid value: '{raw_val}' (must be 10000)",
                "Material": material,
                "Plant": plant,
            })

        elif val != 10000:
            resultados.append({
                "Line": idx + 2,
                "Field": "Costing Lot Size",
                "Error": f"Invalid value: '{val}' (must be 10000)",
                "Material": material,
                "Plant": plant,
            })

def validar_safety_time_marc(marc_df, mara_df, resultados):
    """
    MARC:
    Safety Time rules:
    - Applies to Material Types: FERT, HALB, ROH
    - Value can be:
        - 0 or EMPTY → OK
        - Otherwise → must be 2
    Validation per (Material, Plant).
    """

    required_marc = {
        "Material",
        "Plant",
        "Safety Time",
    }
    required_mara = {
        "Material",
        "Material Type",
    }

    if (
        marc_df is None
        or mara_df is None
        or not required_marc.issubset(marc_df.columns)
        or not required_mara.issubset(mara_df.columns)
    ):
        return

    # ✅ Join para ir buscar Material Type
    df = marc_df.merge(
        mara_df[["Material", "Material Type"]],
        on="Material",
        how="left",
    )

    def parse_val(x):
        if pd.isna(x) or str(x).strip() == "":
            return None
        try:
            return int(float(str(x).strip()))
        except Exception:
            return "INVALID"

    for idx, row in df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant    = str(row.get("Plant", "")).strip()
        mt       = str(row.get("Material Type", "")).strip().upper()

        # ✅ Só FERT / HALB / ROH
        if mt not in {"FERT", "HALB", "ROH"}:
            continue

        raw_val = row.get("Safety Time")
        val     = parse_val(raw_val)

        # ✅ blank ou 0 → OK
        if val in (None, 0):
            continue

        # ✅ valor inválido
        if val == "INVALID":
            resultados.append({
                "Line": idx + 2,
                "Field": "Safety Time",
                "Error": f"Invalid value: '{raw_val}' (must be 0, empty, or 2)",
                "Material": material,
                "Plant": plant,
            })

        # ✅ diferente de 2 → erro
        elif val != 2:
            resultados.append({
                "Line": idx + 2,
                "Field": "Safety Time",
                "Error": f"Invalid value: '{val}' (must be 2 when not 0/empty)",
                "Material": material,
                "Plant": plant,
            })

def validar_material_price_determination_control(mbew_df, mara_df, resultados):
    """
    MBEW:
    Material Price Determination: Control must be '3' for:
    - FERT
    - HALB
    - ROH
    Validation per (Material, Valuation Area / Plant).
    """

    required_mbew = {
        "Material",
        "Valuation Area",
        "Material Price Determination: Control",
    }
    required_mara = {
        "Material",
        "Material Type",
    }

    if (
        mbew_df is None
        or mara_df is None
        or mbew_df.empty
        or mara_df.empty
        or not required_mbew.issubset(mbew_df.columns)
        or not required_mara.issubset(mara_df.columns)
    ):
        return

    # ✅ Join para obter Material Type
    df = mbew_df.merge(
        mara_df[["Material", "Material Type"]],
        on="Material",
        how="left",
    )

    def parse_val(x):
        if pd.isna(x) or str(x).strip() == "":
            return None
        try:
            return int(float(str(x).strip()))
        except Exception:
            return "INVALID"

    for idx, row in df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant    = str(row.get("Valuation Area", "")).strip()  # 👈 equivalente ao Plant
        mt       = str(row.get("Material Type", "")).strip().upper()

        # ✅ Só FERT / HALB / ROH
        if mt not in {"FERT", "HALB", "ROH"}:
            continue

        raw_val = row.get("Material Price Determination: Control")
        val     = parse_val(raw_val)

        if val is None:
            resultados.append({
                "Line": idx + 2,
                "Field": "Material Price Determination: Control",
                "Error": "Mandatory field blank (should be '3')",
                "Material": material,
                "Plant": plant,
            })

        elif val == "INVALID":
            resultados.append({
                "Line": idx + 2,
                "Field": "Material Price Determination: Control",
                "Error": f"Invalid value: '{raw_val}' (should be '3')",
                "Material": material,
                "Plant": plant,
            })

        elif val != 3:
            resultados.append({
                "Line": idx + 2,
                "Field": "Material Price Determination: Control",
                "Error": f"Invalid value: '{val}' (should be '3')",
                "Material": material,
                "Plant": plant,
            })

# =============================================================================
# C) DATA CLEANUP VALIDATIONS
# =============================================================================

def validar_fp_pt01_pt02_status55(zplm_marc_df, marc_df):
    """
    Materials with FP in both PT01 and PT02 cannot have PSMS = 55.
    """
    if zplm_marc_df is None or zplm_marc_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    zplm = zplm_marc_df.copy()
    zplm.columns = [c.strip().upper() for c in zplm.columns]
    marc = marc_df.copy()
    marc.columns = [c.strip().title() for c in marc.columns]

    df_fp = zplm[
        (zplm["MATERIAL PLANT TYPE"].astype(str).str.upper() == "FP") &
        (zplm["PLANT"].isin(["PT01", "PT02"]))
    ]
    fp_both = [mat for mat, plants in df_fp.groupby("MATERIAL")["PLANT"].agg(set).items()
               if {"PT01", "PT02"}.issubset(plants)]

    if not fp_both:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    marc_filtered = marc[marc["Material"].astype(str).str.strip().isin(fp_both)]
    erros = []
    for _, row in marc_filtered.iterrows():
        if str(row["Plant-Specific Material Status"]).strip() == "55":
            erros.append({"Material": str(row["Material"]).strip(),
                          "Plant": str(row["Plant"]).strip(),
                          "Error": "Material with FP in PT01 AND PT02 cannot have status 55"})
    return pd.DataFrame(erros)


def validar_fp55_eop_vs_service_liability(zplm_marc_df, marc_df):
    """
    For FP materials with PSMS = 55:
    (A) Service Liability Date must not be empty.
    (B) EOP Date must be strictly earlier than Service Liability Date.
    """
    empty_cols = ["Material", "Plant", "Material Plant Type", "PSMS", "EOP Date", "Service Liability Date", "Error"]
    if zplm_marc_df is None or zplm_marc_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=empty_cols)

    zplm = zplm_marc_df.copy()
    zplm.columns = [str(c).strip().upper() for c in zplm.columns]
    marc = marc_df.copy()
    marc.columns = [str(c).strip().upper() for c in marc.columns]

    def pick(df, candidates):
        return next((c for c in candidates if c in df.columns), None)

    z_mat    = pick(zplm, ["MATERIAL", "MATNR"])
    z_plant  = pick(zplm, ["PLANT", "WERKS"])
    z_mptype = pick(zplm, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])
    z_eop    = pick(zplm, ["EOP DATE", "EOP", "END OF PROD DATE", "END OF PRODUCTION DATE"])
    z_srv    = pick(zplm, ["SERVICE LIABILITY DATE", "SERVICE LIAB DATE", "SRV LIABILITY DATE"])
    m_mat    = pick(marc, ["MATERIAL", "MATNR"])
    m_plant  = pick(marc, ["PLANT", "WERKS"])
    m_psms   = pick(marc, ["PLANT-SPECIFIC MATERIAL STATUS", "PLANT SPECIFIC MATERIAL STATUS", "MTSTB"])

    if any(c is None for c in [z_mat, z_plant, z_mptype, m_mat, m_plant, m_psms]):
        return pd.DataFrame(columns=empty_cols)

    z = zplm.copy()
    z[z_mat] = z[z_mat].astype(str).str.strip()
    z[z_plant] = z[z_plant].astype(str).str.strip().str.upper()
    z[z_mptype] = z[z_mptype].astype(str).str.strip().str.upper()
    z_fp = z[z[z_mptype] == "FP"].copy()

    m = marc.copy()
    m[m_mat] = m[m_mat].astype(str).str.strip()
    m[m_plant] = m[m_plant].astype(str).str.strip().str.upper()
    m[m_psms] = m[m_psms].apply(
        lambda x: '' if pd.isna(x) else (str(int(float(str(x).strip()))) if str(x).strip().replace('.', '').isdigit() else str(x).strip())
    )
    m55 = m[m[m_psms] == "55"].copy()

    key_int = set(zip(z_fp[z_mat], z_fp[z_plant])) & set(zip(m55[m_mat], m55[m_plant]))
    if not key_int:
        return pd.DataFrame(columns=empty_cols)

    z_fp["_KEY"] = list(zip(z_fp[z_mat], z_fp[z_plant]))
    df_target = z_fp[z_fp["_KEY"].isin(key_int)].copy()

    def _dt(val):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return pd.NaT
        s = str(val).strip()
        if s == '' or s.lower() in ("nan", "none"):
            return pd.NaT
        try:
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                return pd.to_datetime("1899-12-30") + pd.to_timedelta(int(val), unit="D")
        except Exception:
            pass
        return pd.to_datetime(s, errors="coerce", dayfirst=True)

    errors = []
    for _, row in df_target.iterrows():
        material = row[z_mat]
        plant    = row[z_plant]
        mptype   = row[z_mptype]
        srv_raw  = row.get(z_srv, None)
        eop_raw  = row.get(z_eop, None)
        srv = _dt(srv_raw)
        eop = _dt(eop_raw)

        if pd.isna(srv):
            errors.append({"Material": material, "Plant": plant, "Material Plant Type": mptype,
                           "PSMS": "55", "EOP Date": eop_raw, "Service Liability Date": srv_raw,
                           "Error": "Service Liability Date empty for FP with status 55"})
        elif not pd.isna(eop) and eop >= srv:
            errors.append({"Material": material, "Plant": plant, "Material Plant Type": mptype,
                           "PSMS": "55", "EOP Date": eop_raw, "Service Liability Date": srv_raw,
                           "Error": "EOP Date must be strictly earlier than Service Liability Date"})

    return pd.DataFrame(errors, columns=empty_cols)


def validar_fp55_service_liability_expired(zplm_marc_df, marc_df):
    """
    For FP materials with PSMS = 55:
    Service Liability Date must not be <= today.
    """
    empty_cols = ["Material", "Plant", "Material Plant Type", "PSMS", "Service Liability Date", "Error"]
    if zplm_marc_df is None or zplm_marc_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=empty_cols)

    zplm = zplm_marc_df.copy()
    zplm.columns = [str(c).strip().upper() for c in zplm.columns]
    marc = marc_df.copy()
    marc.columns = [str(c).strip().upper() for c in marc.columns]

    def pick(df, candidates):
        return next((c for c in candidates if c in df.columns), None)

    z_mat    = pick(zplm, ["MATERIAL", "MATNR"])
    z_plant  = pick(zplm, ["PLANT", "WERKS"])
    z_mptype = pick(zplm, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])
    z_srv    = pick(zplm, ["SERVICE LIABILITY DATE", "SERVICE LIAB DATE", "SRV LIABILITY DATE"])
    m_mat    = pick(marc, ["MATERIAL", "MATNR"])
    m_plant  = pick(marc, ["PLANT", "WERKS"])
    m_psms   = pick(marc, ["PLANT-SPECIFIC MATERIAL STATUS", "PLANT SPECIFIC MATERIAL STATUS", "MTSTB"])

    if any(c is None for c in [z_mat, z_plant, z_mptype, z_srv, m_mat, m_plant, m_psms]):
        return pd.DataFrame(columns=empty_cols)

    z = zplm.copy()
    z[z_mat] = z[z_mat].astype(str).str.strip()
    z[z_plant] = z[z_plant].astype(str).str.strip().str.upper()
    z[z_mptype] = z[z_mptype].astype(str).str.strip().str.upper()
    z_fp = z[z[z_mptype] == "FP"].copy()

    m = marc.copy()
    m[m_mat] = m[m_mat].astype(str).str.strip()
    m[m_plant] = m[m_plant].astype(str).str.strip().str.upper()
    m[m_psms] = m[m_psms].apply(
        lambda x: '' if pd.isna(x) else (str(int(float(str(x).strip()))) if str(x).strip().replace('.', '').isdigit() else str(x).strip())
    )
    m55 = m[m[m_psms] == "55"]

    key_int = set(zip(z_fp[z_mat], z_fp[z_plant])) & set(zip(m55[m_mat], m55[m_plant]))
    if not key_int:
        return pd.DataFrame(columns=empty_cols)

    z_fp["_KEY"] = list(zip(z_fp[z_mat], z_fp[z_plant]))
    df_target = z_fp[z_fp["_KEY"].isin(key_int)].copy()
    today = pd.Timestamp.today().normalize()

    def _dt(val):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return pd.NaT
        s = str(val).strip()
        if s == '' or s.lower() in ("nan", "none"):
            return pd.NaT
        try:
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                return pd.to_datetime("1899-12-30") + pd.to_timedelta(int(val), unit="D")
        except Exception:
            pass
        return pd.to_datetime(s, errors="coerce", dayfirst=True)

    errors = []
    for _, row in df_target.iterrows():
        srv_raw = row[z_srv]
        srv = _dt(srv_raw)
        if not pd.isna(srv) and srv <= today:
            errors.append({"Material": row[z_mat], "Plant": row[z_plant],
                           "Material Plant Type": row[z_mptype], "PSMS": "55",
                           "Service Liability Date": srv_raw,
                           "Error": "Service Liability Date is already expired (≤ today)"})

    return pd.DataFrame(errors, columns=empty_cols)


def validar_fp50_eop_expired(zplm_marc_df, marc_df):
    """
    FP materials (Material Plant Type = FP) with PSMS in {10,20,30,40,45,50}

    Rules:
    - All plants:
        EOP Date must not be empty
    - PT01 and all plants except PT02:
        EOP Date must be strictly in the future (EOP > today)
    - PT02 only:
        EOP Date must NOT be in the future (EOP <= today)

    Status Groups:
    - Group A: 10, 20, 30
    - Group B: 40, 45, 50
    """

    empty_cols = [
        "Material", "Plant", "Material Plant Type",
        "PSMS", "Status Group", "EOP Date", "Error"
    ]

    if zplm_marc_df is None or zplm_marc_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=empty_cols)

    # -------------------------
    # Setup & column detection
    # -------------------------
    zplm = zplm_marc_df.copy()
    zplm.columns = [str(c).strip().upper() for c in zplm.columns]

    marc = marc_df.copy()
    marc.columns = [str(c).strip().upper() for c in marc.columns]

    def pick(df, candidates):
        return next((c for c in candidates if c in df.columns), None)

    z_mat    = pick(zplm, ["MATERIAL", "MATNR"])
    z_plant  = pick(zplm, ["PLANT", "WERKS"])
    z_mptype = pick(zplm, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])
    z_eop    = pick(zplm, ["EOP DATE", "EOP", "END OF PROD DATE", "END OF PRODUCTION DATE"])

    m_mat    = pick(marc, ["MATERIAL", "MATNR"])
    m_plant  = pick(marc, ["PLANT", "WERKS"])
    m_psms   = pick(
        marc,
        ["PLANT-SPECIFIC MATERIAL STATUS",
         "PLANT SPECIFIC MATERIAL STATUS",
         "MTSTB"]
    )

    if any(c is None for c in [z_mat, z_plant, z_mptype, z_eop, m_mat, m_plant, m_psms]):
        return pd.DataFrame(columns=empty_cols)

    # -------------------------
    # Normalize ZPLM
    # -------------------------
    z = zplm.copy()
    z[z_mat]    = z[z_mat].astype(str).str.strip()
    z[z_plant]  = z[z_plant].astype(str).str.strip().str.upper()
    z[z_mptype] = z[z_mptype].astype(str).str.strip().str.upper()

    z_fp = z[z[z_mptype] == "FP"].copy()
    if z_fp.empty:
        return pd.DataFrame(columns=empty_cols)

    # -------------------------
    # Normalize MARC
    # -------------------------
    m = marc.copy()
    m[m_mat]   = m[m_mat].astype(str).str.strip()
    m[m_plant] = m[m_plant].astype(str).str.strip().str.upper()
    m[m_psms]  = m[m_psms].apply(
        lambda x: '' if pd.isna(x)
        else (str(int(float(str(x).strip())))
              if str(x).strip().replace('.', '').isdigit()
              else str(x).strip())
    )

    GROUP_A = {"10", "20", "30"}
    GROUP_B = {"40", "45", "50"}
    ALL_PS  = GROUP_A | GROUP_B

    m_sel = m[m[m_psms].isin(ALL_PS)].copy()
    if m_sel.empty:
        return pd.DataFrame(columns=empty_cols)

    # -------------------------
    # Intersect keys
    # -------------------------
    key_int = set(zip(z_fp[z_mat], z_fp[z_plant])) & set(zip(m_sel[m_mat], m_sel[m_plant]))
    if not key_int:
        return pd.DataFrame(columns=empty_cols)

    z_fp["_KEY"] = list(zip(z_fp[z_mat], z_fp[z_plant]))
    df_target = z_fp[z_fp["_KEY"].isin(key_int)].copy()

    today = pd.Timestamp.today().normalize()

    def _dt(val):
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return pd.NaT
        s = str(val).strip()
        if s == '' or s.lower() in ("nan", "none"):
            return pd.NaT
        try:
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                return pd.to_datetime("1899-12-30") + pd.to_timedelta(int(val), unit="D")
        except Exception:
            pass
        return pd.to_datetime(s, errors="coerce", dayfirst=True)

    # -------------------------
    # Validation
    # -------------------------
    errors = []

    for _, row in df_target.iterrows():
        psms = m_sel.loc[
            (m_sel[m_mat] == row[z_mat]) &
            (m_sel[m_plant] == row[z_plant]),
            m_psms
        ].iloc[0]

        status_group = (
            "Group A (10/20/30)" if psms in GROUP_A
            else "Group B (40/45/50)"
        )

        plant = row[z_plant]
        eop_raw = row.get(z_eop, None)
        eop = _dt(eop_raw)

        # ---- Empty mantém para todos ----
        if pd.isna(eop):
            errors.append({
                "Material": row[z_mat],
                "Plant": plant,
                "Material Plant Type": row[z_mptype],
                "PSMS": psms,
                "Status Group": status_group,
                "EOP Date": eop_raw,
                "Error": "EOP Date is empty for FP material with active/phase-out status",
            })
            continue

        # ---- PT02: futuro proibido ----
        if plant == "PT02":
            if eop > today:
                errors.append({
                    "Material": row[z_mat],
                    "Plant": plant,
                    "Material Plant Type": row[z_mptype],
                    "PSMS": psms,
                    "Status Group": status_group,
                    "EOP Date": eop_raw,
                    "Error": (
                        f"EOP Date is in the future (EOP {eop.date()} > today {today.date()}) "
                        f"— PT02 does not allow future dates"
                    ),
                })

        # ---- PT01 e restantes: regra original (expired) ----
        else:
            if eop <= today:
                errors.append({
                    "Material": row[z_mat],
                    "Plant": plant,
                    "Material Plant Type": row[z_mptype],
                    "PSMS": psms,
                    "Status Group": status_group,
                    "EOP Date": eop_raw,
                    "Error": (
                        f"EOP Date is expired (EOP {eop.date()} ≤ today {today.date()}) "
                        f"— FP material must have EOP Date in the future"
                    ),
                })

    return pd.DataFrame(errors, columns=empty_cols)


def validar_fppp70_standard_price_mbew(zplm_marc_df, marc_df, mbew_df):
    """
    For FP/MP/PP materials with PSMS = 70:
    Standard Price (MBEW) must not be empty or zero.
    """
    empty_cols = ["Material", "Plant", "Material Plant Type", "PSMS", "Standard Price", "Error"]
    if any(df is None or df.empty for df in [zplm_marc_df, marc_df, mbew_df]):
        return pd.DataFrame(columns=empty_cols)

    zplm = zplm_marc_df.copy(); zplm.columns = [str(c).strip().upper() for c in zplm.columns]
    marc = marc_df.copy();      marc.columns = [str(c).strip().upper() for c in marc.columns]
    mbew = mbew_df.copy();      mbew.columns = [str(c).strip().upper() for c in mbew.columns]

    def pick(df, candidates):
        return next((c for c in candidates if c in df.columns), None)

    z_mat    = pick(zplm, ["MATERIAL", "MATNR"])
    z_plant  = pick(zplm, ["PLANT", "WERKS"])
    z_mptype = pick(zplm, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])
    m_mat    = pick(marc, ["MATERIAL", "MATNR"])
    m_plant  = pick(marc, ["PLANT", "WERKS"])
    m_psms   = pick(marc, ["PLANT-SPECIFIC MATERIAL STATUS", "MTSTB"])
    mb_mat   = pick(mbew, ["MATERIAL", "MATNR"])
    mb_plant = pick(mbew, ["VALUATION AREA", "BWKEY", "PLANT", "WERKS"])
    mb_price = pick(mbew, ["STANDARD PRICE", "STPRS", "STD PRICE"])

    if not all([z_mat, z_plant, z_mptype, m_mat, m_plant, m_psms, mb_mat, mb_plant, mb_price]):
        return pd.DataFrame(columns=empty_cols)

    z = zplm[[z_mat, z_plant, z_mptype]].copy()
    z[z_mat] = z[z_mat].astype(str).str.strip()
    z[z_plant] = z[z_plant].astype(str).str.strip().str.upper()
    z[z_mptype] = z[z_mptype].astype(str).str.strip().str.upper()
    z_sel = z[z[z_mptype].isin(["FP", "MP", "PP"])].copy()

    m = marc[[m_mat, m_plant, m_psms]].copy()
    m[m_mat] = m[m_mat].astype(str).str.strip()
    m[m_plant] = m[m_plant].astype(str).str.strip().str.upper()
    m[m_psms] = m[m_psms].apply(
        lambda x: '' if pd.isna(x) else (str(int(float(str(x).strip()))) if str(x).strip().replace('.','').isdigit() else '')
    )
    m70 = m[m[m_psms] == "70"]

    mb = mbew[[mb_mat, mb_plant, mb_price]].copy()
    mb[mb_mat] = mb[mb_mat].astype(str).str.strip()
    mb[mb_plant] = mb[mb_plant].astype(str).str.strip().str.upper()

    key_ok = set(zip(z_sel[z_mat], z_sel[z_plant])) & set(zip(m70[m_mat], m70[m_plant]))
    if not key_ok:
        return pd.DataFrame(columns=empty_cols)

    z_sel["_KEY"] = list(zip(z_sel[z_mat], z_sel[z_plant]))
    z_final = z_sel[z_sel["_KEY"].isin(key_ok)].drop(columns=["_KEY"])

    merged = z_final.merge(mb, left_on=[z_mat, z_plant], right_on=[mb_mat, mb_plant], how="left")
    merged = merged.merge(m70[[m_mat, m_plant, m_psms]], left_on=[z_mat, z_plant], right_on=[m_mat, m_plant], how="left")

    errors = []
    for _, row in merged.iterrows():
        price = row[mb_price]
        if price in [None, "", " ", 0, 0.0] or (isinstance(price, float) and np.isnan(price)):
            errors.append({"Material": row[z_mat], "Plant": row[z_plant],
                           "Material Plant Type": row[z_mptype], "PSMS": row[m_psms],
                           "Standard Price": price,
                           "Error": "Standard Price is empty or zero for FP/MP/PP with status 70"})

    return pd.DataFrame(errors, columns=empty_cols)


def validar_ltb_fields(zplm_marc, marc, mara):
    """
    LTB fields validation (Data Cleanups).

    Only materials meeting all the following conditions may have LTB values:
        - Material Plant Type = FP
        - Material Type in {FERT, HALB, ROH}
        - Plant-Specific Material Status in {55, 70, 99}

    Any other material with at least one populated LTB field generates
    a Type A error.

    LTB fields validated from ZPLM MARC:
        - LTB Qty - Quarter Calculation
        - LTB Original Qty Date
        - LTB Open Quantity
        - LTB Open Qty Update
        - LTB Original Qty
    """

    z = zplm_marc.copy()
    m = marc.copy()
    a = mara.copy()

    # ------------------------------------------------------------------
    # LTB fields from ZPLM MARC
    # ------------------------------------------------------------------
    ltb_fields = [
        'LTB Qty - Quarter Calculation',
        'LTB Original Qty Date',
        'LTB Open Quantity',
        'LTB Open Qty Update',
        'LTB Original Qty'
    ]

    ltb_quantity_fields = [
        'LTB Qty - Quarter Calculation',
        'LTB Open Quantity',
        'LTB Original Qty'
    ]

    ltb_date_fields = [
        'LTB Original Qty Date',
        'LTB Open Qty Update'
    ]

    # ------------------------------------------------------------------
    # Required columns
    # ------------------------------------------------------------------
    req_z = {
        'Material',
        'Plant',
        'Material Plant Type',
        *ltb_fields
    }

    req_m = {
        'Material',
        'Plant',
        'Plant-Specific Material Status'
    }

    req_a = {
        'Material',
        'Material Type'
    }

    missing = []

    if not req_z.issubset(z.columns):
        missing.append(
            f"ZPLM_MARC: {sorted(req_z - set(z.columns))}"
        )

    if not req_m.issubset(m.columns):
        missing.append(
            f"MARC: {sorted(req_m - set(m.columns))}"
        )

    if not req_a.issubset(a.columns):
        missing.append(
            f"MARA: {sorted(req_a - set(a.columns))}"
        )

    # ------------------------------------------------------------------
    # Output columns
    # ------------------------------------------------------------------
    cols = [
        'Material',
        'Plant',
        'Material Plant Type',
        'Material Type',
        'Plant-Specific Material Status',
        *ltb_fields,
        'Field',
        'Error'
    ]

    if missing:
        return pd.DataFrame(columns=cols)

    # ------------------------------------------------------------------
    # Prepare MARA
    # ------------------------------------------------------------------
    a_use = a[
        [
            'Material',
            'Material Type'
        ]
    ].copy()

    a_use['Material'] = (
        a_use['Material']
        .astype(str)
        .str.strip()
    )

    a_use['Material Type'] = (
        a_use['Material Type']
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ------------------------------------------------------------------
    # Prepare MARC
    # ------------------------------------------------------------------
    m_use = m[
        [
            'Material',
            'Plant',
            'Plant-Specific Material Status'
        ]
    ].copy()

    m_use['Material'] = (
        m_use['Material']
        .astype(str)
        .str.strip()
    )

    m_use['Plant'] = (
        m_use['Plant']
        .astype(str)
        .str.strip()
    )

    status_num = pd.to_numeric(
        m_use['Plant-Specific Material Status'],
        errors='coerce'
    )

    m_use['Plant-Specific Material Status'] = (
        status_num
        .astype('Int64')
        .astype(str)
        .str.replace('<NA>', '', regex=False)
        .str.strip()
    )

    # ------------------------------------------------------------------
    # Prepare ZPLM MARC
    # ------------------------------------------------------------------
    z['Material'] = (
        z['Material']
        .astype(str)
        .str.strip()
    )

    z['Plant'] = (
        z['Plant']
        .astype(str)
        .str.strip()
    )

    z['Material Plant Type'] = (
        z['Material Plant Type']
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ------------------------------------------------------------------
    # Merge reference data
    # ------------------------------------------------------------------
    z = (
        z
        .merge(
            a_use,
            on='Material',
            how='left'
        )
        .merge(
            m_use,
            on=['Material', 'Plant'],
            how='left'
        )
    )

    # ------------------------------------------------------------------
    # Authorization rules
    # ------------------------------------------------------------------
    valid_mtypes = {'FERT', 'HALB', 'ROH'}
    valid_status = {'55', '70', '99'}

    is_fp = (
        z['Material Plant Type']
        .astype(str)
        .str.strip()
        .str.upper()
        .eq('FP')
    )

    is_valid_mt = (
        z['Material Type']
        .astype(str)
        .str.strip()
        .str.upper()
        .isin(valid_mtypes)
    )

    is_valid_st = (
        z['Plant-Specific Material Status']
        .astype(str)
        .str.strip()
        .isin(valid_status)
    )

    authorized = (
        is_fp
        & is_valid_mt
        & is_valid_st
    )

    # ------------------------------------------------------------------
    # Check whether each LTB field contains a value
    # ------------------------------------------------------------------
    field_has_value = pd.DataFrame(
        index=z.index
    )

    # Quantity fields:
    # NaN, blank and zero are considered empty.
    for field in ltb_quantity_fields:
        raw_value = z[field]

        cleaned_value = (
            raw_value
            .astype(str)
            .str.strip()
        )

        numeric_value = pd.to_numeric(
            raw_value,
            errors='coerce'
        )

        is_empty = (
            raw_value.isna()
            | cleaned_value.isin(
                [
                    '',
                    '0',
                    '0.0',
                    'nan',
                    'None',
                    '<NA>'
                ]
            )
            | numeric_value.fillna(0).eq(0)
        )

        field_has_value[field] = ~is_empty

    # Date fields:
    # NaN, blank, zero and invalid dates are considered empty.
    for field in ltb_date_fields:
        raw_value = z[field]

        cleaned_value = (
            raw_value
            .astype(str)
            .str.strip()
        )

        date_value = pd.to_datetime(
            raw_value,
            errors='coerce'
        )

        is_empty = (
            raw_value.isna()
            | cleaned_value.isin(
                [
                    '',
                    '0',
                    '0.0',
                    'nan',
                    'None',
                    '<NA>'
                ]
            )
            | date_value.isna()
        )

        field_has_value[field] = ~is_empty

    # Material has an LTB value if at least one of the five fields
    # contains a valid non-empty value.
    has_any_ltb_value = field_has_value.any(axis=1)

    # ------------------------------------------------------------------
    # Type A error
    # ------------------------------------------------------------------
    mask_A = (
        ~authorized
        & has_any_ltb_value
    )

    output_fields = [
        'Material',
        'Plant',
        'Material Plant Type',
        'Material Type',
        'Plant-Specific Material Status',
        *ltb_fields
    ]

    df_A = z.loc[
        mask_A,
        output_fields
    ].copy()

    if df_A.empty:
        return pd.DataFrame(columns=cols)

    # ------------------------------------------------------------------
    # Identify exactly which LTB fields contain values
    # ------------------------------------------------------------------
    df_A['Field'] = field_has_value.loc[
        mask_A,
        ltb_fields
    ].apply(
        lambda row: ' / '.join(
            field
            for field, has_value in row.items()
            if has_value
        ),
        axis=1
    )

    df_A['Error'] = (
        "Material not authorized to have LTB values "
        "(requires FP + MT∈{FERT/HALB/ROH} "
        "+ Status∈{55,70,99})"
    )

    # ------------------------------------------------------------------
    # Final organization
    # ------------------------------------------------------------------
    df_A = df_A[cols]

    df_A.sort_values(
        by=['Plant', 'Material'],
        inplace=True,
        kind='stable'
    )

    df_A.reset_index(
        drop=True,
        inplace=True
    )

    return df_A


def validar_ekpo_delivery_completion(marc, mara, zplm_marc, ekpo):
    """
    EKPO validation (Cleanups): For materials with Status∈{70,99},
    Plant Type∈{PP,MP,FP}, MT∈{ROH,FERT,HALB}:
    EKPO lines without Deletion Indicator must have Delivery Completed filled.
    """
    marc = marc.copy(); mara = mara.copy(); zpm = zplm_marc.copy(); ekpo = ekpo.copy()

    for df in [marc, mara, zpm, ekpo]:
        df["Material"] = df["Material"].astype(str).str.strip()

    marc["Plant"] = marc["Plant"].astype(str).str.strip()
    zpm["Plant"]  = zpm["Plant"].astype(str).str.strip()

    marc["Plant-Specific Material Status"] = (
        pd.to_numeric(marc["Plant-Specific Material Status"], errors="coerce")
        .astype("Int64").astype(str).str.replace("<NA>", "")
    )
    mara["Material Type"] = mara["Material Type"].astype(str).str.strip().str.upper()
    zpm["Material Plant Type"] = zpm["Material Plant Type"].astype(str).str.strip().str.upper()

    base = (
        marc.merge(zpm[["Material", "Plant", "Material Plant Type"]], on=["Material", "Plant"], how="left")
            .merge(mara[["Material", "Material Type"]], on="Material", how="left")
    )
    base = base[
        base["Plant-Specific Material Status"].isin({"70", "99"}) &
        base["Material Plant Type"].isin({"PP", "MP", "FP"}) &
        base["Material Type"].isin({"ROH", "FERT", "HALB"})
    ]

    empty_out = pd.DataFrame(columns=["Material", "Plant", "Material Plant Type", "Material Type",
                                       "Status", "Purchasing Document", "Deletion Indicator",
                                       "Delivery Completed", "Error", "Rule"])
    if base.empty:
        return empty_out

    ekpo["Deletion Indicator"] = ekpo["Deletion Indicator"].astype(str).str.strip()
    ekpo["Delivery Completed"] = ekpo["Delivery Completed"].astype(str).str.strip()

    merged = base.merge(ekpo, on="Material", how="left", suffixes=("", "_EKPO"))
    merged = merged[merged["Deletion Indicator"].isin(["", " ", None, pd.NA])]

    if merged.empty:
        return empty_out

    merged["Delivery Completed"] = merged["Delivery Completed"].astype(str).str.strip()
    errors = merged.loc[merged["Delivery Completed"].isin(["", " ", None, pd.NA]),
                        ["Material", "Plant", "Material Plant Type", "Material Type",
                         "Plant-Specific Material Status", "Purchasing Document",
                         "Deletion Indicator", "Delivery Completed"]].copy()

    if errors.empty:
        return empty_out

    errors = errors.rename(columns={"Plant-Specific Material Status": "Status"})
    errors["Error"] = "Open PO item — Delivery Completed is blank but should be closed"
    errors["Rule"]  = "EKPO Delivery Completion"
    return errors


def validar_mard_total_stock(mard, marc, mara):
    """
    Data Cleanups: For MT∈{FERT,HALB,ROH} with PSMS=70,
    error if no stock exists in any storage location.
    """
    import re as _re

    def _norm(s):
        return _re.sub(r'[^a-z0-9]+', '', str(s).lower())

    def _pick_first_cols(df, wanted_labels):
        out = {k: None for k in wanted_labels}
        norm_cols = [_norm(c) for c in df.columns]
        seen = set()
        for key, patterns in wanted_labels.items():
            for i, col in enumerate(df.columns):
                if col in seen:
                    continue
                for pat in patterns:
                    if norm_cols[i].startswith(_norm(pat)):
                        out[key] = col
                        seen.add(col)
                        break
                if out[key] is not None:
                    break
        return out

    mard = mard.copy(); marc = marc.copy(); mara = mara.copy()
    for df in (mard, marc, mara):
        if 'Material' in df.columns:
            df['Material'] = df['Material'].astype(str).str.strip()
    for df in (mard, marc):
        if 'Plant' in df.columns:
            df['Plant'] = df['Plant'].astype(str).str.strip()

    mara["Material Type"] = mara["Material Type"].astype(str).str.upper().str.strip()
    st_num = pd.to_numeric(marc["Plant-Specific Material Status"], errors="coerce")
    marc["Plant-Specific Material Status"] = st_num.astype("Int64").astype(str).str.replace("<NA>", "", regex=False).str.strip()

    base = (
        mard[["Material", "Plant"]].drop_duplicates()
            .merge(mara[["Material", "Material Type"]], on="Material", how="left")
            .merge(marc[["Material", "Plant", "Plant-Specific Material Status"]], on=["Material", "Plant"], how="left")
    )
    base = base[base["Material Type"].isin({"FERT", "HALB", "ROH"}) &
                (base["Plant-Specific Material Status"] == "70")]

    if base.empty:
        return pd.DataFrame(columns=["Material", "Plant", "Material Type", "Status", "Error", "Rule"])

    wanted = {"unrestricted": ["Unrestricted"], "stock_transfer": ["Stock in transfer"],
               "quality": ["Quality Inspection"], "blocked": ["Blocked"], "returns": ["Returns"]}
    colmap = _pick_first_cols(mard, wanted)
    stock_cols = [c for c in colmap.values() if c is not None]

    if not stock_cols:
        return pd.DataFrame(columns=["Material", "Plant", "Material Type", "Status", "Error", "Rule"])

    mard_use = mard[["Material", "Plant"] + stock_cols].copy()
    for c in stock_cols:
        mard_use[c] = pd.to_numeric(mard_use[c], errors="coerce").fillna(0)

    row_has_stock = None
    for c in stock_cols:
        row_has_stock = (mard_use[c] > 0) if row_has_stock is None else (row_has_stock | (mard_use[c] > 0))
    mard_use["_row_has_stock"] = row_has_stock.fillna(False)

    agg = mard_use.groupby(["Material", "Plant"])["_row_has_stock"].max().reset_index().rename(columns={"_row_has_stock": "Has_Stock"})
    out = agg.merge(base[["Material", "Plant", "Material Type", "Plant-Specific Material Status"]],
                    on=["Material", "Plant"], how="right").rename(columns={"Plant-Specific Material Status": "Status"})

    df_errors = out[out["Has_Stock"] == False].copy()
    df_errors["Error"] = "Material has no stock in any Storage Location (should have stock)"
    df_errors["Rule"]  = "MARD Total Stock"
    return df_errors.sort_values(["Plant", "Material"])


def validar_fp_existence_zplm_ta_ifs(zplm_marc_df, zplm_ta_ifs_df, marc_df):
    """Cleanup only: report active FP Material+Plant pairs missing in ZPLM_TA_IFS.

    ZPLM_TA_IFS is treated as a material-level interface table. The output remains
    plant-specific because FP classification and Part Status are plant-specific.
    PSMS 75/99 rows are expected to have been removed by the cleanup orchestrator.

    FP Material+Plant pairs marked for transfer in ZPLM MARC (field
    'Material Transfer' = X, see build_material_transfer_map) are skipped
    and not reported by this rule.
    """
    cols = ["Material", "Plant", "Material Plant Type", "Part Status", "Error", "Rule"]
    if any(df is None or df.empty for df in (zplm_marc_df, zplm_ta_ifs_df, marc_df)):
        return pd.DataFrame(columns=cols)

    def prep(df):
        out = df.copy()
        out.columns = [str(c).strip().upper() for c in out.columns]
        return out

    z = prep(zplm_marc_df)
    i = prep(zplm_ta_ifs_df)
    m = prep(marc_df)
    pick = lambda df, candidates: next((c for c in candidates if c in df.columns), None)
    zmat = pick(z, ["MATERIAL", "MATNR"])
    zplant = pick(z, ["PLANT", "WERKS"])
    ztype = pick(z, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])
    imat = pick(i, ["MATERIAL", "MATNR", "MATERIAL NUMBER", "PRODUCT", "PRODUCT NUMBER"])
    mmat = pick(m, ["MATERIAL", "MATNR"])
    mplant = pick(m, ["PLANT", "WERKS"])
    mstatus = pick(m, ["PLANT-SPECIFIC MATERIAL STATUS", "PLANT SPECIFIC MATERIAL STATUS", "MTSTB"])
    if not all([zmat, zplant, ztype, imat, mmat, mplant, mstatus]):
        return pd.DataFrame(columns=cols)

    z[zmat] = z[zmat].map(_normalize_code_as_str_basic).str.strip()
    z[zplant] = z[zplant].astype(str).str.strip().str.upper()
    z[ztype] = z[ztype].astype(str).str.strip().str.upper()
    i[imat] = i[imat].map(_normalize_code_as_str_basic).str.strip()
    m[mmat] = m[mmat].map(_normalize_code_as_str_basic).str.strip()
    m[mplant] = m[mplant].astype(str).str.strip().str.upper()
    m[mstatus] = m[mstatus].map(_normalize_code_as_str_basic)

    ifs_materials = set(i.loc[i[imat] != "", imat])
    status_map = (m.drop_duplicates([mmat, mplant])
                    .set_index([mmat, mplant])[mstatus].to_dict())
    transfer_map = build_material_transfer_map(zplm_marc_df)
    fp = z[z[ztype].eq("FP")].drop_duplicates([zmat, zplant])
    rows = []
    for _, row in fp.iterrows():
        material = row[zmat]
        plant = row[zplant]
        if transfer_map.get((material, plant), False):
            continue  # FP marked for transfer in ZPLM MARC — not counted
        if material and material not in ifs_materials:
            rows.append({
                "Material": material,
                "Plant": plant,
                "Material Plant Type": "FP",
                "Part Status": status_map.get((material, plant), ""),
                "Error": "FP material does not exist in ZPLM_TA_IFS",
                "Rule": "FP existence in ZPLM_TA_IFS",
            })
    return pd.DataFrame(rows, columns=cols)

# =============================================================================
# D) CROSS-EXISTENCE VALIDATIONS (MARA ↔ ZPLM MARA, MARC ↔ ZPLM MARC)
# =============================================================================

def validar_mara_vs_zplm_mara_existence(mara_df, zplm_mara_df):
    """FERT/HALB/ROH materials in MARA that are missing in ZPLM MARA."""
    if mara_df is None or mara_df.empty or zplm_mara_df is None or zplm_mara_df.empty:
        return pd.DataFrame(columns=["Material", "Error"])

    mara = mara_df.copy(); mara.columns = [c.strip().upper() for c in mara.columns]
    zplm = zplm_mara_df.copy(); zplm.columns = [c.strip().upper() for c in zplm.columns]

    m_mat = next((c for c in ["MATERIAL", "MATNR"] if c in mara.columns), None)
    m_mt  = next((c for c in ["MATERIAL TYPE", "MTART"] if c in mara.columns), None)
    z_mat = next((c for c in ["MATERIAL", "MATNR"] if c in zplm.columns), None)

    if not all([m_mat, m_mt, z_mat]):
        return pd.DataFrame(columns=["Material", "Error"])

    mara[m_mat] = mara[m_mat].astype(str).str.strip()
    mara[m_mt]  = mara[m_mt].astype(str).str.strip().str.upper()
    zplm[z_mat] = zplm[z_mat].astype(str).str.strip()

    mara_filtered = mara[mara[m_mt].isin({"FERT", "HALB", "ROH"})]
    missing = sorted(set(mara_filtered[m_mat]) - set(zplm[z_mat]))

    return pd.DataFrame([{"Material": m, "Error": "Missing in ZPLM MARA"} for m in missing],
                        columns=["Material", "Error"])


def validar_zplm_mara_vs_mara_existence(zplm_mara_df, mara_df):
    """Materials in ZPLM MARA that are not FERT/HALB/ROH in MARA."""
    if zplm_mara_df is None or zplm_mara_df.empty or mara_df is None or mara_df.empty:
        return pd.DataFrame(columns=["Material", "Error"])

    zplm = zplm_mara_df.copy(); zplm.columns = [c.strip().upper() for c in zplm.columns]
    mara = mara_df.copy();      mara.columns = [c.strip().upper() for c in mara.columns]

    z_mat = next((c for c in ["MATERIAL", "MATNR"] if c in zplm.columns), None)
    m_mat = next((c for c in ["MATERIAL", "MATNR"] if c in mara.columns), None)
    m_mt  = next((c for c in ["MATERIAL TYPE", "MTART"] if c in mara.columns), None)

    if not all([z_mat, m_mat, m_mt]):
        return pd.DataFrame(columns=["Material", "Error"])

    zplm[z_mat] = zplm[z_mat].astype(str).str.strip()
    mara[m_mat] = mara[m_mat].astype(str).str.strip()
    mara[m_mt]  = mara[m_mt].astype(str).str.strip().str.upper()

    mara_filtered = mara[mara[m_mt].isin({"FERT", "HALB", "ROH"})]
    extra = sorted(set(zplm[z_mat]) - set(mara_filtered[m_mat]))

    return pd.DataFrame(
        [{"Material": m, "Error": "Exists in ZPLM MARA but missing in MARA (FERT/HALB/ROH)"} for m in extra],
        columns=["Material", "Error"]
    )


def validar_marc_vs_zplm_marc_existence(marc_df, zplm_marc_df, mara_df):
    """
    FERT/HALB/ROH (Material, Plant) in MARC that are missing in ZPLM MARC.
    Includes materials that have NO PLANT in MARC → reports as Plant = 'No Plant'.
    """
    if marc_df is None or zplm_marc_df is None or marc_df.empty or zplm_marc_df.empty:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    marc = marc_df.copy();  marc.columns  = [c.strip().upper() for c in marc.columns]
    zplm = zplm_marc_df.copy(); zplm.columns = [c.strip().upper() for c in zplm.columns]
    mara = mara_df.copy();  mara.columns  = [c.strip().upper() for c in mara.columns]

    # Column detection
    m_mat   = next((c for c in ["MATERIAL","MATNR"] if c in marc.columns), None)
    m_plant = next((c for c in ["PLANT","WERKS"] if c in marc.columns), None)
    z_mat   = next((c for c in ["MATERIAL","MATNR"] if c in zplm.columns), None)
    z_plant = next((c for c in ["PLANT","WERKS"] if c in zplm.columns), None)
    mt_col  = next((c for c in ["MATERIAL TYPE","MTART"] if c in mara.columns), None)
    mara_mat= next((c for c in ["MATERIAL","MATNR"] if c in mara.columns), None)

    if not all([m_mat, z_mat, z_plant, mt_col, mara_mat]):
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    # Normalize
    marc[m_mat]   = marc[m_mat].astype(str).str.strip()
    zplm[z_mat]   = zplm[z_mat].astype(str).str.strip()
    zplm[z_plant] = zplm[z_plant].astype(str).str.strip().str.upper()
    mara[mara_mat] = mara[mara_mat].astype(str).str.strip()
    mara[mt_col]   = mara[mt_col].astype(str).str.strip().str.upper()

    # Materials in scope
    valid_mats = set(mara[mara[mt_col].isin({"FERT","HALB","ROH"})][mara_mat])

    # >>> MATERIALS WITHOUT PLANT IN MARC <<<
    marc_with_plants = marc[marc[m_mat].isin(valid_mats)]
    mats_with_plant = set(marc_with_plants[m_mat])
    mats_without_plant = sorted(valid_mats - mats_with_plant)

    # Build keys
    marc_keys = set(zip(marc_with_plants[m_mat], marc_with_plants[m_plant].str.upper()))
    zplm_keys = set(zip(zplm[z_mat], zplm[z_plant]))

    missing = sorted(marc_keys - zplm_keys)

    # Add missing materials with NO PLANT
    for m in mats_without_plant:
        missing.append((m, "No Plant"))

    return pd.DataFrame(
        [{"Material": m, "Plant": p, "Error": "Missing in ZPLM MARC"} for m, p in missing],
        columns=["Material", "Plant", "Error"]
    )


def validar_zplm_marc_vs_marc_existence(zplm_marc_df, marc_df, mara_df):
    """
    Entries in ZPLM MARC that are not in MARC (FERT/HALB/ROH).
    Includes entries where material exists but has NO PLANT in MARC → Plant = 'No Plant'.
    """
    if zplm_marc_df is None or marc_df is None or zplm_marc_df.empty or marc_df.empty:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    zplm = zplm_marc_df.copy(); zplm.columns = [c.strip().upper() for c in zplm.columns]
    marc = marc_df.copy();      marc.columns = [c.strip().upper() for c in marc.columns]
    mara = mara_df.copy();      mara.columns = [c.strip().upper() for c in mara.columns]

    z_mat   = next((c for c in ["MATERIAL","MATNR"] if c in zplm.columns), None)
    z_plant = next((c for c in ["PLANT","WERKS"] if c in zplm.columns), None)
    m_mat   = next((c for c in ["MATERIAL","MATNR"] if c in marc.columns), None)
    m_plant = next((c for c in ["PLANT","WERKS"] if c in marc.columns), None)
    mara_mat= next((c for c in ["MATERIAL","MATNR"] if c in mara.columns), None)
    mt_col  = next((c for c in ["MATERIAL TYPE","MTART"] if c in mara.columns), None)

    if not all([z_mat, z_plant, m_mat, mara_mat, mt_col]):
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    # Normalize
    zplm[z_mat]   = zplm[z_mat].astype(str).str.strip()
    zplm[z_plant] = zplm[z_plant].astype(str).str.strip().str.upper()
    marc[m_mat]   = marc[m_mat].astype(str).str.strip()
    if m_plant:
        marc[m_plant] = marc[m_plant].astype(str).str.strip().str.upper()

    mara[mara_mat] = mara[mara_mat].astype(str).str.strip()
    mara[mt_col]   = mara[mt_col].astype(str).str.strip().str.upper()

    # Material type restriction
    valid_mats = set(mara[mara[mt_col].isin({"FERT","HALB","ROH"})][mara_mat])

    # MATRIZES EM MARC
    marc_filtered = marc[marc[m_mat].isin(valid_mats)]
    marc_keys = set(zip(marc_filtered[m_mat], marc_filtered[m_plant])) if m_plant else set()

    # >>> Materials with NO PLANT in MARC <<<
    mats_in_marc = set(marc_filtered[m_mat])
    mats_without_plant = sorted(valid_mats - mats_in_marc)

    # ZPLM keys
    zplm_keys = set(zip(zplm[z_mat], zplm[z_plant]))

    extra = sorted(zplm_keys - marc_keys)

    # Add "No Plant" cases
    for m in mats_without_plant:
        if (m, "No Plant") in extra:
            continue
        extra.append((m, "No Plant"))

    return pd.DataFrame(
        [{"Material": m, "Plant": p, "Error": "Exists in ZPLM MARC but missing in MARC (FERT/HALB/ROH)"} 
         for m, p in extra],
        columns=["Material", "Plant", "Error"]
    )


# =============================================================================
# E) POST-PROCESSING HELPERS
# =============================================================================

def add_old_material_number_errors(mara_df, errors_df):
    """
    Validates MARA 'Old material number':
    - Must not be empty.
    - Must equal Material without '-' characters.
    Appends errors to errors_df and returns the result.
    """
    if 'Material' not in mara_df.columns or 'Old material number' not in mara_df.columns:
        print("⚠️ MARA must contain 'Material' and 'Old material number' columns.")
        return errors_df

    materials = errors_df['Material'].dropna().astype(str).str.strip().unique()
    mara_indexed = mara_df.set_index('Material')
    additional_rows = []

    for material in materials:
        expected = material.replace('-', '')
        old_val  = mara_indexed.at[material, 'Old material number'] if material in mara_indexed.index else None
        old_str  = '' if pd.isna(old_val) else str(old_val).strip()

        if old_str == '':
            error_msg = 'Mandatory field blank'
        elif old_str != expected:
            error_msg = (f"Invalid value: should equal Material without '-' "
                         f"(expected: '{expected}', current: '{old_str}')")
        else:
            error_msg = None

        if error_msg:
            additional_rows.append({
                'Line': '', 'Field': 'Old material number', 'Error': error_msg,
                'Material': material, 'Plant': '',
                'View': None, 'Justified_By_Missing_View': False,
            })

    if additional_rows:
        errors_df = pd.concat([errors_df, pd.DataFrame(additional_rows)], ignore_index=True)

    return errors_df


def fill_plant_and_validate_enterproj_from_zplm(errors_df, latest_report,
                                                  zplm_marc_file=None, zplm_mara_file=None):
    """
    Enriches 'Enterproj Number' errors: fills Plant from ZPLM MARC and
    cross-validates Enterproj Number against the previous report and ZPLM MARA.
    """
    if errors_df is None or errors_df.empty:
        return errors_df
    for c in ['Field', 'Material', 'Plant', 'Error']:
        if c not in errors_df.columns:
            return errors_df

    mask = (
        errors_df['Field'].astype(str).str.strip() == 'Enterproj Number'
    ) & (
        errors_df['Plant'].astype(str).str.strip().isin(['', 'nan'])
    )
    if not mask.any():
        return errors_df

    subset = errors_df.loc[mask].copy()
    zplm_marc_df = load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None
    zplm_mara_df = load_xlsx_fast(zplm_mara_file, sheet_name=0) if zplm_mara_file else None

    if zplm_marc_df is None or zplm_mara_df is None:
        return errors_df

    zplm_marc_df.columns = [c.strip().upper() for c in zplm_marc_df.columns]
    zplm_mara_df.columns = [c.strip().upper() for c in zplm_mara_df.columns]

    if not {'MATERIAL', 'PLANT'}.issubset(zplm_marc_df.columns):
        return errors_df
    if not {'MATERIAL', 'ENTERPROJ NUMBER'}.issubset(zplm_mara_df.columns):
        return errors_df

    try:
        prev_df = load_xlsx_fast(latest_report, sheet_name="Original Data")
    except Exception:
        prev_df = pd.DataFrame(columns=['Material', 'Enterproj Number'])

    def norm_str(x):
        return '' if pd.isna(x) else str(x).strip()

    prev_map = dict(zip(prev_df.get('Material', pd.Series(dtype=str)).astype(str).str.strip(),
                        prev_df.get('Enterproj Number', pd.Series(dtype=str)).astype(str).str.strip()))
    zplm_map = dict(zip(zplm_mara_df['MATERIAL'].astype(str).str.strip(), zplm_mara_df['ENTERPROJ NUMBER']))

    zplm_marc_filt = zplm_marc_df[['MATERIAL', 'PLANT']].drop_duplicates()
    subset['Material_norm'] = subset['Material'].astype(str).str.strip()
    zplm_marc_filt['MATERIAL_norm'] = zplm_marc_filt['MATERIAL'].astype(str).str.strip()

    merged = subset.merge(zplm_marc_filt[['MATERIAL_norm', 'PLANT']], left_on='Material_norm',
                          right_on='MATERIAL_norm', how='left')
    merged['Plant'] = merged['PLANT'].where(merged['PLANT'].notna(), merged['Plant'])
    merged = merged.drop(columns=['MATERIAL_norm', 'Material_norm', 'PLANT', 'MATERIAL'], errors='ignore')

    updated_rows = []
    for _, r in merged.iterrows():
        mat      = str(r['Material']).strip()
        prev_val = norm_str(prev_map.get(mat, ''))
        zplm_val = norm_str(zplm_map.get(mat, ''))
        extras = []
        if zplm_val == '':
            extras.append("Mandatory field blank in ZPLM MARA (Enterproj Number)")
        elif prev_val != '' and zplm_val != prev_val:
            extras.append(f"Invalid value: should equal first report Enterproj Number (expected: '{prev_val}', current: '{zplm_val}')")
        base_err = norm_str(r['Error'])
        if extras:
            r['Error'] = (base_err + (' \n' if base_err else '') + ' \n'.join(extras)).strip()
        updated_rows.append(r)

    enriched = pd.DataFrame(updated_rows)
    return pd.concat([errors_df.loc[~mask].copy(), enriched], ignore_index=True)


def validar_neg_stocks_in_plant(marc_df, mara_df):
    """
    MARC + MARA: Neg. stocks in plant validation.

    Scope:
      - Validate only Material Types FERT, HALB and ROH.

    Rules:
      - TN02:
          - The field must always be blank for FERT, HALB and ROH.

      - All other plants:
          - ROH and HALB:
              - The field must be marked.
          - FERT:
              - The field must not be marked.

    Accepted marked values:
      - X
      - 1
      - 1.0
      - YES
      - Y
      - TRUE

    Accepted blank/inactive values:
      - Blank
      - 0
      - 0.0
      - NO
      - N
      - FALSE
    """

    if (
        marc_df is None
        or marc_df.empty
        or mara_df is None
        or mara_df.empty
    ):
        return []

    required_marc_columns = {
        'Material',
        'Plant',
        'Neg. stocks in plant',
    }

    required_mara_columns = {
        'Material',
        'Material Type',
    }

    if (
        not required_marc_columns.issubset(marc_df.columns)
        or not required_mara_columns.issubset(mara_df.columns)
    ):
        return []

    # ------------------------------------------------------------
    # Build Material -> Material Type mapping
    # ------------------------------------------------------------
    material_type_map = dict(
        zip(
            mara_df['Material']
            .fillna('')
            .astype(str)
            .str.strip(),

            mara_df['Material Type']
            .fillna('')
            .astype(str)
            .str.strip()
            .str.upper()
        )
    )

    # ------------------------------------------------------------
    # Normalization helpers
    # ------------------------------------------------------------
    def _normalize_string(value):
        if value is None or pd.isna(value):
            return ''

        return str(value).strip()

    def _normalize_checkbox(value):
        """
        Normalize the checkbox value.

        Returns:
          - 'X' when marked
          - '' when blank or inactive
          - 'INVALID' for an unrecognized non-empty value
        """

        normalized_value = _normalize_string(value).upper()

        if normalized_value in {
            'X',
            '1',
            '1.0',
            'YES',
            'Y',
            'TRUE',
        }:
            return 'X'

        if normalized_value in {
            '',
            '0',
            '0.0',
            'NO',
            'N',
            'FALSE',
        }:
            return ''

        return 'INVALID'

    errors = []

    # ------------------------------------------------------------
    # Validate each MARC record
    # ------------------------------------------------------------
    for idx, row in marc_df.iterrows():
        material = _normalize_string(
            row.get('Material')
        )

        plant = _normalize_string(
            row.get('Plant')
        ).upper()

        material_type = material_type_map.get(
            material,
            ''
        )

        # Validate only FERT, HALB and ROH
        if material_type not in {'FERT', 'HALB', 'ROH'}:
            continue

        raw_value = _normalize_string(
            row.get('Neg. stocks in plant')
        ).upper()

        normalized_value = _normalize_checkbox(
            row.get('Neg. stocks in plant')
        )

        current_value = (
            raw_value
            if raw_value != ''
            else 'blank'
        )

        # ========================================================
        # TN02 override
        #
        # For TN02, the field must always be blank, regardless of
        # whether the Material Type is FERT, HALB or ROH.
        # ========================================================
        if plant == 'TN02':
            if normalized_value == 'X':
                errors.append({
                    'Line': idx + 2,
                    'Field': 'Neg. stocks in plant',
                    'Error': (
                        "Field must be blank for TN02 "
                        f"(Material Type: {material_type}, "
                        f"current: '{current_value}')"
                    ),
                    'Material': material,
                    'Plant': plant,
                })

            elif normalized_value == 'INVALID':
                errors.append({
                    'Line': idx + 2,
                    'Field': 'Neg. stocks in plant',
                    'Error': (
                        f"Invalid value: '{current_value}'. "
                        "Field must be blank for TN02 "
                        f"(Material Type: {material_type})"
                    ),
                    'Material': material,
                    'Plant': plant,
                })

            # TN02 must not enter the standard ROH/HALB/FERT rules.
            continue

        # ========================================================
        # All plants except TN02
        # ========================================================

        # --------------------------------------------------------
        # ROH / HALB: field must be marked
        # --------------------------------------------------------
        if material_type in {'ROH', 'HALB'}:
            if normalized_value == '':
                errors.append({
                    'Line': idx + 2,
                    'Field': 'Neg. stocks in plant',
                    'Error': (
                        "Mandatory field blank: must be marked "
                        f"for Material Type {material_type}"
                    ),
                    'Material': material,
                    'Plant': plant,
                })

            elif normalized_value == 'INVALID':
                errors.append({
                    'Line': idx + 2,
                    'Field': 'Neg. stocks in plant',
                    'Error': (
                        f"Invalid value: '{current_value}'. "
                        f"Field must be marked for Material Type "
                        f"{material_type}"
                    ),
                    'Material': material,
                    'Plant': plant,
                })

        # --------------------------------------------------------
        # FERT: field must be blank
        # --------------------------------------------------------
        elif material_type == 'FERT':
            if normalized_value == 'X':
                errors.append({
                    'Line': idx + 2,
                    'Field': 'Neg. stocks in plant',
                    'Error': (
                        "Field must be blank for Material Type FERT "
                        f"(current: '{current_value}')"
                    ),
                    'Material': material,
                    'Plant': plant,
                })

            elif normalized_value == 'INVALID':
                errors.append({
                    'Line': idx + 2,
                    'Field': 'Neg. stocks in plant',
                    'Error': (
                        f"Invalid value: '{current_value}'. "
                        "Field must be blank for Material Type FERT"
                    ),
                    'Material': material,
                    'Plant': plant,
                })

    return errors

# =============================================================================
# SAP VALIDATION ENGINE — PART 3: ORCHESTRATION & WORKFLOWS
# =============================================================================
# Depends on: Part 1 (imports, loaders, helpers) + Part 2 (validation functions)
#
# SECTIONS:
#   1) postprocess_with_views_roles_and_classification
#   2) run_full_validation        — main SAPUI5 validation pipeline
#   3) run_fp_status_validation   — FP family classification
#   4) run_data_cleanups_validation — Data Cleanup pipeline + Excel export
#   5) run_mrp_validation         — MRP field validation
#   6) run_plm_workflow_combined  — ATL / PLM workflow
#   7) compare_adjusted_reports   — diff between two reports
#   8) Log payload builders + persistence helpers
# =============================================================================


# =============================================================================
# 1) POST-PROCESSING: VIEWS / ROLES / CLASSIFICATION
# =============================================================================

def postprocess_with_views_roles_and_classification(
    errors_df: pd.DataFrame,
    mara_df:   pd.DataFrame,
    marc_df:   pd.DataFrame,
    roles_df:  pd.DataFrame,
    views_df:  pd.DataFrame,
    plant_filter: str | None = None,
):
    """
    Full post-processing pipeline:
      - Maps Field -> View (via Roles)
      - Computes Letters_Present / Missing per (Material, Plant) from MARC 'Maintenance status'
      - Generates views_missing_df (per Material/Plant)
      - Flags Justified_By_Missing_View and classifies Error_Type (Missing View / Field Error)
      - Builds coverage_df (how many materials are missing each required view)
      - Builds view_tabs_dict (per-view detail tabs: Missing_<View>)
      - Constructs contagem_final_df (Adjusted Summary) with Role/View_Description overrides
      - Renames 'Export/Import Group' -> 'Intrastat Group'

    Returns:
        errors_df, errors_filtered_df, views_missing_output_df,
        contagem_final_df, coverage_df, view_tabs_dict
    """
    if errors_df is None:
        errors_df = pd.DataFrame()
    if roles_df is None or views_df is None:
        empty = pd.DataFrame()
        return errors_df, empty, empty, empty, empty, {}

    roles_df  = roles_df.copy()
    views_df  = views_df.copy()
    marc_df   = marc_df.copy()
    mara_df   = mara_df.copy()
    errors_df = errors_df.copy()

    # -----------------------------------------------
    # 1) Apply plant filter early
    # -----------------------------------------------
    if plant_filter is not None:
        plant_filter = str(plant_filter).strip().upper()

        mats_in_plant = set()
        if 'Material' in marc_df.columns and 'Plant' in marc_df.columns:
            mats_in_plant = set(
                marc_df.loc[marc_df['Plant'].astype(str).str.strip().str.upper() == plant_filter, 'Material']
                       .astype(str).str.strip()
            )

        if 'Plant' in marc_df.columns:
            marc_df = marc_df[marc_df['Plant'].astype(str).str.strip().str.upper() == plant_filter].copy()

        if 'Plant' in errors_df.columns and 'Material' in errors_df.columns:
            plant_col = errors_df['Plant'].astype(str).str.strip().str.upper()
            mat_col   = errors_df['Material'].astype(str).str.strip()
            errors_df = errors_df[(plant_col == plant_filter) |
                                  ((plant_col == '') & mat_col.isin(mats_in_plant))].copy()

    # Normalise headers
    for df in [roles_df, views_df, marc_df, mara_df, errors_df]:
        df.columns = df.columns.astype(str).str.strip()

    # -----------------------------------------------
    # 2) Base mappings: Field -> View, View -> Description
    # -----------------------------------------------
    if 'Field' in roles_df.columns and 'View' in roles_df.columns:
        roles_df['Field'] = roles_df['Field'].astype(str).str.strip()
        roles_df['View']  = roles_df['View'].astype(str).str.strip()
        field_to_view = roles_df.set_index('Field')['View'].to_dict()
    else:
        field_to_view = {}

    if 'Codigo' in views_df.columns and 'Descricao' in views_df.columns:
        views_df['Codigo']    = views_df['Codigo'].astype(str).str.strip()
        views_df['Descricao'] = views_df['Descricao'].astype(str).str.strip()
        view_to_desc     = views_df.set_index('Codigo')['Descricao'].to_dict()
        expected_letters = set(views_df['Codigo'].dropna().unique())
    else:
        view_to_desc, expected_letters = {}, set()

    if 'Field' in errors_df.columns:
        errors_df['View'] = errors_df['Field'].map(field_to_view).fillna('Not applicable')
    else:
        errors_df['View'] = 'Not applicable'

    # -----------------------------------------------
    # 3) Extract Letters_Present per (Material, Plant)
    # -----------------------------------------------
    if 'Maintenance status' not in marc_df.columns:
        marc_df['Maintenance status'] = ''

    marc_df['Letters_Present'] = marc_df['Maintenance status'].apply(
        lambda s: [] if pd.isna(s) else list(str(s).strip())
    )

    mt_map = dict(zip(
        mara_df.get('Material', pd.Series(dtype=str)).astype(str).str.strip(),
        mara_df.get('Material Type', pd.Series(dtype=str)).astype(str).str.strip()
    ))

    if 'Plant Type' in marc_df.columns:
        marc_df['_KEY'] = list(zip(
            marc_df['Material'].astype(str).str.strip(),
            marc_df['Plant'].astype(str).str.strip()
        ))
        pt_map = dict(zip(marc_df['_KEY'], marc_df['Plant Type'].astype(str).str.strip()))
    else:
        pt_map = {}

    # -----------------------------------------------
    # 4) Views Missing per (Material, Plant)
    # -----------------------------------------------
    valid_types  = {'FERT', 'HALB', 'ROH'}
    view_to_fields = roles_df.groupby('View')['Field'].apply(list).to_dict()

    valid_mats = mara_df.loc[
        mara_df['Material Type'].astype(str).str.strip().isin(valid_types), 'Material'
    ].astype(str).str.strip().unique()

    marc_filtered = marc_df[marc_df['Material'].astype(str).str.strip().isin(valid_mats)].copy()

    missing_rows = []
    for _, row in marc_filtered.iterrows():
        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()
        present  = set(row.get('Letters_Present', []) or [])
        missing  = expected_letters - present

        fields_missing       = []
        descriptions_missing = []
        for letter in missing:
            fields_missing.extend(view_to_fields.get(letter, []))
            descriptions_missing.append(view_to_desc.get(letter, ''))

        missing_rows.append({
            'Material':              material,
            'Plant':                 plant,
            'Letters_Present':       ','.join(sorted(present)),
            'Letters_Missing':       ','.join(sorted(missing)),
            'Fields_Missing':        ','.join(sorted(fields_missing)),
            'Descriptions_Missing':  ','.join(sorted(descriptions_missing)),
        })

    views_missing_df = pd.DataFrame(missing_rows)

    # -----------------------------------------------
    # 5) Classify Error_Type + Justified_By_Missing_View
    # Pre-build a dict (Material, Plant) -> set of missing view letters for O(1) lookup
    # -----------------------------------------------
    _missing_lookup: dict[tuple, set] = {}
    if not views_missing_df.empty:
        for _, _vmrow in views_missing_df.iterrows():
            _key = (_vmrow['Material'], _vmrow['Plant'])
            _raw = _vmrow.get('Letters_Missing', '')
            _missing_lookup[_key] = set(_raw.split(',')) if isinstance(_raw, str) and _raw else set()

    def classify_error(row):
        view = str(row.get('View', '')).strip()
        if view not in ('', 'Not applicable'):
            _key = (str(row.get('Material', '')).strip(), str(row.get('Plant', '')).strip())
            if view in _missing_lookup.get(_key, set()):
                return 'Missing View'
        return 'Field Error'

    if not errors_df.empty:
        errors_df['Error_Type']              = errors_df.apply(classify_error, axis=1)
        errors_df['Justified_By_Missing_View'] = errors_df['Error_Type'] == 'Missing View'
    else:
        errors_df['Error_Type']              = pd.Series(dtype=str)
        errors_df['Justified_By_Missing_View'] = pd.Series(dtype=bool)

    # -----------------------------------------------
    # 6) Split by type
    # -----------------------------------------------
    errors_filtered_df      = errors_df[errors_df['Justified_By_Missing_View']].copy()
    views_missing_output_df = errors_df[~errors_df['Justified_By_Missing_View']].copy()

    # -----------------------------------------------
    # 7) Adjusted Summary (Field Errors only)
    # -----------------------------------------------
    if not views_missing_output_df.empty and 'Field' in views_missing_output_df.columns:
        contagem_final_df = (
            views_missing_output_df['Field']
            .value_counts()
            .rename_axis('Field')
            .reset_index(name='Adjusted Nº of Errors')
        )
    else:
        contagem_final_df = pd.DataFrame(columns=['Field', 'Adjusted Nº of Errors'])

    if not contagem_final_df.empty:
        contagem_final_df['View']             = contagem_final_df['Field'].map(field_to_view).fillna('Not applicable')
        contagem_final_df['View_Description'] = contagem_final_df['View'].map(view_to_desc).fillna('Not applicable')

        if 'Role' in roles_df.columns:
            roles_df['Role']  = roles_df['Role'].astype(str).str.strip()
            field_to_role     = roles_df.set_index('Field')['Role'].to_dict()
            contagem_final_df['Role'] = contagem_final_df['Field'].map(field_to_role).fillna('Not applicable')

        # Specific overrides
        mask_override = contagem_final_df['Field'].isin(['Enterproj Number', 'Plant Type'])
        contagem_final_df.loc[mask_override, 'View_Description'] = 'Product Master Visteon Data'

    # -----------------------------------------------
    # 8) Field renames (Export/Import Group -> Intrastat Group)
    # -----------------------------------------------
    field_renames = {'Export/Import Group': 'Intrastat Group'}

    def apply_field_renames(df_):
        df_ = df_.copy()
        if 'Field' in df_.columns:
            df_['Field'] = df_['Field'].replace(field_renames)
        return df_

    errors_df               = apply_field_renames(errors_df)
    errors_filtered_df      = apply_field_renames(errors_filtered_df)
    views_missing_output_df = apply_field_renames(views_missing_output_df)
    contagem_final_df       = apply_field_renames(contagem_final_df)

    for col in ['View', 'View_Description', 'Role']:
        if col in contagem_final_df.columns:
            contagem_final_df[col] = contagem_final_df[col].astype(str).replace(
                {'nan': 'Not applicable', 'None': 'Not applicable', '': 'Not applicable'}
            )

    # -----------------------------------------------
    # 9) Coverage per View
    # -----------------------------------------------
    field_rules = {
        'Production Scheduling Profile':         lambda mt, pt: mt in ['FERT', 'HALB'],
        'REM Profile':                           lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Repetitive Manufacturing':              lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Enterproj Number':                      lambda mt, pt: mt == 'FERT',
        'Procurement Type':                      lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'] and pt in ['PP', 'MP'],
        'Account Assignment Group':              lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Automatic Purchase Order Allowed':      lambda mt, pt: pt == 'PP',
        'Batch Management Required Ind.':        lambda mt, pt: pt in ['PP', 'COI'], 
        'Source List Required':                  lambda mt, pt: pt == 'PP',
        'Purchasing Group':                      lambda mt, pt: pt == 'PP',
        'MRP Type':                              lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Planning Time Fence':                   lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Lot Sizing Procedure':                  lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Per.Lot Sz.Plng Cal.':                  lambda mt, pt: mt in ['FERT', 'HALB'],
        'With Qty Structure':                    lambda mt, pt: mt in ['FERT', 'HALB'],
        'Item Category Group':                   lambda mt, pt: True,
        'General Item Category Group':           lambda mt, pt: True,
        'Item Category Group (MVKE)':            lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Loading Group':                         lambda mt, pt: True,
        'Division':                              lambda mt, pt: True,
        'Export/Import Group':                   lambda mt, pt: True,
        'Intrastat Group':                       lambda mt, pt: True,
        'Product Hierarchy':                     lambda mt, pt: mt == 'FERT',
        'Price Control':                         lambda mt, pt: True,
        'Price Unit':                            lambda mt, pt: True,
        'Valuation Class':                       lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Gross Weight':                          lambda mt, pt: pt == 'FP',
        'Net Weight':                            lambda mt, pt: pt == 'FP',
        'Planning Calendar':                     lambda mt, pt: True,
        'Period Type':                           lambda mt, pt: True,
        'Strategy Group':                        lambda mt, pt: True,
        'MRP Controller':                        lambda mt, pt: True,
        'Plant Specific Material Status':        lambda mt, pt: True,
        'Backflush':                             lambda mt, pt: True,
        'Product-Related Origin':                lambda mt, pt: True,
        'Availability Check':                    lambda mt, pt: True,
        'Neg. stocks in plant':                  lambda mt, pt: mt in ['ROH', 'HALB'],
        'Consumption Mode':                      lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Bwd consumption per.':                  lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Fwd consumption per.':                  lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Profit Center':                         lambda mt, pt: True,
        'Material Group':                        lambda mt, pt: True,
        'JIT Delivery Sched.':                   lambda mt, pt: pt == 'PP',
        'Issue Storage Location':                lambda mt, pt: True,
        'Cntry/Reg of Origin':                   lambda mt, pt: pt == 'FP',
        #'Commodity Code':                        lambda mt, pt: pt == 'FP',
        'LIFO/FIFO-relevant':                    lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Stock Determination Group':             lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Storage Location for Ext. Procurement': lambda mt, pt: pt == 'PP',
        'Individual/ Collective Requirements':   lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Packaging Mat. Type':                   lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],   
        'WH Material Group':                     lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Production Supervisor':                 lambda mt, pt: mt in ['FERT', 'HALB'],   
        'Batch Management(Plant)':               lambda mt, pt: pt in ['PP', 'COI'],
        'Advanced Planning':                     lambda mt, pt: True,
        'Heuristic':                             lambda mt, pt: True,
        'PP Plng Procedure':                     lambda mt, pt: True,
        'Proposed Strategy':                     lambda mt, pt: True,
        'Version Indicator':                     lambda mt, pt: pt in ['FP', 'MP'],
        'Costing Lot Size':                      lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Safety Time':                           lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
    }

    def is_view_required(material_type, plant_type, view_code):
        for field in view_to_fields.get(view_code, []):
            rule = field_rules.get(field)
            if rule and rule(material_type, plant_type):
                return True
        return False

    view_missing_count = {letter: 0 for letter in expected_letters}
    for _, r in marc_filtered.iterrows():
        mt      = str(mt_map.get(str(r['Material']).strip(), '')).strip().upper()
        pt      = str(pt_map.get((str(r['Material']).strip(), str(r['Plant']).strip()), '')).strip().upper()
        present = set(r.get('Letters_Present', []) or [])
        for v in expected_letters:
            if v not in present and is_view_required(mt, pt, v):
                view_missing_count[v] += 1

    coverage_df = pd.DataFrame([
        {'View': v, 'View_Description': view_to_desc.get(v, ''),
         'Materiais without Active View': cnt,
         'Link': f'=HYPERLINK("#Missing_{v}!A1","Check Details")'}
        for v, cnt in view_missing_count.items() if cnt > 0
    ])

    # -----------------------------------------------
    # 10) Missing_<View> detail tabs
    # -----------------------------------------------
    view_tabs_dict = {}
    for v in expected_letters:
        rows = []
        for _, r in marc_filtered.iterrows():
            mt      = str(mt_map.get(str(r['Material']).strip(), '')).strip().upper()
            pt      = str(pt_map.get((str(r['Material']).strip(), str(r['Plant']).strip()), '')).strip().upper()
            present = set(r.get('Letters_Present', []) or [])
            if v not in present and is_view_required(mt, pt, v):
                rows.append({
                    'Material':         str(r['Material']).strip(),
                    'Plant':            str(r['Plant']).strip(),
                    'Missing_View':     v,
                    'View_Description': view_to_desc.get(v, ''),
                    'Fields_Expected':  ','.join(view_to_fields.get(v, [])),
                })
        if rows:
            df_view = pd.DataFrame(rows)
            if not df_view.empty:
                df_view = df_view.drop_duplicates(subset=['Material', 'Plant', 'Missing_View']).reset_index(drop=True)
            view_tabs_dict[v] = df_view
    
    # -------------------------------------------------
    # ADD MRP Controller from MARC (Material + Plant)
    # -------------------------------------------------
    if marc_df is not None and not marc_df.empty and "MRP Controller" in marc_df.columns:

        marc_tmp = marc_df[["Material", "Plant", "MRP Controller"]].copy()

        marc_tmp["Material"] = marc_tmp["Material"].astype(str).str.strip()
        marc_tmp["Plant"] = marc_tmp["Plant"].astype(str).str.strip()

        # Avoid repeated merges when MARC contains duplicate Material+Plant rows.
        marc_tmp = marc_tmp.drop_duplicates(subset=["Material", "Plant"], keep="first")

        views_missing_output_df["Material"] = views_missing_output_df["Material"].astype(str).str.strip()
        views_missing_output_df["Plant"] = views_missing_output_df["Plant"].astype(str).str.strip()

        views_missing_output_df = views_missing_output_df.merge(
            marc_tmp,
            on=["Material", "Plant"],
            how="left"
        )

        if not views_missing_output_df.empty:
            dedup_cols = [c for c in ['Material', 'Plant', 'Field', 'Error'] if c in views_missing_output_df.columns]
            if dedup_cols:
                views_missing_output_df = views_missing_output_df.drop_duplicates(subset=dedup_cols).reset_index(drop=True)
    else:
        views_missing_output_df["MRP Controller"] = ""

    # -------------------------------------------------
    # REMOVE "View" only from Field With Error dataset
    # -------------------------------------------------
    if "View" in views_missing_output_df.columns:
        views_missing_output_df = views_missing_output_df.drop(columns=["View"])

    return errors_df, errors_filtered_df, views_missing_output_df, contagem_final_df, coverage_df, view_tabs_dict


# =============================================================================
# 2) RUN FULL VALIDATION — Main SAPUI5 Pipeline
# =============================================================================

def run_full_validation(
    mara_file,
    marc_file,
    mvke_file,
    marm_file=None,
    mbew_file=None,
    roles_file=None,
    zplm_marc_file=None,
    zplm_mara_file=None,
    previous_report=None,
    plant_filter: str | None = None,
    progress_callback=None,
    csr_item_file=None,
    csr_head_file=None,
    matplsta_file=None,
    mard_file=None,
):
    """
    Runs the complete Palmela validation pipeline:
      1. Load files
      2. Pre-normalisation + build sap_like_df (MARC + MARA + ZPLM MARC)
      3. All field validations
      4. Post-processing (Views / Roles / Summary)
      5. Excel generation in-memory
    Returns: (excel_bytes, errors_df, summary_df)

    NOTE: This engine no longer depends on the Item Browse Report / SAPUI5
    Export. Like the Tunisia engine, it builds an equivalent flat dataframe
    (`sap_like_df`) from MARC + MARA + ZPLM MARC via build_sap_like_df(),22
    and runs the exact same validar_* functions against it.
    """
    def _cb(pct, msg):
        if progress_callback:
            progress_callback(pct, msg)

    # ------------------------------------------------------------------
    # 1) Load DataFrames
    # ------------------------------------------------------------------
    _cb(7,  "Loading MARA...")
    mara_df = load_xlsx_fast(mara_file, sheet_name=0)

    _cb(10, "Loading MARC...")
    marc_df = load_xlsx_fast(marc_file, sheet_name=0)

    _cb(12, "Loading MVKE...")
    mvke_df = load_xlsx_fast(mvke_file, sheet_name=0)

    if roles_file is None:
        roles_file = get_rolesviews_bytesio()
    rolesviews_dict = load_xlsx_fast(roles_file, sheet_name=None)

    _cb(14, "Loading ZPLM data...")
    zplm_marc_df = load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None

    _cb(16, "Preparing ZPLM data...")
    zplm_mara_df = load_xlsx_fast(zplm_mara_file, sheet_name=0) if zplm_mara_file else None

    _cb(18, "Loading CSR data...")
    csr_item_df = load_xlsx_fast(csr_item_file, sheet_name=0) if csr_item_file else None
    csr_head_df = load_xlsx_fast(csr_head_file, sheet_name=0) if csr_head_file else None

    _cb(19, "Loading MATPLSTA data...")
    matplsta_df = load_xlsx_fast(matplsta_file, sheet_name=0) if matplsta_file else None

    _cb(20, "Loading MARM data...")
    marm_df = load_xlsx_fast(marm_file, sheet_name=0) if marm_file else None

    _cb(21, "Loading MBEW data...")
    mbew_df = load_xlsx_fast(mbew_file, sheet_name=0) if mbew_file else None

    _cb(21, "Loading MARD data...")
    mard_df = load_xlsx_fast(mard_file, sheet_name=0) if mard_file else None

    # ------------------------------------------------------------------
    # 2) Basic sanity checks
    # ------------------------------------------------------------------
    if mara_df is None:
        raise ValueError("MARA.xlsx not loaded.")
    if marc_df is None:
        raise ValueError("MARC.xlsx not loaded.")
    if mvke_df is None:
        raise ValueError("MVKE.xlsx not loaded.")
    if marm_df is None:
        raise ValueError("MARM.xlsx not loaded.")
    if mbew_df is None:
        raise ValueError("MBEW.xlsx not loaded.")
    if "Roles" not in rolesviews_dict or "Views" not in rolesviews_dict:
        raise ValueError("RolesViews.xlsx must contain sheets 'Roles' and 'Views'.")

    roles_df = rolesviews_dict["Roles"]
    views_df = rolesviews_dict["Views"]

    # ------------------------------------------------------------------
    # 2b) Column normalisation — same rename maps used by the TN engine,
    # since raw MARC/MBEW/ZPLM MARC/MVKE/MARM exports may use column names
    # that differ slightly from what the validar_* functions expect (these
    # used to be normalised for free by the Item Browse Report's own column
    # layout; now that we build sap_like_df ourselves, we normalise here).
    # ------------------------------------------------------------------
    marc_df = _normalize_df_columns(marc_df, {
        'plant': 'Plant',
        'werks': 'Plant',
        'intrastat group': 'Export/Import Group',
        'period indicator': 'Period Type',
        'material plant type': 'Plant Type',
        'purchasing group': 'Purchasing Group',
        'source list': 'Source list',
        'source list required': 'Source list required',
        'automatic po': 'Automatic Purchase Order Allowed',
        'automatic purchase order': 'Automatic Purchase Order Allowed',
        'issue storage location': 'Issue Storage Location',
        'production storage location': 'Issue Storage Location',
        'planning strategy group': 'Planning Strategy Group',
        'plant-specific material status': 'Plant Specific Material Status',
    })
    # ✅ 'Strategy Group' alias — validar_strategy_group() (generic, Palmela-only)
    # expects this exact name (matching the old Item Browse Report column),
    # while everything else (including TN's own tn_validar_strategy_group)
    # was written against 'Planning Strategy Group'. Keep both.
    if 'Planning Strategy Group' in marc_df.columns:
        marc_df['Strategy Group'] = marc_df['Planning Strategy Group']

    mbew_df = _normalize_df_columns(mbew_df, {
        'plant': 'Plant',
        'werks': 'Plant',
        'valuation area': 'Plant',
        'bwkey': 'Plant',
        'price control': 'Price Control',
        'price unit': 'Price Unit',
        'valuation class': 'Valuation Class',
        'cost est. with qty structure': 'With Qty Structure',
    })
    # ✅ 'Valuation Area' alias — validar_lifo_fifo_relevant_mbew() hard-requires
    # a column literally named 'Valuation Area' (MBEW's own SAP field name),
    # while everything else here was normalized to 'Plant'. Keep both.
    if mbew_df is not None and 'Plant' in mbew_df.columns:
        mbew_df['Valuation Area'] = mbew_df['Plant']

    zplm_marc_df = _normalize_df_columns(zplm_marc_df, {
        'plant': 'Plant',
        'werks': 'Plant',
        'material plant type': 'Plant Type',
        'enterproj number': 'Enterproj Number',
        'material transfer': 'Material Transfer',
        'eop date': 'EOP Date',
        'service liability date': 'Service Liability Date',
        'ltb quantity': 'LTB Quantity',
        'ltb start date': 'LTB Start Date',
    })
    # ✅ Keep a 'Material Plant Type' alias alongside the normalized 'Plant Type'.
    # validar_repetitive_manufacturing() and validar_rem_profile() read
    # zplm_marc_df directly (not sap_like_df) and expect the original raw
    # column name; every other zplm_marc_df consumer expects 'Plant Type'.
    if zplm_marc_df is not None and 'Plant Type' in zplm_marc_df.columns:
        zplm_marc_df['Material Plant Type'] = zplm_marc_df['Plant Type']

    # ✅ DEDUPLICATE ZPLM MARC to one row per (Material, Plant) — BEFORE any
    # consumer touches it (build_sap_like_df AND the several validar_*
    # functions that merge marc_df with zplm_marc_df directly: Storage
    # Location for Ext. Procurement, Version Indicator, Batch Management
    # (Plant), Country/Region of Origin, Repetitive Manufacturing, REM
    # Profile). ZPLM MARC carries change-history columns (Original/Change
    # Number, Created/Changed By/On), which strongly suggests it can
    # legitimately hold more than one historical row per (Material, Plant).
    # A left.merge(zplm_marc_df, on=['Material','Plant']) multiplies every
    # matching row, so without this, every one of those validations would
    # inflate massively. We keep the most recently changed row per
    # (Material, Plant) when a date/change column is available, else just
    # the first occurrence.
    if zplm_marc_df is not None and not zplm_marc_df.empty and {'Material', 'Plant'}.issubset(zplm_marc_df.columns):
        _sort_col = next(
            (c for c in ['Changed On', 'Change Number', 'Created On'] if c in zplm_marc_df.columns),
            None
        )
        if _sort_col:
            # Convert to string to handle mixed datetime/time types
            zplm_marc_df[_sort_col] = zplm_marc_df[_sort_col].fillna('').astype(str)
            zplm_marc_df = zplm_marc_df.sort_values(_sort_col, ascending=False, na_position='last')
        zplm_marc_df = zplm_marc_df.drop_duplicates(subset=['Material', 'Plant'], keep='first')

    mvke_df = _normalize_df_columns(mvke_df, {'plant': 'Plant', 'werks': 'Plant', 'material': 'Material'})
    marm_df = _normalize_df_columns(marm_df, {'plant': 'Plant', 'werks': 'Plant', 'material': 'Material'})
    mard_df = _normalize_df_columns(mard_df, {'plant': 'Plant', 'werks': 'Plant', 'material': 'Material'})

    # ------------------------------------------------------------------
    # 2c) Hard-scope this engine to Palmela plants only (PT01/PT02), and
    # narrow further to a single plant if `plant_filter` asks for one.
    #
    # marc_file / mbew_file / zplm_marc_file may be *shared, global*
    # exports (e.g. when app.py runs the combined "All Sites" report and
    # hands the very same MARC file to both this engine and
    # run_tn_validation()). Without this guard, this engine would validate
    # TN01/TN02 rows too — using Palmela's generic rules instead of the
    # TN-exception-aware rules — producing spurious errors for Tunisia
    # materials that don't match (and therefore aren't removed by) the
    # app.py dedup step, since the error text differs from what the TN
    # engine legitimately produces for the same Material/Plant/Field.
    #
    # For a standalone Palmela-only run this is a no-op, since the
    # uploaded files already only contain PT01/PT02.
    # ------------------------------------------------------------------
    PALMELA_PLANTS = {"PT01", "PT02"}
    pt_scope = set(PALMELA_PLANTS)
    if plant_filter:
        pf = str(plant_filter).strip().upper()
        if pf in PALMELA_PLANTS:
            pt_scope = {pf}

    def _fpt(df, col='Plant'):
        if df is None or df.empty or col not in df.columns:
            return df
        return df[df[col].astype(str).str.strip().str.upper().isin(pt_scope)].copy()

    marc_df      = _fpt(marc_df)
    mbew_df      = _fpt(mbew_df)
    mard_df      = _fpt(mard_df) if mard_df is not None else None
    zplm_marc_df = _fpt(zplm_marc_df) if zplm_marc_df is not None else None

    # ------------------------------------------------------------------
    # 3) Pre-normalisation
    # ------------------------------------------------------------------
    marc_df = ensure_basic_data_k(marc_df)
    marc_df_full = marc_df.copy()

    # ------------------------------------------------------------------
    # 3b) Build sap_like_df (replaces the Item Browse Report / SAPUI5 Export)
    # ------------------------------------------------------------------
    _cb(18, "Building SAP-like dataframe from MARC + MARA + ZPLM MARC...")
    sap_like_df = build_sap_like_df(marc_df, mara_df, zplm_marc_df)

    # ✅ GLOBAL FILTER — ONLY FERT / HALB / ROH (Indirect Materials — ERSA,
    # HIBE, VERP — are handled by the separate Indirect Materials engine)
    sap_like_df = _filter_valid_material_types(sap_like_df, mara_df)
    advanced_planning_df = sap_like_df.copy()

    # Exclude Material Plant Status 99 rows from all validations except
    # Advanced Planning.
    exclude_mps99_pairs = _get_mps99_material_plant_pairs(marc_df_full)
    if exclude_mps99_pairs:
        sap_like_df = _exclude_material_plant_status_99(sap_like_df, exclude_mps99_pairs)
        marc_df = _exclude_material_plant_status_99(marc_df, exclude_mps99_pairs)
        if zplm_marc_df is not None:
            zplm_marc_df = _exclude_material_plant_status_99(
                zplm_marc_df,
                exclude_mps99_pairs,
            )
        if mvke_df is not None:
            mvke_df = _exclude_material_plant_status_99(
                mvke_df,
                exclude_mps99_pairs,
            )
        if mbew_df is not None:
            if 'Plant' in mbew_df.columns:
                mbew_df = _exclude_material_plant_status_99(
                    mbew_df,
                    exclude_mps99_pairs,
                    plant_col='Plant',
                )
            elif 'Valuation Area' in mbew_df.columns:
                mbew_df = _exclude_material_plant_status_99(
                    mbew_df,
                    exclude_mps99_pairs,
                    plant_col='Valuation Area',
                )

    marc_df     = _filter_valid_material_types(marc_df, mara_df)

    if mvke_df is not None:
        mvke_df = _filter_valid_material_types(mvke_df, mara_df)

    if mbew_df is not None:
        mbew_df = _filter_valid_material_types(mbew_df, mara_df)

    if mbew_df is not None and not mbew_df.empty:
        mbew_cols = [c for c in [
            'Material', 'Plant', 'Price Control', 'Price Unit',
            'Valuation Class', 'With Qty Structure'
        ] if c in mbew_df.columns]
        if 'Material' in mbew_cols:
            mbew_small = mbew_df[mbew_cols].copy()
            if 'Plant' in mbew_small.columns and 'Plant' in sap_like_df.columns:
                mbew_small = mbew_small.drop_duplicates(subset=['Material', 'Plant'], keep='first')
                sap_like_df = sap_like_df.merge(mbew_small, on=['Material', 'Plant'], how='left')
            else:
                mbew_small = mbew_small.drop_duplicates(subset=['Material'], keep='first')
                sap_like_df = sap_like_df.merge(mbew_small, on='Material', how='left')

    

    # ------------------------------------------------------------------
    # 4) Run all field validations
    # ------------------------------------------------------------------
    resultados = []

    # Mandatory fields
    for coluna in ['Material', 'Material Type', 'Description', 'Base UOM', 'Plant', 'Plant Type']:
        _cb(20, f"Validating mandatory field: {coluna}...")
        validar_campo_obrigatorio(sap_like_df, coluna, resultados)

    validations = [
        (22, lambda: validar_procurement_type(sap_like_df, resultados)),
        (24, lambda: validar_account_assignment_group(sap_like_df, resultados, mvke_df)),
        (26, lambda: validar_automatic_purchase_order(sap_like_df, resultados)),
        (28, lambda: validar_batch_management(sap_like_df, resultados)),
        (30, lambda: validar_consumption_mode(sap_like_df, resultados)),
        (32, lambda: validar_consumption_per(marc_df, mara_df, resultados)),
        (33, lambda: validar_division(sap_like_df, resultados)),
        (34, lambda: validar_export_import_group(sap_like_df, resultados)),
        (36, lambda: validar_general_item_category_group(sap_like_df, resultados)),
        (38, lambda: validar_gross_weight(sap_like_df, resultados)),
        (40, lambda: validar_item_category_group(sap_like_df, resultados, mvke_df)),
        (42, lambda: validar_jit_delivery_sched(sap_like_df, resultados)),
        (44, lambda: validar_loading_group(sap_like_df, resultados)),
        (46, lambda: validar_lot_sizing_procedure(sap_like_df, resultados)),
        (47, lambda: validar_per_lot_sz_plng_cal(advanced_planning_df, resultados, mara_df, matplsta_df)),
        (48, lambda: validar_material_group(sap_like_df, resultados)),
        (50, lambda: validar_mrp_type(sap_like_df, resultados, zplm_marc_df, csr_item_df, csr_head_df)),
        (51, lambda: validar_individual_collective_requirements_marc(marc_df, mara_df, resultados)),
        (52, lambda: validar_net_weight(sap_like_df, resultados)),
        (54, lambda: validar_period_type(sap_like_df, resultados)),
        (55, lambda: validar_storage_location_ext_procurement(marc_df, zplm_marc_df, mara_df, resultados)),
        (56, lambda: validar_planning_calendar(sap_like_df, resultados)),
        (57, lambda: validar_stock_determination_group_marc(marc_df, mara_df, resultados)),
        (58, lambda: validar_planning_time_fence(sap_like_df, resultados)),
        (59, lambda: validar_version_indicator_marc(marc_df, zplm_marc_df, mara_df, resultados)),
        (60, lambda: validar_price_control(sap_like_df, resultados)),
        (62, lambda: validar_price_unit(sap_like_df, resultados)),
        (63, lambda: validar_product_hierarchy(sap_like_df, resultados, mvke_df, mara_df)),
        (63, lambda: validar_product_hierarchy_mvke(sap_like_df, resultados, mvke_df)),
        (64, lambda: validar_production_supervisor_marc(marc_df, mara_df, resultados)),
        (66, lambda: validar_product_hierarchy_vs_enterproj(mara_df, zplm_mara_df, resultados)),
        (67, lambda: validar_packaging_material_type_mara(mara_df, resultados)),
        (67, lambda: validar_wh_material_group_mara(mara_df, resultados)),
        (68, lambda: validar_production_scheduling_profile(sap_like_df, resultados)),
        (69, lambda: validar_country_region_origin_fp(marc_df, zplm_marc_df, mara_df, resultados)),
        (69, lambda: validar_commodity_code_fp(marc_df, zplm_marc_df, mara_df, resultados)),
        (70, lambda: validar_profit_center(sap_like_df, resultados)),
        (71, lambda: validar_lifo_fifo_relevant_mbew(mbew_df, mara_df, resultados)),
        (72, lambda: validar_purchasing_group(sap_like_df, resultados)),
        (74, lambda: validar_repetitive_manufacturing(marc_df, zplm_marc_df, mara_df, resultados)),
        (76, lambda: validar_rem_profile(marc_df, zplm_marc_df, mara_df, resultados)),
        (78, lambda: validar_source_list_required_marc(marc_df, mara_df, resultados)),
        (80, lambda: validar_storage_location(sap_like_df, resultados)),
        (82, lambda: validar_valuation_class(sap_like_df, resultados)),
        (84, lambda: validar_with_qty_structure(sap_like_df, resultados)),
        (86, lambda: validar_enterproj_number(sap_like_df, resultados)),
        (87, lambda: validar_mrp_controller(sap_like_df, resultados, mard_df)),
        (88, lambda: validar_plant_specific_material_status(sap_like_df, resultados)),
        (89, lambda: validar_backflush(sap_like_df, resultados)),
        (90, lambda: validar_strategy_group(sap_like_df, resultados)),
        (91, lambda: validar_product_related_origin(sap_like_df, resultados)),
        (92, lambda: validar_availability_check(sap_like_df, resultados)),
        (93, lambda: validar_material_group_appoint(sap_like_df, resultados)),
        (94, lambda: validar_purchasing_group_dxx(sap_like_df, resultados)),
    ]

    for pct, fn in validations:
        _cb(pct, "Running field validations...")
        fn()

    # ------------------------------------------------------------------
    # 4B) Cross-file validations (non-SAPUI5 tables)
    # ------------------------------------------------------------------
    _cb(95, "Validating MARM Alternative UoM rules...")
    validar_marm_pc_ea_roh_halb(
        marm_df=marm_df,
        mara_df=mara_df,
        resultados=resultados
    )

    _cb(96, "Validating Advanced Planning fields (MARC)...")
    validar_advanced_planning_fields(
        marc_df=advanced_planning_df,
        mara_df=mara_df,
        resultados=resultados,
        matplsta_df=matplsta_df,
    )

    _cb(97, "Validating Neg. stocks in plant (MARC)...")
    neg_errors_early = validar_neg_stocks_in_plant(marc_df, mara_df)
    if neg_errors_early:
        resultados.extend(neg_errors_early)

    # ------------------------------------------------------------------
    # 5) Build errors DataFrame
    # ------------------------------------------------------------------
    errors_df = _coerce_df_for_display(pd.DataFrame(resultados))

    # ✅ Safety-net dedup: remove exact duplicates (same Material, Plant,
    # Field and Error). Different plants for the same material are fine
    # and expected (e.g. PT01 vs PT02) — only an exact repeat of the same
    # (Material, Plant, Field, Error) combination is removed here.
    _dedup_key = [c for c in ['Material', 'Plant', 'Field', 'Error'] if c in errors_df.columns]
    if len(_dedup_key) == 4 and not errors_df.empty:
        for _c in _dedup_key:
            errors_df[_c] = errors_df[_c].fillna('').astype(str).str.strip()
        errors_df = errors_df.drop_duplicates(subset=_dedup_key).reset_index(drop=True)

    # ------------------------------------------------------------
    # ✅ DETECT ERRORS NOT MAPPED IN ROLES (parity with TN engine)
    # ------------------------------------------------------------
    if roles_df is not None and not roles_df.empty:
        errors_df['Field_clean'] = errors_df['Field'].astype(str).str.strip()
        roles_set = set(roles_df['Field'].astype(str).str.strip())

        missing_role_df = errors_df[~errors_df['Field_clean'].isin(roles_set)].copy()

        if not missing_role_df.empty:
            print("\n--- ⚠️ ERRORS NOT MAPPED IN ROLES ---\n")
            for f in sorted(missing_role_df['Field_clean'].unique()):
                print(f)
            print("\n--------------------------------------\n")

        errors_df = errors_df.drop(columns=['Field_clean'], errors='ignore')
    else:
        missing_role_df = pd.DataFrame()

    # Enterproj enrichment (optional — needs previous report)
    if previous_report:
        errors_df = fill_plant_and_validate_enterproj_from_zplm(
            errors_df, latest_report=previous_report,
            zplm_marc_file=zplm_marc_file, zplm_mara_file=zplm_mara_file
        )

    # Additional enrichments
    if not errors_df.empty:
        errors_df = add_old_material_number_errors(mara_df, errors_df)

    # ------------------------------------------------------------------
    # 6) Post-processing
    # ------------------------------------------------------------------
    (
        errors_df, errors_filtered_df, views_missing_output_df,
        contagem_final_df, coverage_df, view_tabs_dict
    ) = postprocess_with_views_roles_and_classification(
        errors_df, mara_df, marc_df, roles_df, views_df, plant_filter=plant_filter
    )

    # Consolidate Missing_<View> tabs
    def _concat_missing_tabs(vtabs):
        cols = ['Material', 'Plant', 'Missing_View', 'View_Description', 'Fields_Expected']
        frames = [dfv[cols] for dfv in vtabs.values() if dfv is not None and not dfv.empty]
        if not frames:
            return pd.DataFrame(columns=cols)
        out = pd.concat(frames, ignore_index=True).drop_duplicates()
        out.sort_values(['Material', 'Plant', 'Missing_View'], inplace=True)
        return out

    view_is_missing_df = _concat_missing_tabs(view_tabs_dict)

    # ------------------------------------------------------------------
    # 7) Write Excel in-memory  (sheet order: Exec Summary first, raw data last)
    # ------------------------------------------------------------------
    _cb(100, "Generating final Excel report...")
    file_bytes = io.BytesIO()

    # Sort Adjusted Summary before writing (% of Total removed per user request)
    if not contagem_final_df.empty and 'Adjusted Nº of Errors' in contagem_final_df.columns:
        contagem_final_df = contagem_final_df.sort_values(
            'Adjusted Nº of Errors', ascending=False
        ).reset_index(drop=True)

    with pd.ExcelWriter(file_bytes, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, sheet_name="Executive Summary", index=False)

        base_row = 1
        if not contagem_final_df.empty:
            contagem_final_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=1)
            base_row = len(contagem_final_df) + 5
        if not coverage_df.empty:
            coverage_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=base_row)

        views_missing_output_df.to_excel(writer, sheet_name="Field With Error", index=False)
        errors_df.to_excel(writer, sheet_name="All Errors", index=False)
        view_is_missing_df.to_excel(writer, sheet_name="View is Missing", index=False)

        # Missing Role Mapping sheet (parity with TN engine)
        if not missing_role_df.empty:
            missing_role_df[['Field', 'Material', 'Plant', 'Error']].to_excel(
                writer, sheet_name="Missing Role Mapping", index=False
            )

        for view_code, df_view in view_tabs_dict.items():
            df_view.to_excel(writer, sheet_name=f"Missing_{view_code}"[:31], index=False)

        # Assignments sheet — enriched with owner/status from master.json
        try:
            _enriched = merge_assignments(errors_df, roles_df=roles_df)
            _asgn_df  = build_assignments_sheet(_enriched)
            _asgn_df.to_excel(writer, sheet_name="Assignments", index=False)
            save_open_assignments(_enriched)
        except Exception:
            pass

        marc_df.to_excel(writer, sheet_name="Original Data (MARC)", index=False)

    # Apply full professional styling
    file_bytes.seek(0)
    file_bytes = _style_sap_report(
        file_bytes,
        errors_df=errors_df,
        views_missing_output_df=views_missing_output_df,
        contagem_final_df=contagem_final_df,
        coverage_df=coverage_df,
        view_tabs_dict=view_tabs_dict,
        sap_df=marc_df,
        plant_filter=plant_filter,
    )

    # ------------------------------------------------------------------
    # 8) Build summary for Streamlit display
    # ------------------------------------------------------------------
    if 'Field' in errors_df.columns and not errors_df.empty:
        summary_df = (
            errors_df['Field'].value_counts()
            .rename_axis('Field').reset_index(name='Adjusted Nº of Errors')
        )
    else:
        summary_df = pd.DataFrame({"Field": [], "Adjusted Nº of Errors": []})

    return file_bytes, errors_df, summary_df


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
        from sap_validation_engine import list_run_logs as _lrl
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

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# =============================================================================
# 3) FP STATUS VALIDATION (Family Classification)
# =============================================================================

def _normalize_code_as_str_basic(x):
    """Converts numeric-like values (e.g. '55.0') to plain integers ('55')."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    try:
        return str(int(float(s)))
    except Exception:
        return s


def _pick_col_upper(df, candidates):
    """Returns the first matching column name from a list of candidates."""
    return next((c for c in candidates if c in df.columns), None)


def run_fp_status_validation(mara_file, marc_file, zplm_marc_file,
                              zplm_mara_file=None, plant_filter=None):
    """
    FP Family Classification:
    Groups FP materials by family (same base, different last character).
    Classifies each as 'Correct', 'Outdated', or 'Need verify'.
    Returns: (df_fp, kpis)
    """
    mara = load_xlsx_fast(mara_file, sheet_name=0)
    marc = load_xlsx_fast(marc_file, sheet_name=0)
    zplm = load_xlsx_fast(zplm_marc_file, sheet_name=0)

    if mara is None or marc is None or zplm is None:
        raise ValueError("Missing required input files.")

    for df in [mara, marc, zplm]:
        df.columns = df.columns.astype(str).str.strip().str.upper()

    col_mat_mara  = _pick_col_upper(mara, ["MATERIAL", "MATNR"])
    col_mat_marc  = _pick_col_upper(marc, ["MATERIAL", "MATNR"])
    col_plant_marc = _pick_col_upper(marc, ["PLANT", "WERKS"])
    col_psms      = _pick_col_upper(marc, ["PLANT-SPECIFIC MATERIAL STATUS",
                                            "PLANT SPECIFIC MATERIAL STATUS", "MTSTB"])
    col_mat_zplm  = _pick_col_upper(zplm, ["MATERIAL", "MATNR"])
    col_plant_zplm = _pick_col_upper(zplm, ["PLANT", "WERKS"])
    col_mptype    = _pick_col_upper(zplm, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])

    pf = str(plant_filter).strip().upper() if plant_filter else None
    mara_set = set(mara[col_mat_mara].dropna().astype(str).str.strip())

    # Filter FP rows from ZPLM
    df_fp = zplm[zplm[col_mptype].astype(str).str.strip().str.upper() == "FP"][
        [col_mat_zplm, col_plant_zplm, col_mptype]
    ].copy()
    df_fp[col_mat_zplm]   = df_fp[col_mat_zplm].astype(str).str.strip()
    df_fp[col_plant_zplm] = df_fp[col_plant_zplm].astype(str).str.strip().str.upper()
    df_fp = df_fp[df_fp[col_mat_zplm].isin(mara_set)]
    if pf:
        df_fp = df_fp[df_fp[col_plant_zplm] == pf]

    # Filter MARC for allowed PSMS
    marc_small = marc[[col_mat_marc, col_plant_marc, col_psms]].copy()
    marc_small[col_mat_marc]   = marc_small[col_mat_marc].astype(str).str.strip()
    marc_small[col_plant_marc] = marc_small[col_plant_marc].astype(str).str.strip().str.upper()
    marc_small[col_psms]       = marc_small[col_psms].apply(_normalize_code_as_str_basic)
    marc_small = marc_small[marc_small[col_psms].isin({"40", "45", "50", "55"})]
    if pf:
        marc_small = marc_small[marc_small[col_plant_marc] == pf]

    merged = df_fp.merge(
        marc_small,
        left_on=[col_mat_zplm, col_plant_zplm],
        right_on=[col_mat_marc, col_plant_marc],
        how="inner"
    ).rename(columns={
        col_mat_zplm: "Material", col_plant_zplm: "Plant",
        col_mptype: "Material Plant Type", col_psms: "PSMS"
    })[["Material", "Plant", "Material Plant Type", "PSMS"]]

    merged["Material"] = merged["Material"].astype(str).str.strip()
    merged = merged[~merged["Material"].str.match(r"^\d+$")]

    PSMS_NAME = {"40": "PREPROD", "45": "PROD", "50": "BALOUT", "55": "SERVICE"}
    merged["PSMS Name"] = merged["PSMS"].map(PSMS_NAME)

    # Keep only families with 2+ members
    merged["_base"] = merged["Material"].str[:-1]
    base_counts = merged["_base"].value_counts()
    merged = merged[merged["_base"].isin(set(base_counts[base_counts >= 2].index))].copy()

    def classify_family(g):
        g = g.copy()
        g["suffix"] = g["Material"].str[-1]
        g = g.sort_values("suffix")
        psms_set = set(g["PSMS"])

        if psms_set == {"55"} and len(g) >= 2:
            g["Status"] = "Need verify"
            return g.drop(columns=["suffix"])

        priority = ["45", "40", "55", "50"]
        correct_psms = next((p for p in priority if p in psms_set), None)
        g["Status"] = "Correct"

        has40, has45, has55 = "40" in psms_set, "45" in psms_set, "55" in psms_set

        if correct_psms in ("45", "40") and has55:
            g.loc[g["PSMS"] == "55", "Status"] = "Need verify"
        if correct_psms == "55" and (has40 or has45):
            g.loc[g["PSMS"].isin(["40", "45"]), "Status"] = "Need verify"
        if correct_psms == "45" and has40:
            g.loc[g["PSMS"] == "40", "Status"] = "Need verify"
        if correct_psms == "40" and has45:
            g.loc[g["PSMS"] == "45", "Status"] = "Need verify"

        # Balout (PSMS=50): only the LAST one (highest suffix) is Correct.
        # All others in the same family with status 50 are Outdated.
        dup50 = g[g["PSMS"] == "50"]
        if len(dup50) >= 2:
            # Mark all as Outdated first, then restore Correct for the last one
            g.loc[dup50.index, "Status"] = "Outdated"
            g.loc[dup50.index[-1], "Status"] = "Correct"

        return g.drop(columns=["suffix"])

    df = merged.groupby(["Plant", "_base"], group_keys=False).apply(classify_family)
    df = df.drop(columns=["_base"])

    kpis = {} if df.empty else {
        "materials": df["Material"].nunique(),
        "plants":    df["Plant"].nunique(),
        "psms":      df["PSMS"].nunique(),
        "outdated":  int((df["Status"] == "Outdated").sum()),
        "correct":   int((df["Status"] == "Correct").sum()),
    }

    return df, kpis


# =============================================================================
# 4) DATA CLEANUPS VALIDATION — Full Pipeline + Excel Export
# =============================================================================

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

    out2 = io.BytesIO()
    wb.save(out2)
    out2.seek(0)

    return out2, results, kpi_fp

# =============================================================================
# 5) ADVANCED PLANNING FIELD VALIDATIONS
# (formerly the standalone MRP Validation page — now integrated into
#  run_full_validation so these fields appear in the main SAP report)
# =============================================================================

def validar_advanced_planning_fields(
    marc_df,
    mara_df,
    resultados,
    matplsta_df=None
):
    """
    Validate Advanced Planning fields exclusively by Plant Type.

    PP rules:
    - Planning Status 1:
        - Heuristic must be empty.
        - PP Plng Procedure must be empty.
        - Proposed Strategy must be empty.
        - Advanced Planning must be OFF.
    - Planning Status 9 or blank:
        - Advanced Planning-related fields are not validated.

    PSMS rules:
    - PSMS 25/75:
        - Advanced Planning may be OFF.
        - Advanced Planning may only be ON when Planning Status is 9.

    FP/MP rules:
    - Planning Status 9 is not allowed unless PSMS is 25 or 75.
    - Remaining existing validation rules are preserved.
    """
    if marc_df is None or marc_df.empty:
        return

    def _pick(candidates):
        return next(
            (column for column in candidates if column in marc_df.columns),
            None
        )

    col_adv = _pick([
        'Advanced Planning'
    ])

    col_heur = _pick([
        'Heuristic',
        'Product Heuristic'
    ])

    col_pp = _pick([
        'PP Plng Procedure',
        'PP Planning Procedure'
    ])

    col_req = _pick([
        'Requirement Strategy',
        'Proposed Strategy'
    ])

    col_psms = _pick([
        'Plant-Specific Material Status',
        'Plant Specific Material Status'
    ])

    col_pt = _pick([
        'Plant Type',
        'Material Plant Type'
    ])

    required_columns = [
        col_adv,
        col_heur,
        col_pp,
        col_req,
        col_psms,
        col_pt
    ]

    if any(column is None for column in required_columns):
        return

    def _s(value):
        return '' if pd.isna(value) else str(value).strip()

    def is_true(value):
        return _s(value).upper() in {
            'TRUE',
            'X',
            '1',
            'YES',
            'Y'
        }

    def is_false(value):
        return _s(value).upper() in {
            'FALSE',
            '',
            '0',
            'NO',
            'N'
        }

    def add_error(idx, field, error, material, plant):
        resultados.append({
            'Line': idx + 2,
            'Field': field,
            'Error': error,
            'Material': material,
            'Plant': plant,
        })

    # -------------------------------------------------------------------------
    # Build Material + Plant -> Planning Status mapping
    # -------------------------------------------------------------------------
    planning_status_map = {}

    if matplsta_df is not None and not matplsta_df.empty:
        mps = matplsta_df.copy()
        mps.columns = mps.columns.astype(str).str.strip()

        required_matplsta_columns = {
            'Product Number',
            'Location',
            'Planning Status'
        }

        if required_matplsta_columns.issubset(mps.columns):
            for _, source_row in mps.iterrows():
                key = (
                    _s(source_row['Product Number']),
                    _s(source_row['Location']).upper()
                )

                planning_status_map[key] = normalize_code(
                    source_row['Planning Status'],
                    as_int=True
                )

    # -------------------------------------------------------------------------
    # Validate each MARC row
    # -------------------------------------------------------------------------
    for idx, row in marc_df.iterrows():
        material = _s(row.get('Material'))
        plant = _s(row.get('Plant')).upper()
        plant_type = _s(row.get(col_pt)).upper()

        if plant_type not in {'FP', 'MP', 'PP'}:
            continue

        advanced = row.get(col_adv)
        heuristic = _s(row.get(col_heur))

        procedure = normalize_code(
            row.get(col_pp),
            as_int=True
        )

        strategy = normalize_code(
            row.get(col_req),
            as_int=True
        )

        psms = normalize_code(
            row.get(col_psms),
            as_int=True
        )

        planning_status = planning_status_map.get(
            (material, plant),
            ''
        )

        # =====================================================================
        # PLANT TYPE PP
        # =====================================================================
        if plant_type == 'PP':

            # PP with Planning Status 9 or blank:
            # do not validate Advanced Planning, Heuristic,
            # PP Plng Procedure or Proposed Strategy.
            if planning_status in {'', '9'}:
                continue

            # PP with Planning Status 1:
            # dependent fields must be empty and Advanced Planning must be OFF.
            if planning_status == '1':

                if heuristic != '':
                    add_error(
                        idx,
                        'Heuristic',
                        (
                            "Must be empty for Plant Type 'PP' with "
                            f"Planning Status 1 (current: '{heuristic}')"
                        ),
                        material,
                        plant
                    )

                if procedure != '':
                    add_error(
                        idx,
                        'PP Plng Procedure',
                        (
                            "Must be empty for Plant Type 'PP' with "
                            f"Planning Status 1 (current: '{procedure}')"
                        ),
                        material,
                        plant
                    )

                if strategy != '':
                    add_error(
                        idx,
                        'Proposed Strategy',
                        (
                            "Must be empty for Plant Type 'PP' with "
                            f"Planning Status 1 (current: '{strategy}')"
                        ),
                        material,
                        plant
                    )

                if is_true(advanced):
                    add_error(
                        idx,
                        'Advanced Planning',
                        (
                            "Plant Type 'PP' with Planning Status 1 "
                            "must have Advanced Planning OFF"
                        ),
                        material,
                        plant
                    )

                continue

            # PP with any other Planning Status is outside
            # the currently defined validation rules.
            continue

        # =====================================================================
        # PLANT TYPES FP AND MP
        # =====================================================================

        # ---------------------------------------------------------------------
        # PSMS 25/75 precedence
        #
        # Advanced Planning may be:
        # - OFF, regardless of Planning Status.
        # - ON, but only when Planning Status is 9.
        # ---------------------------------------------------------------------
        if psms in {'25', '75'}:
            advanced_is_allowed = (
                is_false(advanced)
                or (
                    is_true(advanced)
                    and planning_status == '9'
                )
            )

            if not advanced_is_allowed:
                add_error(
                    idx,
                    'Advanced Planning',
                    (
                        "PSMS 25/75: Advanced Planning must be OFF, "
                        "or may only be ON when Planning Status is 9"
                    ),
                    material,
                    plant
                )

                continue

        # ---------------------------------------------------------------------
        # Existing restriction for MP in PT01 with Planning Status 1
        # ---------------------------------------------------------------------
        if (
            plant_type == 'MP'
            and plant == 'PT01'
            and planning_status == '1'
        ):
            add_error(
                idx,
                'Advanced Planning',
                (
                    "Planning Status 1 not allowed for "
                    "Plant Type 'MP' in PT01"
                ),
                material,
                plant
            )

            continue

        # ---------------------------------------------------------------------
        # MP in PT01 with Planning Status 9 and Advanced Planning ON is allowed.
        # ---------------------------------------------------------------------
        if (
            plant_type == 'MP'
            and plant == 'PT01'
            and planning_status == '9'
            and is_true(advanced)
        ):
            continue

        # ---------------------------------------------------------------------
        # Planning Status 9 is not allowed for FP/MP unless PSMS is 25 or 75.
        #
        # This condition prevents a conflict with the PSMS 25/75 rule above.
        # ---------------------------------------------------------------------
        if (
            plant_type in {'FP', 'MP'}
            and planning_status == '9'
            and psms not in {'25', '75'}
        ):
            add_error(
                idx,
                'Advanced Planning',
                (
                    "Planning Status 9 not allowed for Plant Type 'FP/MP' "
                    "unless PSMS is 25 or 75"
                ),
                material,
                plant
            )

            continue

        # ---------------------------------------------------------------------
        # Determine where Advanced Planning is mandatory
        # ---------------------------------------------------------------------
        required_advanced = (
            (
                plant == 'PT01'
                and plant_type == 'FP'
            )
            or (
                plant == 'PT02'
                and plant_type in {'FP', 'MP'}
            )
        )

        if required_advanced and not is_true(advanced):
            add_error(
                idx,
                'Advanced Planning',
                'Mandatory: must be TRUE/X',
                material,
                plant
            )

            continue

        # ---------------------------------------------------------------------
        # Do not validate dependent fields when Advanced Planning is OFF.
        # ---------------------------------------------------------------------
        if is_false(advanced):
            continue

        # ---------------------------------------------------------------------
        # FP/MP dependent fields
        # ---------------------------------------------------------------------
        if heuristic == '':
            add_error(
                idx,
                'Heuristic',
                'Mandatory',
                material,
                plant
            )

        if procedure != '4':
            add_error(
                idx,
                'PP Plng Procedure',
                "Should be '4'",
                material,
                plant
            )

        if strategy != '20':
            add_error(
                idx,
                'Proposed Strategy',
                "Should be '20'",
                material,
                plant
            )


# -----------------------------------------------------------------------------
# run_mrp_validation — kept for backward compatibility only.
# The validation logic now lives in validar_advanced_planning_fields()
# which is called directly from run_full_validation().
# -----------------------------------------------------------------------------
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


# =============================================================================
# 6) PLM / ATL WORKFLOW
# =============================================================================

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

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def run_plm_workflow_combined(auto_file, plm_file, timing_file,
                              workflowtask_file=None):
    """
    PLM / ATL Workflow Report.
    Source of truth for groups + Technical Status: C_WORKFLOWTASK (optional).
    Returns: (excel_bytes, df_pt01, df_pt02)

    If workflowtask_file is provided, groups (X marks) and Approval Status
    are derived from C_WORKFLOWTASK (Technical Status = READY / SELECTED).
    If not provided, falls back to the original approver-column logic.
    """
    # ------------------------------------------------------------------
    # 1) LOAD FILES
    # ------------------------------------------------------------------
    df_auto   = load_xlsx_fast(auto_file,   sheet_name=0)
    df_plm    = load_xlsx_fast(plm_file,    sheet_name=0)
    df_timing = load_xlsx_fast(timing_file, sheet_name=0)

    if any(df is None or df.empty for df in [df_auto, df_plm, df_timing]):
        raise ValueError("ATL files missing: Automotive / TimingLetter / Timing Letters")

    for col in ['Plant', 'Trigger ID', 'Creation Date']:
        if col not in df_auto.columns:
            raise ValueError(f"Column '{col}' missing in Automotive file.")
        if col not in df_plm.columns:
            raise ValueError(f"Column '{col}' missing in TimingLetter PLM file.")

    # Enterproj Number and Project Name are optional — present when already
    # linked to a CN in PLM. If the columns don't exist, we create them empty.
    for ep_col in ['Enterproj Number', 'Project Name']:
        if ep_col not in df_auto.columns:
            df_auto[ep_col] = ""
        if ep_col not in df_plm.columns:
            df_plm[ep_col] = ""

    # ── Build EP map HERE — before any filter or dedup ────────────────────────
    # The Automotive file is the only source of Enterproj Number / Project Name.
    # We build the lookup now, using ALL rows (including non-INAPPROVAL ones),
    # so that every CN that has these fields populated can later enrich the
    # TimingLetter rows — regardless of Approval Status.
    # Key = normalised CN string (same format used throughout the rest of the fn)
    _auto_ep_map: dict[str, dict] = {}
    for _, _ar in df_auto.iterrows():
        _key = str(_ar.get('Trigger ID', '')).strip()
        try:
            _key = str(int(float(_key)))
        except Exception:
            pass
        if not _key or _key in ('', 'nan'):
            continue
        _ep = str(_ar.get('Enterproj Number', '')).strip()
        _pn = str(_ar.get('Project Name', '')).strip()
        # Normalise numeric Enterproj values (e.g. '1234.0' → '1234')
        try:
            _ep = str(int(float(_ep))) if _ep not in ('', 'nan') else ''
        except Exception:
            pass
        if _ep not in ('', 'nan') or _pn not in ('', 'nan'):
            _auto_ep_map[_key] = {
                'Enterproj Number': _ep if _ep not in ('', 'nan') else '',
                'Project Name':     _pn if _pn not in ('', 'nan') else '',
            }
    if 'Timing Letter Number' not in df_timing.columns or 'Change Number' not in df_timing.columns:
        raise ValueError("Timing Letters file must contain 'Timing Letter Number' and 'Change Number'.")

    # ------------------------------------------------------------------
    # 2) MAP TL ↔ CN (active Timing Letters only)
    # ------------------------------------------------------------------
    EXCLUDED_TL_STATUS = {'Complete', 'Rejected'}
    status_col = next(
        (c for c in df_timing.columns
         if c.strip().lower() in ('timing letter status', 'status')), None
    )
    active_timing = (
        df_timing[~df_timing[status_col].astype(str).str.strip().isin(EXCLUDED_TL_STATUS)].copy()
        if status_col else df_timing.copy()
    )

    def _norm_id(x):
        """Normalize numeric IDs — avoids scientific notation (9.9e+11 → 990000000120)."""
        s = str(x).strip()
        try:
            return str(int(float(s)))
        except Exception:
            return s

    map_tl_to_cn = dict(zip(
        active_timing['Timing Letter Number'].apply(_norm_id),
        active_timing['Change Number'].apply(_norm_id)
    ))
    map_cn_to_tl = dict(zip(
        active_timing['Change Number'].apply(_norm_id),
        active_timing['Timing Letter Number'].apply(_norm_id)
    ))

    # Normalise PLM Trigger ID → CN number, and populate Timing Letter Number
    df_plm['_orig_trigger'] = df_plm['Trigger ID'].astype(str).str.strip()
    df_plm['Trigger ID'] = df_plm['_orig_trigger'].apply(
        lambda x: map_tl_to_cn.get(x, x)
    )
    df_plm['Timing Letter Number'] = df_plm['_orig_trigger'].apply(
        lambda x: _norm_id(x) if x in map_tl_to_cn else _norm_id(map_cn_to_tl.get(x, ""))
    )
    df_plm.drop(columns=['_orig_trigger'], inplace=True)

    # ------------------------------------------------------------------
    # 3) FILTER — keep only INAPPROVAL / REJECTED from source files
    # ------------------------------------------------------------------
    if 'Approval Status' in df_auto.columns:
        df_auto = df_auto[df_auto['Approval Status'].isin(['INAPPROVAL', 'REJECTED'])].copy()
    if 'Approval Status' in df_plm.columns:
        df_plm  = df_plm[df_plm['Approval Status'].isin(['INAPPROVAL', 'REJECTED'])].copy()
    # Note: REJECTED rows are filtered OUT of the final report after merging

    # ------------------------------------------------------------------
    # 4) DEDUPE — keep most recent per Trigger ID
    # ------------------------------------------------------------------
    for df in (df_auto, df_plm):
        df['Creation Date'] = pd.to_datetime(df['Creation Date'], errors='coerce')
        df.sort_values('Creation Date', ascending=False, inplace=True)
        df.drop_duplicates(subset=['Trigger ID'], keep='first', inplace=True)

    # ------------------------------------------------------------------
    # 5) INITIALISE GROUP COLUMNS
    # ------------------------------------------------------------------
    auto_groups = ["Pricing", "Purchasing", "MPL", "Customs",
                   "Sales", "Manufacturing", "Costing", "MDS"]
    tl_groups   = ["MPL-TL", "MDS-TL"]

    for g in auto_groups:
        df_auto[g] = ""
    for g in tl_groups:
        df_plm[g] = ""

    # ------------------------------------------------------------------
    # 6) APPLY GROUP X MARKS
    # ------------------------------------------------------------------
    if workflowtask_file is not None:
        # ── NEW LOGIC: derive groups from C_WORKFLOWTASK ──────────────
        df_wf = load_xlsx_fast(workflowtask_file, sheet_name=0)
        if df_wf is None or df_wf.empty:
            raise ValueError("C_WORKFLOWTASK file is empty or could not be loaded.")

        df_wf.columns = [str(c).strip() for c in df_wf.columns]

        # The CN/TL key is in 'Instance ID.1' (last Instance ID column)
        # 'Instance ID' (col 4) is NaN for workflow rows — use Instance ID.1
        instance_col_key = 'Instance ID.1'
        if instance_col_key not in df_wf.columns:
            # Fallback: pick last column named 'Instance ID' (case-insensitive)
            candidates = [c for c in df_wf.columns if 'Instance ID' in c.lower()]
            instance_col_key = candidates[-1] if candidates else None
        if instance_col_key is None:
            raise ValueError("C_WORKFLOWTASK: no Instance ID column found.")

        # Normalise columns
        df_wf[instance_col_key]   = df_wf[instance_col_key].astype(str).str.strip()
        df_wf['Technical Status'] = df_wf['Technical Status'].astype(str).str.upper().str.strip()

        # Work item text column — SAP exports it as 'Work item text' (lowercase 'item')
        wit_col = next((c for c in df_wf.columns
                        if c.lower().replace(' ', '') == 'workitemtext'), None)
        if wit_col is None:
            raise ValueError("C_WORKFLOWTASK: no 'Work item text' column found.")
        df_wf[wit_col] = df_wf[wit_col].astype(str).str.strip()

        # Keep only READY / SELECTED rows
        df_wf = df_wf[df_wf['Technical Status'].isin({'READY', 'SELECTED'})].copy()

        # Build workflow_map: {CN_or_TL_id: {group: "X", _status: ...}}
        # Key = Instance ID.1 value (e.g. 'CN00141782' or '990000000051')
        # For each CN/TL, collect all groups marked X across multiple rows
        workflow_map = {}
        for _, row in df_wf.iterrows():
            iid    = row[instance_col_key]
            text   = row[wit_col]

            if iid in ('nan', '', 'NaN'):
                continue

            entry = workflow_map.setdefault(iid, {})

            # Skip non-approval rows — only READY/SELECTED with group info matter
            if 'ready to approve' not in text.lower():
                continue

            # Extract group from Work item text.
            # Known patterns:
            #   TL:        "MPL-TL 990000000051/CN00142474 Ready to Approve"
            #   Standard:  "Pricing-Ready to Approve CN00141782"
            #   RT suffix: "MDS-RT4 400000000257 Ready to Approve"
            #   Two-word:  "Master Data-Ready to Approve CN00142919" → MDS
            import re as _re
            _GROUP_ALIASES = {'Master Data': 'MDS'}
            _first_word = text.split()[0].strip()
            if _first_word.endswith('-TL'):
                # Timing Letter group — always identifiable by first word
                entry[_first_word] = "X"
            else:
                _m = _re.match(r'^(.+?)(?:-Ready|-RT\d+|\s+Ready)', text, _re.IGNORECASE)
                _group_raw = _m.group(1).strip() if _m else _first_word
                entry[_GROUP_ALIASES.get(_group_raw, _group_raw)] = "X"

        # Apply workflow_map to each dataframe
        def _tln_key(raw):
            """Normalize a raw TLN cell value (float or string) to a clean string key."""
            import math
            if raw is None or (isinstance(raw, float) and math.isnan(raw)):
                return ''
            if isinstance(raw, (float, __builtins__.__class__)):
                try:
                    return str(int(raw))
                except Exception:
                    pass
            s = str(raw).strip()
            if s in ('', 'nan', 'NaN'):
                return ''
            try:
                return str(int(float(s)))
            except Exception:
                return s

        # Reset index after dedup to avoid stale index issues with df.at[]
        df_auto = df_auto.reset_index(drop=True)
        df_plm  = df_plm.reset_index(drop=True)

        def apply_wf_map_auto(df):
            """Automotive: look up by Trigger ID (CN number)."""
            for idx, row in df.iterrows():
                key = str(row['Trigger ID']).strip()
                wf  = workflow_map.get(key)
                if not wf:
                    continue
                for col, val in wf.items():
                    if col in df.columns:
                        df.at[idx, col] = val
            return df

        def apply_wf_map_tl(df):
            """TimingLetter: look up by Timing Letter Number (990* key).
            Never use Trigger ID — a TL's Trigger ID is a CN that also exists
            in workflow_map as Automotive, which would apply wrong groups."""
            if 'Timing Letter Number' not in df.columns:
                return df
            for idx, row in df.iterrows():
                tln = _tln_key(row['Timing Letter Number'])
                wf  = workflow_map.get(tln) if tln else None
                if not wf:
                    continue
                for col, val in wf.items():
                    if col in df.columns:
                        df.at[idx, col] = val
            return df

        df_auto = apply_wf_map_auto(df_auto)
        df_plm  = apply_wf_map_tl(df_plm)

    else:
        # ── FALLBACK: original approver-column logic ───────────────────
        categories_auto    = ["Pricing", "Purchasing", "MPL", "Customs",
                               "Sales", "Manufacturing", "Costing", "MDS"]
        approver_cols_auto = ['Approver1'] + [f'approver{i}' for i in range(2, 8)] + ['Approver8']

        for idx, row in df_auto.iterrows():
            for col in approver_cols_auto:
                if col in df_auto.columns:
                    clean = str(row[col]).split("(")[0].strip() if pd.notna(row[col]) else ""
                    if clean in categories_auto:
                        df_auto.at[idx, clean] = "X"

        for idx, row in df_plm.iterrows():
            for col in [c for c in df_plm.columns if "approver" in c.lower()]:
                val = str(row[col]) if pd.notna(row.get(col)) else ""
                if val.startswith('MPL'): df_plm.at[idx, 'MPL-TL'] = "X"
                if val.startswith('MDS'): df_plm.at[idx, 'MDS-TL'] = "X"

    # ------------------------------------------------------------------
    # 7) MERGE + FINAL COLUMNS
    # ------------------------------------------------------------------
    # Build Creation Date map from C_WORKFLOWTASK (earliest 'Created' per CN/TL)
    _wf_date_map: dict = {}
    if workflowtask_file is not None:
        try:
            _df_wf_dates = load_xlsx_fast(workflowtask_file, sheet_name=0)
            _df_wf_dates.columns = [str(c).strip() for c in _df_wf_dates.columns]
            _iid_col = 'Instance ID.1' if 'Instance ID.1' in _df_wf_dates.columns else None
            if _iid_col and 'Created' in _df_wf_dates.columns:
                def _norm_iid(x):
                    if pd.isna(x): return ''
                    s = str(x).strip()
                    try: return str(int(float(s)))
                    except Exception: return s
                _df_wf_dates[_iid_col] = _df_wf_dates[_iid_col].apply(_norm_iid)
                _df_wf_dates = _df_wf_dates[_df_wf_dates[_iid_col] != '']
                _df_wf_dates['Created'] = pd.to_datetime(_df_wf_dates['Created'], errors='coerce')
                _wf_date_map = _df_wf_dates.groupby(_iid_col)['Created'].min().to_dict()
        except Exception:
            _wf_date_map = {}

    def _get_creation_date(trigger_id, tln, fallback):
        """Return earliest workflow creation date for a CN or TL."""
        key_cn  = str(trigger_id).strip() if trigger_id else ''
        key_tln = str(tln).strip() if tln and str(tln).strip() not in ('', 'nan', '0') else ''
        dt = _wf_date_map.get(key_cn) or (_wf_date_map.get(key_tln) if key_tln else None)
        return dt if dt is not None else fallback

    df_auto['Source'] = "Automotive"
    df_plm['Source']  = "TimingLetter"

    group_cols = ["MPL-TL", "MDS-TL",
                  "Pricing", "Purchasing", "MPL", "Customs",
                  "Sales", "Manufacturing", "Costing", "MDS"]

    final_cols = [
        "Trigger ID", "Timing Letter Number", "Enterproj Number", "Project Name", "Plant",
        "Creation Date", "Approval Status",
    ] + group_cols + ["Aging"]

    # Ensure all columns exist in both dataframes
    for col in final_cols + ["Source"]:
        if col not in df_auto.columns: df_auto[col] = ""
        if col not in df_plm.columns:  df_plm[col]  = ""

    df_final = pd.concat([df_auto, df_plm], ignore_index=True)

    # Propagate TLN from TimingLetter rows to Automotive rows with same Trigger ID
    tln_map = (
        df_final[df_final['Source'] == 'TimingLetter']
        [['Trigger ID', 'Timing Letter Number']]
        .dropna(subset=['Timing Letter Number'])
        .drop_duplicates(subset=['Trigger ID'])
        .set_index('Trigger ID')['Timing Letter Number']
        .to_dict()
    )
    mask_auto = df_final['Source'] == 'Automotive'
    df_final.loc[mask_auto, 'Timing Letter Number'] = df_final.loc[mask_auto].apply(
        lambda row: tln_map.get(str(row['Trigger ID']).strip(), row['Timing Letter Number'])
        if (pd.isna(row['Timing Letter Number']) or str(row['Timing Letter Number']).strip() == '')
        else row['Timing Letter Number'],
        axis=1
    )

    # Remove REJECTED rows
    df_final = df_final[df_final['Approval Status'] != 'REJECTED'].copy()

    # ── Merge duplicate rows (same Trigger ID in Automotive + TimingLetter) ──
    # Combine group X marks from both sources into one row per CN/TL.
    merged_rows = []
    for trigger_id, grp in df_final.groupby('Trigger ID', sort=False):
        if len(grp) == 1:
            row_dict = grp.iloc[0].to_dict()
            # If this is a TimingLetter-only row, enrich from the Automotive map
            if str(row_dict.get('Source', '')).strip() == 'TimingLetter':
                ep_data = _auto_ep_map.get(str(trigger_id).strip(), {})
                for ep_col in ('Enterproj Number', 'Project Name'):
                    if not row_dict.get(ep_col) or \
                       str(row_dict[ep_col]).strip() in ('', 'nan'):
                        fallback = ep_data.get(ep_col, '')
                        if fallback not in ('', 'nan'):
                            row_dict[ep_col] = fallback
            merged_rows.append(row_dict)
            continue

        # Start from the Automotive row if present (for Approval Status / Plant)
        auto_rows = grp[grp['Source'] == 'Automotive']
        base_row  = auto_rows.iloc[0].to_dict() if not auto_rows.empty else grp.iloc[0].to_dict()

        # Merge ALL group X marks across all rows
        for _, other in grp.iterrows():
            for gc in group_cols:
                if str(other.get(gc, '')).strip().upper() == 'X':
                    base_row[gc] = 'X'
            # Fill TLN if missing in base
            if not base_row.get('Timing Letter Number') or \
               str(base_row['Timing Letter Number']).strip() in ('', 'nan', '0'):
                tln_other = str(other.get('Timing Letter Number', '')).strip()
                if tln_other not in ('', 'nan', '0'):
                    base_row['Timing Letter Number'] = tln_other
            # Fill Enterproj Number / Project Name if missing in base
            for ep_col in ('Enterproj Number', 'Project Name'):
                if not base_row.get(ep_col) or \
                   str(base_row[ep_col]).strip() in ('', 'nan'):
                    ep_other = str(other.get(ep_col, '')).strip()
                    if ep_other not in ('', 'nan'):
                        base_row[ep_col] = ep_other

        merged_rows.append(base_row)

    df_final = pd.DataFrame(merged_rows)

    # ── Universal EP enrichment — apply map to every row ─────────────────────
    # After all merging, do a final pass to ensure every CN that has
    # Enterproj Number / Project Name in the Automotive file gets it populated,
    # regardless of which code path the row went through above.
    def _enrich_ep(row):
        trigger = str(row.get('Trigger ID', '')).strip()
        try:
            trigger = str(int(float(trigger)))
        except Exception:
            pass
        ep_data = _auto_ep_map.get(trigger, {})
        for ep_col in ('Enterproj Number', 'Project Name'):
            current = str(row.get(ep_col, '')).strip()
            if current in ('', 'nan') and ep_data.get(ep_col, '') not in ('', 'nan'):
                row[ep_col] = ep_data[ep_col]
        return row

    if not df_final.empty:
        df_final = df_final.apply(_enrich_ep, axis=1)

    # Apply Creation Date from C_WORKFLOWTASK (earliest workflow date per CN/TL)
    df_final['Creation Date'] = df_final.apply(
        lambda row: _get_creation_date(
            row.get('Trigger ID'),
            row.get('Timing Letter Number'),
            pd.to_datetime(row.get('Creation Date'), errors='coerce')
        ),
        axis=1
    )
    df_final['Creation Date'] = pd.to_datetime(df_final['Creation Date'], errors='coerce')

    # Compute Aging and sort by Aging descending
    today = datetime.now()
    df_final['Aging'] = (today - df_final['Creation Date']).dt.days
    df_final.sort_values('Aging', ascending=False, inplace=True)
    df_final['Creation Date'] = df_final['Creation Date'].dt.strftime('%Y-%m-%d')

    # Final column selection — no Source column
    df_final = df_final[[c for c in final_cols if c in df_final.columns]]

    df_pt01  = df_final[df_final['Plant'] == 'PT01'].copy()
    df_pt02  = df_final[df_final['Plant'] == 'PT02'].copy()

    # ------------------------------------------------------------------
    # 8) EXCEL OUTPUT
    # ------------------------------------------------------------------
    xlsx_bytes = io.BytesIO()
    with pd.ExcelWriter(xlsx_bytes, engine="openpyxl") as writer:
        df_pt01.to_excel(writer, sheet_name="PT01", index=False)
        df_pt02.to_excel(writer, sheet_name="PT02", index=False)
    xlsx_bytes.seek(0)
    xlsx_bytes = _apply_conditional_formatting_in_memory(xlsx_bytes)

    return xlsx_bytes, df_pt01, df_pt02

# =============================================================================
# 7) COMPARE ADJUSTED REPORTS (Progress Tracking)
# =============================================================================

def _norm_str(x) -> str:
    """Normalises value to stripped string; NaN -> ''."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    return str(x).strip()


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


# =============================================================================
# 8) LOG PAYLOAD BUILDERS + PERSISTENCE
# =============================================================================

def _ts() -> str:
    return datetime.now().isoformat(timespec="seconds")

def _ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def build_sap_log_payload(*, user, plant_filter, mara_file, marc_file,
                           mvke_file, errors_df, summary_df, report_path=None):
    """Log payload for a SAP Validation run."""
    def _name(f, default):
        return getattr(f, "name", default) if f is not None else default

    total_errors    = int(errors_df.shape[0])            if errors_df is not None else 0
    total_fields    = int(errors_df['Field'].nunique())   if (errors_df is not None and 'Field' in errors_df.columns) else 0
    total_materials = int(errors_df['Material'].nunique()) if (errors_df is not None and 'Material' in errors_df.columns) else 0

    payload = {
        "run_type": "SAP_VALIDATION",
        "timestamp": _ts(),
        "user": user,
        "plant_filter": plant_filter or "ALL",
        "metrics": {"total_errors": total_errors, "fields_with_error": total_fields,
                    "materials_with_error": total_materials},
        "files": {"mara": _name(mara_file, "MARA.xlsx"),
                  "marc": _name(marc_file, "MARC.xlsx"),
                  "mvke": _name(mvke_file, "MVKE.xlsx")},
        "top_fields": summary_df.head(5).to_dict(orient="records") if summary_df is not None else [],
    }
    if report_path:
        payload["report_path"] = report_path
    return payload


def build_atl_log_payload(*, user, plant_scope, df_pt01, df_pt02):
    """Log payload for an ATL/PLM Workflow run."""
    def _tln_count(df):
        if df is None or df.empty or 'Timing Letter Number' not in df.columns:
            return 0
        return int(df['Timing Letter Number'].astype(str).str.strip().replace('', pd.NA).notna().sum())

    def _cn_count(df):
        if df is None or df.empty:
            return 0
        return int(df['Trigger ID'].nunique()) if 'Trigger ID' in df.columns else 0

    return {
        "run_type": "PLM_WORKFLOW_ATL",
        "timestamp": _ts(),
        "user": user,
        "plant_scope": plant_scope,
        "metrics": {
            "PT01_rows":       int(df_pt01.shape[0]) if df_pt01 is not None else 0,
            "PT02_rows":       int(df_pt02.shape[0]) if df_pt02 is not None else 0,
            "PT01_with_TLN":   _tln_count(df_pt01),
            "PT02_with_TLN":   _tln_count(df_pt02),
            "PT01_unique_CNs": _cn_count(df_pt01),
            "PT02_unique_CNs": _cn_count(df_pt02),
        },
    }


def build_progress_log_payload(*, user, current_name, previous_name, diff_dfs):
    """Log payload for a Progress Comparison run."""
    def _n(name):
        df = diff_dfs.get(name)
        return int(df.shape[0]) if df is not None else 0
    return {
        "run_type": "SAP_VALIDATION_PROGRESS",
        "timestamp": _ts(),
        "user": user,
        "files": {"current": current_name, "previous": previous_name},
        "metrics": {
            "new_errors":      _n("New Errors"),
            "resolved_errors": _n("Resolved Errors"),
            "changed_errors":  _n("Changed Errors"),
            "new_materials":   _n("New Materials"),
            "views_lost":      _n("Views Lost"),
        },
    }


def build_mrp_log_payload(*, user, plant_filter, mara_file, marc_file,
                            zplm_marc_file, errors_df, summary_df):
    """Log payload for an MRP Validation run."""
    def _name(f, default):
        return getattr(f, "name", default) if f is not None else default
    return {
        "run_type": "MRP_VALIDATION",
        "timestamp": _ts(),
        "user": user,
        "plant_filter": plant_filter or "ALL",
        "metrics": {
            "total_errors":       int(errors_df.shape[0])            if errors_df is not None else 0,
            "fields_with_error":  int(errors_df['Field'].nunique())   if (errors_df is not None and 'Field' in errors_df.columns) else 0,
            "materials_with_error": int(errors_df['Material'].nunique()) if (errors_df is not None and 'Material' in errors_df.columns) else 0,
        },
        "files": {"mara": _name(mara_file, "MARA.xlsx"),
                  "marc": _name(marc_file, "MARC.xlsx"),
                  "zplm_marc": _name(zplm_marc_file, "ZPLM MARC.xlsx")},
    }


def save_run_log(payload: dict, *, base_dir: str = ".history",
                  max_logs: int = 200) -> str | None:
    """
    Saves a log payload as a JSON file under .history/<run_type>/.

    FIX 12: Retention policy — keeps only the most recent `max_logs` files
    per run type. Older files are deleted automatically to prevent unbounded
    growth of the .history/ directory over time.
    """
    run_type = payload.get("run_type", "UNKNOWN")
    subdir   = {"SAP_VALIDATION": "sap", "PLM_WORKFLOW_ATL": "atl",
                 "SAP_VALIDATION_PROGRESS": "progress", "MRP_VALIDATION": "mrp",
                 "DATA_CLEANUPS": "dataclean",
                 "INDIRECT_MATERIALS": "indirect"}.get(run_type, "misc")

    out_dir = os.path.join(base_dir, subdir)
    _ensure_dir(out_dir)
    fname = f"{run_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    path  = os.path.join(out_dir, fname)

    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        return None

    # Retention: delete oldest files beyond max_logs
    try:
        all_logs = sorted(
            [os.path.join(out_dir, fn) for fn in os.listdir(out_dir)
             if fn.endswith(".json")],
            key=os.path.getmtime
        )
        for old_file in all_logs[:-max_logs]:
            try:
                os.remove(old_file)
            except Exception:
                pass
    except Exception:
        pass  # retention failure is non-critical

    return path

def list_run_logs(run_type: str, *, base_dir: str = ".history", limit: int = 25) -> list:
    """Lists the last `limit` log entries for a given run type."""
    subdir_map = {"sap": "sap", "atl": "atl", "progress": "progress",
                  "mrp": "mrp", "dataclean": "dataclean", "indirect": "indirect"}
    folder = os.path.join(base_dir, subdir_map.get(run_type, ""))
    if not os.path.isdir(folder):
        return []

    files = sorted(
        (os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(".json")),
        reverse=True
    )
    out = []
    for p in files[:limit]:
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
                data["_path"] = p
                out.append(data)
        except Exception:
            continue
    return out

def save_report_to_vault(excel_bytes: io.BytesIO, *, report_type: str = "sap",
                          plant_suffix: str = "ALL") -> str:
    """Saves the generated Excel to ./reports/<report_type>/<YYYY-MM>/ and returns the path."""
    _type_labels = {
        "sap":       "SAP_Validation",
        "atl":       "PLM_Workflow",
        "mrp":       "MRP_Validation",
        "dataclean": "Data_Cleanups",
        "progress":  "SAP_Progress",
        "indirect":  "Indirect_Materials",
    }
    label = _type_labels.get(report_type, report_type.upper())
    out_dir = os.path.join(".", "reports", report_type, datetime.now().strftime("%Y-%m"))
    os.makedirs(out_dir, exist_ok=True)
    fname = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{label}_{plant_suffix}.xlsx"
    fpath = os.path.join(out_dir, fname)
    excel_bytes.seek(0)
    with open(fpath, "wb") as f:
        f.write(excel_bytes.getvalue())
    return os.path.abspath(fpath)

# =============================================================================
# SAP VALIDATION ENGINE — PART 4: FINAL ASSEMBLY
# =============================================================================
# This file combines Parts 1–3 into a single, self-contained module.
#
# USAGE (development):
#   from sap_validation_engine_part1 import *
#   from sap_validation_engine_part2 import *
#   from sap_validation_engine_part3 import *
#   from sap_validation_engine_part4 import *
#
# USAGE (production — single file):
#   Concatenate Part 1 + Part 2 + Part 3 + Part 4 into sap_validation_engine.py
#
# SECTIONS:
#   A) Remaining utility helpers (_safe_name, _filter_materials_by_base_duplicates)
#   B) build_fp_log_payload
#   C) build_dc_log_payload
#   D) __all__ export list  (optional — useful for IDE discoverability)
# =============================================================================


# =============================================================================
# A) REMAINING UTILITY HELPERS
# =============================================================================

def _safe_name(uploaded, default_name: str) -> str:
    """
    Returns the filename of an uploaded file object (Streamlit UploadedFile
    or any object with a .name attribute).  Falls back to default_name.
    """
    return getattr(uploaded, "name", default_name) if uploaded is not None else default_name


def _coerce_df_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalises DataFrame column dtypes so PyArrow / Streamlit can serialise
    without errors.

    Problems addressed:
      - 'Line' column: mixed int/str object dtype  -> Int64 (nullable int)
      - Categorical columns -> converted to string to avoid Arrow edge-cases
      - Mixed-type object columns -> converted to string as last resort
    """
    if df is None or df.empty:
        return df
    df = df.copy()

    for col in df.columns:
        if col == 'Line':
            try:
                df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')
            except Exception:
                df[col] = df[col].astype(str)
            continue

        if hasattr(df[col], 'cat'):
            df[col] = df[col].astype(str)
            continue

        if df[col].dtype == object:
            types_present = {type(v) for v in df[col].dropna()}
            if len(types_present) > 1:
                df[col] = df[col].astype(str).replace('nan', '')
            elif types_present == {int} or types_present == {float}:
                try:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                except Exception:
                    pass

    return df


def _filter_materials_by_base_duplicates(
    df: pd.DataFrame,
    material_col: str = "Material",
) -> pd.DataFrame:
    """
    Keeps only rows whose Material 'base' (all chars except the last) appears
    at least twice in the current scope.  Used to identify FP version families.
    The original Material value is preserved unchanged.
    """
    if df is None or df.empty or material_col not in df.columns:
        return df

    tmp = df.copy()
    tmp["_base"] = tmp[material_col].astype(str).str.strip().str[:-1]
    base_counts  = tmp["_base"].value_counts(dropna=False)
    bases_keep   = set(base_counts[base_counts >= 2].index)
    return tmp[tmp["_base"].isin(bases_keep)].drop(columns=["_base"])


# =============================================================================
# A2) _classify_fp_versions  (simple last-char ordering, alternative to the
#     full classify_family logic in run_fp_status_validation)
# =============================================================================

def _classify_fp_versions(df: pd.DataFrame) -> pd.DataFrame:
    """
    Marks each FP material version within its family (all sharing the same base,
    i.e. all chars except the last).  Scope: (Plant, Base).

    Rules:
      - Sort by last character (A -> Z) within the group.
      - The lexicographically latest entry is 'Correct'; all others are 'Outdated'.

    Used as a lightweight alternative to the full PSMS-aware classify_family
    logic inside run_fp_status_validation.
    """
    if df.empty:
        df["Status"] = pd.Series(dtype=str)
        return df

    work = df.copy()
    work["Material"] = work["Material"].astype(str).str.strip()
    work = work[work["Material"].str.len() > 0].copy()

    work["Base"]     = work["Material"].str[:-1]
    work["LastChar"] = work["Material"].str[-1]
    work["Status"]   = "Outdated"

    for (_, _), grp in work.groupby(["Plant", "Base"]):
        if grp.empty:
            continue
        latest_idx = grp.sort_values("LastChar").index[-1]
        work.loc[latest_idx, "Status"] = "Correct"

    return work.drop(columns=["Base", "LastChar"])


# =============================================================================
# B) FP LOG PAYLOAD
# =============================================================================

def build_fp_log_payload(
    user: str,
    plant_filter: str | None,
    mara_file,
    marc_file,
    zplm_marc_file,
    zplm_mara_file,
    df_result: pd.DataFrame,
) -> dict:
    """
    Builds a structured log payload for an FP Status Validation run.

    Parameters
    ----------
    user          : username / identifier
    plant_filter  : 'PT01', 'PT02', or None (all plants)
    mara_file … zplm_mara_file : file objects (must have .name attribute, or None)
    df_result     : DataFrame returned by run_fp_status_validation
    """
    def _n(df):
        return 0 if (df is None or df.empty) else len(df)

    def _nuniq(df, col):
        return 0 if (df is None or df.empty or col not in df.columns) else int(df[col].nunique())

    def _count_status(df, status):
        if df is None or df.empty or "Status" not in df.columns:
            return 0
        return int((df["Status"] == status).sum())

    return {
        "run_type":     "FP_STATUS",
        "kind":         "fp",
        "timestamp":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "user":         user,
        "plant_filter": plant_filter or "ALL",
        "files": {
            "mara":      _safe_name(mara_file,      "MARA.xlsx"),
            "marc":      _safe_name(marc_file,      "MARC.xlsx"),
            "zplm_marc": _safe_name(zplm_marc_file, "ZPLM MARC.xlsx"),
            "zplm_mara": _safe_name(zplm_mara_file, "ZPLM MARA.xlsx"),
        },
        "metrics": {
            "materials": _nuniq(df_result, "Material"),
            "plants":    _nuniq(df_result, "Plant"),
            "outdated":  _count_status(df_result, "Outdated"),
            "correct":   _count_status(df_result, "Correct"),
            "need_verify": _count_status(df_result, "Need verify"),
            "total_rows":  _n(df_result),
        },
    }


# =============================================================================
# C) DATA CLEANUPS LOG PAYLOAD
# =============================================================================

def build_dc_log_payload(
    user: str,
    plant_filter: str | None,
    mara_file,
    marc_file,
    zplm_marc_file,
    zplm_mara_file,
    mbew_file,
    ekpo_file,
    mard_file,
    df_result: pd.DataFrame,
    kpis: dict,
) -> dict:
    """
    Builds a structured log payload for a Data Cleanups validation run.

    Parameters
    ----------
    user … mard_file : standard identifiers / file objects
    df_result        : combined results DataFrame (or None)
    kpis             : dict returned alongside df_result, e.g.
                       {"materials": int, "plants": int, "outdated": int, "correct": int}
    """
    return {
        "run_type":     "DATA_CLEANUPS",
        "timestamp":    datetime.now().isoformat(timespec="seconds"),
        "user":         user,
        "plant_filter": plant_filter or "ALL",
        "metrics": {
            "materials":   int(kpis.get("materials",   0)),
            "plants":      int(kpis.get("plants",      0)),
            "outdated":    int(kpis.get("outdated",    0)),
            "correct":     int(kpis.get("correct",     0)),
            "need_verify": int(kpis.get("need_verify", 0)),
        },
        "files": {
            "mara":      _safe_name(mara_file,      None),
            "marc":      _safe_name(marc_file,      None),
            "zplm_marc": _safe_name(zplm_marc_file, None),
            "zplm_mara": _safe_name(zplm_mara_file, None),
            "mbew":      _safe_name(mbew_file,      None),
            "ekpo":      _safe_name(ekpo_file,      None),
            "mard":      _safe_name(mard_file,      None),
        },
    }


# =============================================================================
# D) PUBLIC API  (__all__)
# =============================================================================

__all__ = [
    # --- Part 1: Loaders & Helpers ---
    "load_xlsx",
    "load_xlsx_fast",
    "get_rolesviews_bytesio",
    "normalize_code",
    "ensure_basic_data_k",
    "build_material_transfer_map",

    # --- Part 2: SAPUI5 field validations ---
    "validar_campo_obrigatorio",
    "validar_procurement_type",
    "validar_automatic_purchase_order",
    "validar_batch_management",
    "validar_consumption_mode",
    "validar_consumption_per",
    "validar_division",
    "validar_export_import_group",
    "validar_general_item_category_group",
    "validar_gross_weight",
    "validar_net_weight",
    "validar_jit_delivery_sched",
    "validar_loading_group",
    "validar_lot_sizing_procedure",
    "validar_per_lot_sz_plng_cal",
    "validar_material_group",
    "validar_material_group_appoint",
    "validar_purchasing_group_dxx",
    "validar_mrp_type",
    "validar_period_type",
    "validar_planning_calendar",
    "validar_planning_time_fence",
    "validar_price_control",
    "validar_price_unit",
    "validar_product_hierarchy",
    "validar_product_hierarchy_mvke",
    "validar_product_hierarchy_vs_enterproj",
    "validar_production_scheduling_profile",
    "validar_profit_center",
    "validar_purchasing_group",
    "validar_repetitive_manufacturing",
    "validar_repetitive_manufacturing_profile",
    "validar_source_list_required",
    "validar_storage_location",
    "validar_valuation_class",
    "validar_with_qty_structure",
    "validar_backflush",
    "validar_strategy_group",
    "validar_product_related_origin",
    "validar_availability_check",
    "validar_mrp_controller",
    "validar_plant_specific_material_status",
    "validar_enterproj_number",
    "validar_wh_material_group_mara",

    # --- Part 2: Cross-file validations ---
    "validar_account_assignment_group",
    "validar_item_category_group",

    # --- Part 2: Data Cleanup validations ---
    "validar_fp_pt01_pt02_status55",
    "validar_fp55_eop_vs_service_liability",
    "validar_fp55_service_liability_expired",
    "validar_fp50_eop_expired",
    "validar_fppp70_standard_price_mbew",
    "validar_ltb_fields",
    "validar_ekpo_delivery_completion",
    "validar_mard_total_stock",
    "validar_fp_existence_zplm_ta_ifs",

    # --- Part 2: Cross-existence validations ---
    "validar_mara_vs_zplm_mara_existence",
    "validar_zplm_mara_vs_mara_existence",
    "validar_marc_vs_zplm_marc_existence",
    "validar_zplm_marc_vs_marc_existence",

    # --- Part 2: Post-processing helpers ---
    "add_old_material_number_errors",
    "fill_plant_and_validate_enterproj_from_zplm",
    "validar_neg_stocks_in_plant",

    # --- Part 3: Orchestration ---
    "postprocess_with_views_roles_and_classification",
    "run_full_validation",
    "run_fp_status_validation",
    "run_data_cleanups_validation",
    "run_mrp_validation",
    "validar_advanced_planning_fields",
    "run_plm_workflow_combined",
    "compare_adjusted_reports",

    # --- Part 3: Styling helpers ---
    "_style_adjusted_summary_sheet",
    "_apply_conditional_formatting_in_memory",

    # --- Part 3 + 4: Log payloads & persistence ---
    "build_sap_log_payload",
    "build_atl_log_payload",
    "build_progress_log_payload",
    "build_mrp_log_payload",
    "build_fp_log_payload",
    "build_dc_log_payload",
    "save_run_log",
    "list_run_logs",
    "save_report_to_vault",

    # --- Part 4: Utility helpers ---
    "_safe_name",
    "_coerce_df_for_display",
    "_classify_fp_versions",
    "_filter_materials_by_base_duplicates",
    "_load_rolesviews_maps",
    "_extract_missing_views_from_workbook",
    "_normalize_code_as_str_basic",
    "_pick_col_upper",
    "_norm_str",
    "_ts",
    "_ensure_dir",
    # --- Assignment system ---
    "get_owner",
    "set_owner",
    "remove_owner",
    "list_all_assignments",
    "merge_assignments",
    "save_open_assignments",
    "build_assignments_sheet",
    "style_assignments_sheet",
    "_load_master",
    "_save_master",
]

# =============================================================================
# ASSIGNMENT SYSTEM — "Assign once, persist forever"
# =============================================================================
#
# Architecture:
#   .history/assignments/master.json  — permanent field→owner map
#   Sheet "Assignments" in every SAP Excel report
#
# master.json structure:
# {
#   "field_assignments": {
#     "MRP Type": {"owner": "João Silva", "team": "MRP", "since": "2026-01-15"}
#   },
#   "plant_overrides": {
#     "PT02": { "MRP Type": {"owner": "Carlos Mota"} }
#   }
# }
# =============================================================================

_ASSIGNMENTS_PATH = os.path.join(".history", "assignments", "master.json")


def _load_master() -> dict:
    """Load the assignment master file. Returns empty structure if missing."""
    if not os.path.exists(_ASSIGNMENTS_PATH):
        return {"field_assignments": {}, "plant_overrides": {}}
    try:
        with open(_ASSIGNMENTS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("field_assignments", {})
        data.setdefault("plant_overrides", {})
        return data
    except Exception:
        return {"field_assignments": {}, "plant_overrides": {}}


def _save_master(master: dict) -> None:
    """Persist the assignment master file atomically."""
    os.makedirs(os.path.dirname(_ASSIGNMENTS_PATH), exist_ok=True)
    tmp = _ASSIGNMENTS_PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(master, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _ASSIGNMENTS_PATH)
    except Exception:
        pass


def _get_assignment_info(field: str, plant: str = "", master: dict | None = None) -> dict:
    """Returns the resolved assignment info for a field/plant combination."""
    if master is None:
        master = _load_master()
    field = (field or "").strip()
    plant = (plant or "").strip().upper()

    if not field:
        return {"owner": "", "team": "", "level": "Unassigned", "plant": ""}

    override = master.get("plant_overrides", {}).get(plant, {}).get(field)
    if isinstance(override, dict) and override.get("owner"):
        return {
            "owner": override["owner"],
            "team": override.get("team", ""),
            "level": "Plant Override",
            "plant": plant,
        }

    default = master.get("field_assignments", {}).get(field)
    if isinstance(default, dict) and default.get("owner"):
        return {
            "owner": default["owner"],
            "team": default.get("team", ""),
            "level": "Default",
            "plant": "",
        }

    return {"owner": "", "team": "", "level": "Unassigned", "plant": ""}


def get_owner(field: str, plant: str = "", master: dict | None = None) -> str:
    """
    Returns the owner for a (field, plant) combination.
    Plant-level override takes priority over field-level default.
    Returns "" if not assigned.
    """
    return _get_assignment_info(field, plant, master=master)["owner"]


def set_owner(field: str, owner: str, *,
              plant: str = "",
              team: str = "",
              master: dict | None = None) -> dict:

    if master is None:
        master = _load_master()

    field = (field or "").strip()
    plant = (plant or "").strip().upper()
    owner = (owner or "").strip()
    now   = datetime.now().strftime("%Y-%m-%d")

    if not field:
        raise ValueError("Field name is required to set an assignment.")

    # ✅ FIX: reload after remove
    if not owner:
        remove_owner(field, plant=plant)
        return _load_master()

    if plant:
        prev = master.setdefault("plant_overrides", {}).setdefault(plant, {}).get(field, {})

        # ✅ PATCH
        if team is not None:
            team_val = team
        else:
            team_val = prev.get("team", "")

        master["plant_overrides"][plant][field] = {
            "owner": owner,
            "team": team_val,
            "since": now
        }

    else:
        prev = master.setdefault("field_assignments", {}).get(field, {})

        # ✅ PATCH
        if team is not None:
            team_val = team
        else:
            team_val = prev.get("team", "")

        master["field_assignments"][field] = {
            "owner": owner,
            "team": team_val,
            "since": now
        }

    _save_master(master)
    return master


def remove_owner(field: str, *, plant: str = "") -> None:
    master = _load_master()
    field = (field or "").strip()
    plant = (plant or "").strip().upper()

    # ✅ FIX: proteção
    if not field:
        return

    if plant:
        plant_dict = master.get("plant_overrides", {}).get(plant, {})
        if isinstance(plant_dict, dict):
            plant_dict.pop(field, None)
            if not plant_dict:
                master.get("plant_overrides", {}).pop(plant, None)
    else:
        master.get("field_assignments", {}).pop(field, None)

    if not master.get("plant_overrides"):
        master["plant_overrides"] = {}

    _save_master(master)


def list_all_assignments() -> pd.DataFrame:
    """
    Returns a flat DataFrame of all current assignments:
    Field | Plant | Owner | Team | Since | Level
    """
    master = _load_master()
    rows = []
    for field, info in master.get("field_assignments", {}).items():
        rows.append({
            "Field":  field,
            "Plant":  "(all plants)",
            "Owner":  info.get("owner", ""),
            "Team":   info.get("team", ""),
            "Since":  info.get("since", ""),
            "Level":  "Default",
        })
    for plant, fields in master.get("plant_overrides", {}).items():
        for field, info in fields.items():
            rows.append({
                "Field":  field,
                "Plant":  plant,
                "Owner":  info.get("owner", ""),
                "Team":   info.get("team", ""),
                "Since":  info.get("since", ""),
                "Level":  "Plant Override",
            })
    return pd.DataFrame(rows, columns=["Field", "Plant", "Owner", "Team", "Since", "Level"])


def _build_field_to_role_map(roles_df: pd.DataFrame | None) -> dict:
    """
    Builds Field -> Role mapping from Roles DataFrame.
    Returns dict: {field_name: role_name}
    """
    if roles_df is None or roles_df.empty:
        return {}
    
    field_to_role = {}
    if 'Field' in roles_df.columns and 'Role' in roles_df.columns:
        roles_copy = roles_df.copy()
        roles_copy['Field'] = roles_copy['Field'].astype(str).str.strip()
        roles_copy['Role'] = roles_copy['Role'].astype(str).str.strip()
        # If multiple roles per field, take the first one
        field_to_role = roles_copy.drop_duplicates(subset=['Field']).set_index('Field')['Role'].to_dict()
    
    return field_to_role


def merge_assignments(errors_df: pd.DataFrame,
                      run_date: str | None = None,
                      roles_df: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    Enriches errors_df with assignment data from master.json.
    """

    if errors_df is None or errors_df.empty:
        return errors_df

    # ✅ AUTO-LOAD ROLES (PATCH)
    if roles_df is None:
        try:
            from utils.fast_loader import load_xlsx_fast

            _rv = load_xlsx_fast(get_rolesviews_bytesio(), sheet_name=None)
            roles_df = _rv.get("Roles")
        except Exception:
            roles_df = None

    master    = _load_master()
    today_str = run_date or datetime.now().strftime("%Y-%m-%d")
    today     = datetime.strptime(today_str, "%Y-%m-%d")

    field_to_role = _build_field_to_role_map(roles_df)

    # Load history
    open_path = os.path.join(".history", "assignments", "open_items.json")
    open_map: dict[tuple, dict] = {}

    if os.path.exists(open_path):
        try:
            with open(open_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            for k, v in raw.items():
                parts = k.split("||")
                if len(parts) == 3:
                    open_map[tuple(parts)] = v
        except Exception:
            pass

    df = errors_df.copy()

    owners, teams, statuses, days_open_list, notes_list, levels = [], [], [], [], [], []

    for _, row in df.iterrows():
        field    = str(row.get("Field", "")).strip()
        plant    = str(row.get("Plant", "")).strip().upper()
        material = str(row.get("Material", "")).strip()

        assignment = _get_assignment_info(field, plant, master=master)

        owner = assignment["owner"]
        team  = assignment["team"]
        level = assignment["level"]

        # ✅ FINAL TEAM LOGIC
        assignment_team = (team or "").strip()
        role_team       = (field_to_role.get(field, "") or "").strip()

        if assignment_team:
            final_team = assignment_team
        elif role_team:
            final_team = role_team
        else:
            final_team = ""

        # Carry forward
        item_key = (material, plant, field)
        prev     = open_map.get(item_key, {})

        notes = prev.get("notes", "")
        first_seen_str = prev.get("first_seen", today_str)

        try:
            first_seen = datetime.strptime(first_seen_str, "%Y-%m-%d")
            d_open = (today - first_seen).days
        except Exception:
            d_open = 0

        OVERDUE_DAYS = 7

        if not owner:
            status = "Unassigned"
        elif d_open == 0:
            status = "New"
        elif d_open > OVERDUE_DAYS:
            status = "Overdue"
        else:
            status = "Open"

        owners.append(owner)
        teams.append(final_team)
        statuses.append(status)
        days_open_list.append(d_open)
        notes_list.append(notes)
        levels.append(level)

    df["Owner"]     = owners
    df["Team"]      = teams
    df["Level"]     = levels
    df["Status"]    = statuses
    df["Days_Open"] = days_open_list
    df["Notes"]     = notes_list

    return df


def save_open_assignments(enriched_df: pd.DataFrame) -> None:
    """
    Persists the current open assignment items to
    .history/assignments/open_items.json for carry-forward in the next run.

    Only saves rows that have an Owner (assigned items).
    Preserves first_seen date from existing entries.
    Saves: owner, team, status, days_open, notes, first_seen, last_seen
    """
    open_path = os.path.join(".history", "assignments", "open_items.json")
    os.makedirs(os.path.dirname(open_path), exist_ok=True)

    # Load existing
    existing: dict = {}
    if os.path.exists(open_path):
        try:
            with open(open_path, "r", encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = {}

    today_str = datetime.now().strftime("%Y-%m-%d")
    updated: dict = {}

    for _, row in enriched_df.iterrows():
        owner = str(row.get("Owner", "")).strip()
        if not owner:
            continue
        material = str(row.get("Material", "")).strip()
        plant    = str(row.get("Plant",    "")).strip().upper()
        field    = str(row.get("Field",    "")).strip()
        key      = f"{material}||{plant}||{field}"

        prev_first_seen = existing.get(key, {}).get("first_seen", today_str)
        updated[key] = {
            "owner":      owner,
            "team":       str(row.get("Team",     "")).strip(),
            "notes":      str(row.get("Notes",     "")).strip(),
            "status":     str(row.get("Status",    "")).strip(),
            "days_open":  int(row.get("Days_Open", 0)),
            "first_seen": prev_first_seen,
            "last_seen":  today_str,
        }

    try:
        tmp = open_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(updated, f, ensure_ascii=False, indent=2)
        os.replace(tmp, open_path)
    except Exception:
        pass


def build_assignments_sheet(enriched_df: pd.DataFrame) -> pd.DataFrame:
    """
    Builds the Assignments sheet DataFrame for Excel export.
    Columns: Material | Plant | Field | Error | Owner | Team |
             Level | Status | Days_Open | Notes
    Sorted: Overdue first, then Open, then New, then Unassigned.
    """
    cols = ["Material", "Plant", "Field", "Error",
            "Owner", "Team", "Level", "Status", "Days_Open", "Notes"]

    present = [c for c in cols if c in enriched_df.columns]
    df = enriched_df[present].copy()

    # Add missing cols as empty
    for c in cols:
        if c not in df.columns:
            df[c] = ""

    status_order = {"Overdue": 0, "Open": 1, "New": 2, "Unassigned": 3}
    df["_sort"] = df["Status"].map(status_order).fillna(9)
    df = df.sort_values(["_sort", "Days_Open"], ascending=[True, False]).drop(columns=["_sort"])

    return df[cols].reset_index(drop=True)


def style_assignments_sheet(ws, df: pd.DataFrame) -> None:
    """
    Applies colour-coding to the Assignments sheet in an openpyxl workbook.
    Open = orange, New = blue, Unassigned = grey.
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    STATUS_FILLS = {
        "Overdue":    PatternFill(start_color="FECACA", fill_type="solid"),  # 🔴
        "Open":       PatternFill(start_color="FFE5B4", fill_type="solid"),  # 🟠
        "New":        PatternFill(start_color="DBEAFE", fill_type="solid"),  # 🔵
        "Unassigned": PatternFill(start_color="F3F4F6", fill_type="solid"),  # ⚪
    }
    header_fill  = PatternFill(start_color="1F4E79", fill_type="solid")
    header_font  = Font(bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin  = Side(style="thin", color="B8CCE4")
    bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 26

    # Header
    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = header_align
        cell.border    = bdr

    # Find Status column index
    status_col = next(
        (i for i, cell in enumerate(ws[1], start=1) if str(cell.value) == "Status"),
        None
    )

    for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        status_val = ""
        if status_col:
            status_val = str(ws.cell(row=r_idx, column=status_col).value or "")
        fill = STATUS_FILLS.get(status_val, PatternFill())
        for cell in row:
            cell.border = bdr
            if fill:
                cell.fill = fill

    # Auto-width
    for col_idx, col in enumerate(ws.columns, start=1):
        max_len = max((len(str(cell.value or "")) for cell in col), default=8)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 50)


# =============================================================================
# ERROR HISTORY & TREND TRACKING
# =============================================================================

_ERROR_HISTORY_PATH = os.path.join(".history", "assignments", "error_history.json")


def _load_error_history() -> dict:
    """Load error history file. Returns empty dict if missing."""
    if not os.path.exists(_ERROR_HISTORY_PATH):
        return {}
    try:
        with open(_ERROR_HISTORY_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_error_history(history: dict) -> None:
    """Persist error history file atomically."""
    os.makedirs(os.path.dirname(_ERROR_HISTORY_PATH), exist_ok=True)
    tmp = _ERROR_HISTORY_PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _ERROR_HISTORY_PATH)
    except Exception:
        pass


def track_error_count(errors_df: pd.DataFrame, run_date: str | None = None) -> None:
    """
    Records error count for each (Material, Plant, Field) combination.
    Determines status: "appeared" | "improving" | "stable" | "worsening" | "resolved" | "recurring"
    
    Should be called at end of validation run, BEFORE save_open_assignments().
    
    Args:
        errors_df: Enriched errors DataFrame with Material, Plant, Field columns
        run_date: ISO date string; defaults to today
    """
    if errors_df is None or errors_df.empty:
        return
    
    today_str = run_date or datetime.now().strftime("%Y-%m-%d")
    history = _load_error_history()
    
    # Count errors per (Material, Plant, Field)
    error_counts = {}
    for _, row in errors_df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant = str(row.get("Plant", "")).strip().upper()
        field = str(row.get("Field", "")).strip()
        key = f"{material}||{plant}||{field}"
        error_counts[key] = error_counts.get(key, 0) + 1
    
    # Update history
    for key, current_count in error_counts.items():
        if key not in history:
            history[key] = {"runs": []}
        
        runs = history[key]["runs"]
        
        # Determine status
        if not runs:
            status = "appeared"
        else:
            last_run = runs[-1]
            prev_count = last_run.get("error_count", 0)
            
            if prev_count == 0:
                status = "recurring"
            elif current_count < prev_count:
                status = "improving"
            elif current_count > prev_count:
                status = "worsening"
            else:
                status = "stable"
        
        # Add run record (avoid duplicates on same date)
        if not runs or runs[-1].get("date") != today_str:
            runs.append({
                "date": today_str,
                "error_count": current_count,
                "status": status
            })
    
    # Remove keys that no longer have errors (they were resolved)
    for key in list(history.keys()):
        if key not in error_counts:
            # This field had errors before, but doesn't now
            runs = history[key]["runs"]
            if runs and runs[-1].get("error_count", 0) > 0:
                # Last run had errors; this run has none
                if not runs or runs[-1].get("date") != today_str:
                    runs.append({
                        "date": today_str,
                        "error_count": 0,
                        "status": "resolved"
                    })
    
    _save_error_history(history)


def get_error_trend(material: str, plant: str, field: str) -> dict:
    """
    Returns trend analysis for a specific (Material, Plant, Field).
    
    Returns dict with:
    - first_seen: ISO date of first error
    - last_seen: ISO date of last event
    - current_count: Current error count (0 if resolved)
    - peak_count: Highest count ever recorded
    - status: Current status (appeared/improving/stable/worsening/resolved/recurring)
    - trend: "improving" | "worsening" | "stable" | "volatile"
    - resolution_rate: Percentage solved from peak
    - history: List of all runs
    """
    history = _load_error_history()
    key = f"{material}||{plant}||{field}"
    
    if key not in history:
        return {
            "material": material,
            "plant": plant,
            "field": field,
            "first_seen": None,
            "last_seen": None,
            "current_count": 0,
            "peak_count": 0,
            "status": "never_tracked",
            "trend": "n/a",
            "resolution_rate": "0%",
            "history": []
        }
    
    runs = history[key]["runs"]
    if not runs:
        return {
            "material": material,
            "plant": plant,
            "field": field,
            "first_seen": None,
            "last_seen": None,
            "current_count": 0,
            "peak_count": 0,
            "status": "never_tracked",
            "trend": "n/a",
            "resolution_rate": "0%",
            "history": []
        }
    
    # Calculate metrics
    first_seen = runs[0]["date"]
    last_seen = runs[-1]["date"]
    current_count = runs[-1]["error_count"]
    peak_count = max(r["error_count"] for r in runs)
    status = runs[-1]["status"]
    
    # Determine trend
    if len(runs) < 2:
        trend = "stable"
    else:
        recent = runs[-3:] if len(runs) >= 3 else runs
        counts = [r["error_count"] for r in recent]
        if counts[-1] < counts[0]:
            trend = "improving"
        elif counts[-1] > counts[0]:
            trend = "worsening"
        else:
            trend = "stable"
        
        # Check for volatile (up/down/up or similar)
        if len(counts) >= 3:
            diffs = [counts[i+1] - counts[i] for i in range(len(counts)-1)]
            if any(d > 0 for d in diffs) and any(d < 0 for d in diffs):
                trend = "volatile"
    
    # Resolution rate
    if peak_count > 0:
        resolution_rate = f"{int(100 * (peak_count - current_count) / peak_count)}%"
    else:
        resolution_rate = "0%"
    
    return {
        "material": material,
        "plant": plant,
        "field": field,
        "first_seen": first_seen,
        "last_seen": last_seen,
        "current_count": current_count,
        "peak_count": peak_count,
        "status": status,
        "trend": trend,
        "resolution_rate": resolution_rate,
        "history": runs
    }


def get_owner_metrics(owner: str) -> dict:
    """
    Returns aggregated metrics for all fields assigned to an owner.
    """

    if not owner or not owner.strip():
        return {}

    master = _load_master()
    history = _load_error_history()

    # ─────────────────────────────────────────────
    # FIND ALL ASSIGNED FIELDS
    # ─────────────────────────────────────────────
    assigned_fields = set()

    for field, info in master.get("field_assignments", {}).items():
        if info.get("owner") == owner:
            assigned_fields.add((field, ""))

    for plant, fields in master.get("plant_overrides", {}).items():
        for field, info in fields.items():
            if info.get("owner") == owner:
                assigned_fields.add((field, plant))

    # ─────────────────────────────────────────────
    # METRICS
    # ─────────────────────────────────────────────
    by_field = []
    total_active = 0

    improving = worsening = stable = resolved = recurring = 0

    # ✅ PATCH — evitar duplicação
    seen_keys = set()

    for field, plant in assigned_fields:

        for key, hist_data in history.items():

            if key in seen_keys:
                continue

            parts = key.split("||")
            if len(parts) != 3:
                continue

            hist_material, hist_plant, hist_field = parts

            # Match field e plant
            if hist_field != field:
                continue

            if plant and hist_plant != plant:
                continue

            runs = hist_data.get("runs", [])
            if not runs:
                continue

            seen_keys.add(key)

            last_run = runs[-1]
            current_count = int(last_run.get("error_count", 0))
            status = last_run.get("status", "")

            # ✅ agregação
            total_active += current_count

            if status == "improving":
                improving += 1
            elif status == "worsening":
                worsening += 1
            elif status == "stable":
                stable += 1
            elif status == "resolved":
                resolved += 1
            elif status == "recurring":
                recurring += 1

            by_field.append({
                "field": hist_field,
                "material": hist_material,
                "plant": hist_plant,
                "current_count": current_count,
                "peak_count": max(r.get("error_count", 0) for r in runs),
                "status": status,
                "first_seen": runs[0].get("date") if runs else None,
                "last_seen": runs[-1].get("date") if runs else None,
            })

    # ─────────────────────────────────────────────
    # OUTPUT
    # ─────────────────────────────────────────────
    return {
        "owner": owner,
        "total_active_errors": total_active,
        "fields_count": len(assigned_fields),
        "improving": improving,
        "worsening": worsening,
        "stable": stable,
        "resolved": resolved,
        "recurring": recurring,
        "by_field": sorted(by_field, key=lambda x: -x["current_count"])
    }

def get_error_summary_df(run_date: str | None = None) -> pd.DataFrame:
    """
    Returns a summary DataFrame of all tracked errors with trend info.
    Useful for dashboards.
    
    Columns: Material | Plant | Field | Current | Peak | Status | Trend | 
             First_Seen | Last_Seen | Days_Today
    """
    history = _load_error_history()
    today_str = run_date or datetime.now().strftime("%Y-%m-%d")
    today = datetime.strptime(today_str, "%Y-%m-%d")
    
    rows = []
    for key, hist_data in history.items():
        parts = key.split("||")
        if len(parts) == 3:
            material, plant, field = parts
            runs = hist_data.get("runs", [])
            if runs:
                last_run = runs[-1]
                first_run = runs[0]
                peak_count = max(r["error_count"] for r in runs)
                
                try:
                    days_active = (today - datetime.strptime(first_run["date"], "%Y-%m-%d")).days
                except Exception:
                    days_active = 0
                
                # Determine trend
                if len(runs) < 2:
                    trend = "stable"
                else:
                    recent = runs[-3:] if len(runs) >= 3 else runs
                    counts = [r["error_count"] for r in recent]
                    if counts[-1] < counts[0]:
                        trend = "improving"
                    elif counts[-1] > counts[0]:
                        trend = "worsening"
                    else:
                        trend = "stable"
                
                rows.append({
                    "Material": material,
                    "Plant": plant,
                    "Field": field,
                    "Current_Count": last_run["error_count"],
                    "Peak_Count": peak_count,
                    "Status": last_run["status"],
                    "Trend": trend,
                    "First_Seen": first_run["date"],
                    "Last_Seen": last_run["date"],
                    "Days_Active": days_active
                })
    
    return pd.DataFrame(rows) if rows else pd.DataFrame()

def build_owner_dashboard_df(enriched_df: pd.DataFrame) -> pd.DataFrame:
    """
    Builds a dashboard-ready DataFrame with aggregated metrics per owner.
    
    Requires enriched_df from merge_assignments().
    
    Output columns:
    Owner | Active_Errors | Fields | Improving | Worsening | Stable |
    Overdue | Priority | Score
    """

    if enriched_df is None or enriched_df.empty:
        return pd.DataFrame()

    if "Owner" not in enriched_df.columns:
        return pd.DataFrame()

    owners = sorted([o for o in enriched_df["Owner"].dropna().unique() if str(o).strip()])

    rows = []

    for owner in owners:

        # 👉 usa a tua função robusta
        metrics = get_owner_metrics(owner)

        if not metrics:
            continue

        total = metrics.get("total_active_errors", 0)
        improving = metrics.get("improving", 0)
        worsening = metrics.get("worsening", 0)
        stable = metrics.get("stable", 0)

        # ✅ OVERDUE EXTRAÍDO DO DF (mais preciso)
        owner_df = enriched_df[enriched_df["Owner"] == owner]
        overdue = int((owner_df["Status"] == "Overdue").sum()) if "Status" in owner_df.columns else 0

        fields = metrics.get("fields_count", 0)

        # ✅ SCORING (simples e eficaz)
        score = (
            total
            + (worsening * 20)
            + (overdue * 30)
            - (improving * 10)
        )

        # ✅ PRIORIDADE
        if overdue > 5 or (worsening >= 3 and total > 100):
            priority = "🔴 High"
        elif overdue > 0 or total > 50:
            priority = "🟠 Medium"
        else:
            priority = "🟢 Low"

        rows.append({
            "Owner": owner,
            "Active_Errors": total,
            "Fields": fields,
            "Improving": improving,
            "Worsening": worsening,
            "Stable": stable,
            "Overdue": overdue,
            "Score": score,
            "Priority": priority,
        })

    df = pd.DataFrame(rows)

    if df.empty:
        return df

    # ✅ ordenação (mais crítico primeiro)
    df = df.sort_values(["Score", "Active_Errors"], ascending=[False, False])

    return df.reset_index(drop=True)

# =============================================================================
# SSA VALIDATION ENGINE
# Sales Orders and Schedule Agreements (SO & SA) validation module
# Tables: VBAK, VBAP, VBKD, VBPA, KNA1, MARC, VBFA
# =============================================================================

def _ssa_norm(x):
    """Normalise a cell value to stripped string; NaN/None -> ''."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    return str(x).strip()


def _ssa_pick(df, candidates):
    """Return first matching column name (case-insensitive)."""
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None


def _ssa_load(df):
    """Normalise column names: strip whitespace."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


# -----------------------------------------------------------------------------
# SSA structural diagnostics - source tables and fields used by every rule
# -----------------------------------------------------------------------------
SSA_DIAGNOSTIC_COLUMNS = [
    "Rule",
    "Check",
    "Table",
    "Logical Field",
    "Accepted Column Names",
    "Status",
    "Impact",
]


def _ssa_diag_check(df, rule, check, table, fields, source_required=True):
    """
    Inspect one table dependency and return one diagnostic row per logical field.

    Status values:
      - OK: field exists and the corresponding check can use it.
      - MISSING: required field does not exist; the check is skipped/partial.
      - SOURCE MISSING: required source table was not supplied or is empty.
      - OPTIONAL SOURCE MISSING: optional source was not supplied; fallback applies.
    """
    rows = []
    source_available = df is not None and not df.empty
    work = _ssa_load(df) if source_available else None

    if not source_available:
        status = "SOURCE MISSING" if source_required else "OPTIONAL SOURCE MISSING"
        for logical_field, definition in fields.items():
            candidates = definition.get("candidates", [logical_field])
            required = definition.get("required", True)
            impact = definition.get("impact", "Validation cannot use this field.")
            if not required and source_required:
                status_for_field = "OPTIONAL SOURCE MISSING"
            else:
                status_for_field = status
            rows.append({
                "Rule": rule,
                "Check": check,
                "Table": table,
                "Logical Field": logical_field,
                "Accepted Column Names": ", ".join(candidates),
                "Status": status_for_field,
                "Impact": impact,
            })
        return rows

    for logical_field, definition in fields.items():
        candidates = definition.get("candidates", [logical_field])
        required = definition.get("required", True)
        column = _ssa_pick(work, candidates)
        if column is not None:
            status = "OK"
            impact = f"Using source column '{column}'."
        elif required:
            status = "MISSING"
            impact = definition.get("impact", "This validation check is skipped.")
        else:
            status = "OPTIONAL MISSING"
            impact = definition.get("impact", "Optional behaviour is unavailable.")
        rows.append({
            "Rule": rule,
            "Check": check,
            "Table": table,
            "Logical Field": logical_field,
            "Accepted Column Names": ", ".join(candidates),
            "Status": status,
            "Impact": impact,
        })
    return rows


def _ssa_run_input_diagnostics(vbak, vbap, vbkd, kna1, marc=None, vbpa=None, vbfa=None, vbep=None, mvke=None):
    """Validate all table/field dependencies used by the current SSA engine."""
    diagnostics = []

    def add(df, rule, check, table, fields, source_required=True):
        diagnostics.extend(_ssa_diag_check(
            df, rule, check, table, fields, source_required=source_required
        ))

    # Shared applicability and active-document logic.
    add(vbak, "Engine Setup", "Document type mapping", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Type": {"candidates": ["Sales Document Type", "Order Type", "AUART"],
                                "impact": "Rules based on AUART cannot determine applicability."},
    })
    add(vbap, "Engine Setup", "Closed-document and item scope", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"],
                                 "impact": "Closed/rejected documents and items cannot be excluded correctly."},
        "Material": {"candidates": ["Material", "MATNR"],
                     "impact": "S-* material scoping cannot be applied."},
        "Plant": {"candidates": ["Plant", "WERKS"], "required": False,
                  "impact": "Plant filter cannot be applied."},
    })

    # Mandatory Fields, split by independently executable sub-check.
    for field, groups in [
        ("Customer Group 1", "ZLZ/ZLZM, ZTAM/ZOR/..., ZDIC/ZMC/..., ZMD/ZDSD/... (all order types)"),
        ("Customer Group 2", "ZLZ/ZLZM and ZTAM/ZOR/..."),
        ("Customer Group 3", "ZLZ/ZLZM"),
    ]:
        fields = {
            "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
            "Overall Header": {
                "candidates": ["Overall Header", "UVALL"],
                "required": False,
                "impact": "Completed VBAK documents with Overall Header='C' cannot be skipped.",
            },
            field: {"candidates": [field], "impact": f"Mandatory check for {field} ({groups}) is skipped."},
        }
        if field == "Customer Group 3":
            fields["MRP for DlvSchType"] = {"candidates": ["MRP for DlvSchType", "MRPforDS", "ABDIS"], "required": False,
                                    "impact": "ZLZ exception ABDIS=E cannot be identified; blank ABDIS is treated as not E."}
        add(vbak, "Mandatory Fields", field, "VBAK", fields)

    mandatory_vbap_checks = [
        ("Route", ["Route", "ROUTE"]),
        ("IFS ID", ["IFS ID"]),
        ("Supplier_code", ["Supplier_code"]),
        ("Plant", ["Plant", "WERKS"]),
        ("Acct Assmt Grp Mat.", ["Acct Assmt Grp Mat.", "Account Assignment Group", "Acct Assignment Grp", "Acc.Assignment Grp.", "KTGRM"]),
    ]
    for logical, candidates in mandatory_vbap_checks:
        add(vbap, "Mandatory Fields", logical, "VBAP", {
            "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
            "Sales Document Item": {"candidates": ["Sales Document Item", "Item (SD)", "POSNR"], "required": False,
                                    "impact": "Item identifier will be blank in reported errors."},
            "Material": {"candidates": ["Material", "MATNR"], "required": logical == "Acct Assmt Grp Mat.",
                         "impact": "Material-dependent KTGRM logic cannot run." if logical == "Acct Assmt Grp Mat." else "Material will be blank in reported errors."},
            "Plant": {"candidates": ["Plant", "WERKS"], "required": logical == "Plant",
                      "impact": "Plant mandatory check is skipped." if logical == "Plant" else "Plant will be blank in reported errors."},
            "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                     "impact": "Rejected items cannot be excluded from this check."},
            logical: {"candidates": candidates, "impact": f"Mandatory check for {logical} is skipped."},
        })

    add(vbak, "Terms of Payment", "Applicable Sales Documents", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "Sales Doc.", "VBELN"]},
    })
    add(vbkd, "Terms of Payment", "Header value", "VBKD", {
        "Sales Document": {"candidates": ["SD Document", "Sales Document", "Sales Doc.", "VBELN"]},
        "Item": {"candidates": ["Item (SD)", "Sales Document Item", "POSNR", "Item"], "required": False,
                 "impact": "Cannot restrict the lookup to header POSNR; first document value may be used."},
        "Terms of Payment": {"candidates": ["Terms of Payment", "Terms of Payment Key", "ZTERM"],
                             "impact": "Terms of Payment check is skipped."},
    })

    add(vbap, "Shipping Point", "Mandatory and equal to Plant", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Item": {"candidates": ["Sales Document Item", "POSNR"]},
        "Material": {"candidates": ["Material", "MATNR"]},
        "Plant": {"candidates": ["Plant", "WERKS"]},
        "Shipping Point": {"candidates": ["Shipping Point/Receiving Pt", "Shipping Point", "VSTEL"]},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                 "impact": "Rejected items cannot be excluded."},
    })

    add(vbpa, "Unloading Point Consistency", "Header versus item", "VBPA", {
        "Sales Document": {"candidates": ["SD Document", "VBELN"]},
        "Item": {"candidates": ["Item (SD)", "POSNR"]},
        "Partner Function": {"candidates": ["Partner Function", "PARVW"]},
        "Unloading Point": {"candidates": ["Unloading Point"]},
    }, source_required=True)

    add(vbkd, "Customer Reference", "VBKD-BSTKD item/header fallback", "VBKD", {
        "Sales Document": {"candidates": ["SD Document", "Sales Document", "VBELN"]},
        "Item": {"candidates": ["Item (SD)", "Sales Document Item", "POSNR"], "required": False,
                 "impact": "Item-specific lookup is unavailable; document-level fallback only."},
        "Customer Reference": {"candidates": ["Customer Reference", "Customer Purchase Order Number", "Customer PO Number", "Cust. Reference", "BSTKD"],
                               "impact": "Customer Reference validation is skipped."},
    })
    add(vbap, "Customer Reference", "Items to validate", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Item": {"candidates": ["Sales Document Item", "Item (SD)", "POSNR"]},
        "Material": {"candidates": ["Material", "MATNR"], "required": False, "impact": "Material is blank in errors."},
        "Plant": {"candidates": ["Plant", "WERKS"], "required": False, "impact": "Plant is blank in errors."},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                 "impact": "Rejected items cannot be excluded."},
    })

    add(vbap, "Weights", "Gross/Net values", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Item": {"candidates": ["Sales Document Item", "POSNR"]},
        "Material": {"candidates": ["Material", "MATNR"]},
        "Plant": {"candidates": ["Plant", "WERKS"]},
        "Gross Weight": {"candidates": ["Gross Weight"]},
        "Net Weight": {"candidates": ["Net Weight"]},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                 "impact": "Rejected items cannot be excluded."},
    })
    add(vbfa, "Weights", "OBD gating", "VBFA", {
        "Preceding Document": {"candidates": ["Preceding Document", "VBELV"]},
        "Preceding Item": {"candidates": ["Item (SD) Preceding Doc.", "Preceding Item", "POSNV"], "required": False,
                           "impact": "OBD gating can only be applied at document level."},
        "Subsequent Document": {"candidates": ["Subsequent Document", "VBELN"]},
        "Preceding Document Category": {
            "candidates": ["Subsequent Doc. Cat.", "Preceding Document Category", "VBTYP_N"],
            "impact": "OBD gating is skipped because the Subsequent Document cannot be confirmed as a delivery (category J).",
        },
    }, source_required=False)

    add(vbep, "Weights", "Active demand gating", "VBEP", {
        "Sales Document": {"candidates": ["Sales Document", "SD Document", "VBELN"],
                           "impact": "ZLZ/ZLZM active demand cannot be matched to VBAP."},
        "Sales Document Item": {"candidates": ["Sales Document Item", "Item (SD)", "POSNR"],
                                "impact": "ZLZ/ZLZM active demand cannot be matched at item level."},
        "Release Type": {"candidates": ["Release Type", "ABART"],
                         "impact": "ZLZ/ZLZM active demand cannot be identified; weight checks are skipped for those document types."},
    }, source_required=False)
    add(vbak, "Incoterms", "DDP + GB header checks", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sold-to Party": {"candidates": ["Sold-to Party", "KUNNR"]},
        "Alt.Tax Classific.": {"candidates": ["Alt.Tax Classific.", "TAXK1"]},
        "Tax Departure C/R": {"candidates": ["Tax Departure C/R", "STCEG_L"]},
    })
    add(vbkd, "Incoterms", "Mandatory Incoterms", "VBKD", {
        "Sales Document": {"candidates": ["SD Document", "VBELN"]},
        "Incoterms": {"candidates": ["Incoterms", "INCO1"]},
    })
    add(kna1, "Incoterms", "Customer country", "KNA1", {
        "Customer": {"candidates": ["Customer", "KUNNR"]},
        "Country/Region Key": {"candidates": ["Country/Region Key", "LAND1"]},
    })
    add(vbpa, "Incoterms", "Optional SP country override", "VBPA", {
        "Sales Document": {"candidates": ["SD Document", "VBELN"]},
        "Partner Function": {"candidates": ["Partner Function", "PARVW"]},
        "Country/Region Key": {"candidates": ["Country/Region Key", "LAND1"]},
    }, source_required=False)

    add(vbak, "Incompletion Log", "Overall Header", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"], "required": False,
                           "impact": "Document identifier is blank in errors."},
        "Overall Header": {"candidates": ["Overall Header", "UVALL"], "impact": "Header incompletion check is skipped."},
    })
    add(vbap, "Incompletion Log", "Item status", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"], "required": False},
        "Sales Document Item": {"candidates": ["Sales Document Item", "POSNR"], "required": False},
        "Material": {"candidates": ["Material", "MATNR"], "required": False},
        "Plant": {"candidates": ["Plant", "WERKS"], "required": False},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                 "impact": "Rejected items cannot be excluded."},
        "Item Incompletion Status": {"candidates": ["UVPOS", "Item Incompletion Status", "Item Status"],
                                     "impact": "Item incompletion check is skipped."},
    })

    add(vbap, "Rejected + Status 70-75-99", "Open item identification", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Item": {"candidates": ["Sales Document Item", "POSNR"], "required": False},
        "Material": {"candidates": ["Material", "MATNR"]},
        "Plant": {"candidates": ["Plant", "WERKS"]},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"]},
    })
    add(marc, "Rejected + Status 70-75-99", "Plant material status", "MARC", {
        "Material": {"candidates": ["Material", "MATNR"]},
        "Plant": {"candidates": ["Plant", "WERKS"]},
        "Plant-Specific Material Status": {"candidates": ["Plant-Specific Material Status", "Plant Specific Material Status", "MTSTB"]},
    }, source_required=True)

    add(mvke, "Mandatory Fields", "S-material Acct Assmt Grp Mat. lookup", "MVKE", {
        "Material": {"candidates": ["Material", "MATNR"]},
        "Sales Organization": {"candidates": ["Sales Organization", "Sales Org", "VKORG"]},
        "Distribution Channel": {"candidates": ["Distribution Channel", "Distr. Channel", "VTWEG"]},
        "Acct Assmt Grp Mat.": {"candidates": ["Acct Assmt Grp Mat.", "Account Assignment Group", "KTGRM"],
                                "impact": "S-materials in SSA_S_MATERIAL_LIST fall back to the plain non-blank check instead of the MVKE comparison."},
    }, source_required=False)

    add(vbak, "Order Without Items", "Header documents", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
    })
    add(vbap, "Order Without Items", "Item documents", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
    })

    add(vbap, "Customer Material", "Mandatory KDMAT", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Item": {"candidates": ["Sales Document Item", "POSNR"]},
        "Material": {"candidates": ["Material", "MATNR"]},
        "Plant": {"candidates": ["Plant", "WERKS"]},
        "Customer Material Number": {"candidates": ["Customer Material Number", "Material by Customer", "KDMAT"]},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                 "impact": "Rejected items cannot be excluded."},
    })

    add(vbak, "MRP for DS Type", "Mandatory ABDIS", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "MRP for DlvSchType": {"candidates": ["MRP for DlvSchType", "MRPforDS", "ABDIS"]},
    })

    return pd.DataFrame(diagnostics, columns=SSA_DIAGNOSTIC_COLUMNS)


def _ssa_print_input_diagnostics(diagnostics_df):
    """Print a concise execution/readiness report for every SSA check."""
    line = "=" * 110
    print("\n" + line)
    print("SSA INPUT FIELD DIAGNOSTICS")
    print(line)

    if diagnostics_df is None or diagnostics_df.empty:
        print("[SSA][WARNING] No diagnostics were produced.")
        print(line + "\n")
        return

    problem_statuses = {"MISSING", "SOURCE MISSING", "OPTIONAL MISSING", "OPTIONAL SOURCE MISSING"}
    problems = diagnostics_df[diagnostics_df["Status"].isin(problem_statuses)]

    grouped = diagnostics_df.groupby(["Rule", "Check"], sort=False)
    for (rule, check), group in grouped:
        required_problem = group["Status"].isin({"MISSING", "SOURCE MISSING"}).any()
        optional_problem = group["Status"].isin({"OPTIONAL MISSING", "OPTIONAL SOURCE MISSING"}).any()
        state = "SKIPPED/PARTIAL" if required_problem else ("RUN WITH FALLBACK" if optional_problem else "READY")
        print(f"[SSA][{state}] {rule} | {check}")

    if problems.empty:
        print("\n[SSA][OK] All configured source fields were found.")
    else:
        print("\nMissing or unavailable dependencies:")
        for _, row in problems.iterrows():
            print(
                f" - [{row['Status']}] {row['Rule']} | {row['Check']} | "
                f"{row['Table']}.{row['Logical Field']}"
            )
            print(f"   Accepted: {row['Accepted Column Names']}")
            print(f"   Impact: {row['Impact']}")

    print(line + "\n")

def _ssa_to_float(x):
    """Convert SAP/Excel numeric-like value to float; blank/invalid -> None."""
    s = _ssa_norm(x)
    if s == '':
        return None
    try:
        return float(str(s).replace(',', '.'))
    except Exception:
        return None


def _ssa_is_zero_or_blank(x):
    """True if value is blank, None, NaN or numeric zero."""
    s = _ssa_norm(x)
    if s == '':
        return True
    try:
        return float(str(s).replace(',', '.')) == 0.0
    except Exception:
        return False


def _ssa_float_equal(a, b, tol=1e-6):
    """Numeric comparison with tolerance."""
    fa = _ssa_to_float(a)
    fb = _ssa_to_float(b)
    if fa is None or fb is None:
        return False
    return abs(fa - fb) <= tol


def _ssa_first_nonempty(*values):
    """Return first non-empty normalized value."""
    for v in values:
        s = _ssa_norm(v)
        if s != '':
            return s
    return ''


def _ssa_get_vbpa_value(vbpa_df, doc, item, partner_func, field_name="Unloading Point"):
    """
    Returns the first matching VBPA value for a given document/item/partner function.
    Prefers exact item match; if not found, returns header-level item=0/blank.
    """
    if vbpa_df is None or vbpa_df.empty:
        return ''

    vbpa = _ssa_load(vbpa_df)

    doc_col = _ssa_pick(vbpa, ["SD Document", "VBELN"])
    item_col = _ssa_pick(vbpa, ["Item (SD)", "POSNR"])
    pf_col   = _ssa_pick(vbpa, ["Partner Function", "PARVW"])
    val_col  = _ssa_pick(vbpa, [field_name])

    if not all([doc_col, item_col, pf_col, val_col]):
        return ''

    df = vbpa[vbpa[doc_col].astype(str).str.strip() == str(doc).strip()].copy()
    if df.empty:
        return ''

    df = df[df[pf_col].astype(str).str.strip().upper() == str(partner_func).strip().upper()]
    if df.empty:
        return ''

    item_str = str(item).strip()
    exact = df[df[item_col].astype(str).str.strip() == item_str]
    if not exact.empty:
        return _ssa_first_nonempty(*exact[val_col].tolist())

    hdr = df[df[item_col].astype(str).str.strip().isin(['0', '', '000000', '0000000000'])]
    if not hdr.empty:
        return _ssa_first_nonempty(*hdr[val_col].tolist())

    return ''


# -----------------------------------------------------------------------------
# Shared helpers — Sales Document prefix / Sales Document Type / VBFA (OBD)
# -----------------------------------------------------------------------------
SSA_ORDER_TYPES_ZLZ_GROUP = {"ZLZ", "ZLZM"}
SSA_ORDER_TYPES_ZTAM_GROUP = {"ZTAM", "ZOR", "ZPRO", "ZFD", "ZINI", "ZORA", "ZFD2", "ZORM"}
SSA_ORDER_TYPES_ZDIC_GROUP = {"ZDIC", "ZMC", "ZRE", "ZDSC", "ZRE2", "ZMC2"}
SSA_ORDER_TYPES_ZMD_GROUP = {"ZMD", "ZDSD", "ZDID", "ZMD2"}

# Convenience union used by several rules that apply to both the ZLZ/ZLZM
# and ZTAM/ZOR/... Sales Document Types.
SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP = SSA_ORDER_TYPES_ZLZ_GROUP | SSA_ORDER_TYPES_ZTAM_GROUP

# Sales Document Types for which Customer Group 2 is mandatory.
SSA_ORDER_TYPES_CUSTOMER_GROUP_2 = {"ZLZ", "ZLZM", "ZOR", "ZTAM"}

# Every Sales Document Type known to the SSA engine.
SSA_ORDER_TYPES_ALL = (
    SSA_ORDER_TYPES_ZLZ_GROUP
    | SSA_ORDER_TYPES_ZTAM_GROUP
    | SSA_ORDER_TYPES_ZDIC_GROUP
    | SSA_ORDER_TYPES_ZMD_GROUP
)

# S-materials validated, for the applicable Sales Document Types, against
# the "Acct Assmt Grp Mat." value carried for them in MVKE (see
# _ssa_build_s_material_mvke_map and Rule 3 inside ssa_validar_mandatory_fields).
SSA_S_MATERIAL_LIST = [
    "S-306", "S-503", "S-304", "S-300", "S-200", "S-202", "S-301", "S-309",
    "S-100", "S-302", "S-507", "S-506", "S-305", "S-303", "S-502", "S-504",
    "S-505", "S-101", "S-102", "S-500", "S-501", "S-308", "S-509", "S-310",
    "S-311",
]

# Populated by run_ssa_validation. The compatibility helper below deliberately
# uses AUART and no longer reads the first digit of the document number.
_SSA_ACTIVE_DOC_TYPE_MAP = {}


def _ssa_doc_type_in(doc, allowed_types):
    """
    True if the Sales Document's Sales Document Type (AUART) is a member of
    allowed_types (e.g. SSA_ORDER_TYPES_ZLZ_GROUP).

    Rule applicability depends solely on AUART — the Sales Document number
    itself is never inspected, so no rule depends on Sales Document number
    ranges.
    """
    doc_type = _SSA_ACTIVE_DOC_TYPE_MAP.get(_ssa_norm(doc), '').upper()
    return bool(doc_type) and doc_type in allowed_types


def _ssa_filter_docs(df, allowed_docs, doc_candidates):
    """Return a copy containing only allowed Sales Documents."""
    if df is None or df.empty:
        return df
    work = _ssa_load(df)
    doc_col = _ssa_pick(work, doc_candidates)
    if not doc_col:
        return work
    return work[
        work[doc_col].astype(str).str.strip().isin(allowed_docs)
    ].copy()


def _ssa_build_doc_plant_map(vbap_df):
    """
    Builds a Sales Document -> Plant lookup from VBAP, keeping the first
    non-blank Plant found for each Sales Document. Used to backfill the
    'Plant' column in the final error report for rules that do not report
    Plant themselves (e.g. header-level checks).
    """
    if vbap_df is None or vbap_df.empty:
        return {}

    work = _ssa_load(vbap_df)
    doc_col = _ssa_pick(work, ["Sales Document", "VBELN"])
    plt_col = _ssa_pick(work, ["Plant", "WERKS"])

    if doc_col is None or plt_col is None:
        return {}

    doc_plant_map = {}
    for doc, plant in zip(
        work[doc_col].map(_ssa_norm),
        work[plt_col].map(_ssa_norm),
    ):
        if doc == "" or plant == "":
            continue
        doc_plant_map.setdefault(doc, plant)

    return doc_plant_map


def _ssa_fill_missing_plant(errors_df, doc_plant_map):
    """
    Fills blank 'Plant' values in errors_df using the Sales Document ->
    Plant lookup built from VBAP via _ssa_build_doc_plant_map. Rows whose
    Sales Document has no match in VBAP, or that already carry a Plant
    value, are left unchanged.
    """
    if errors_df is None or errors_df.empty or not doc_plant_map:
        return errors_df
    if "Plant" not in errors_df.columns or "Sales Document" not in errors_df.columns:
        return errors_df

    errors_df = errors_df.copy()
    blank_mask = errors_df["Plant"].map(_ssa_norm) == ""
    errors_df.loc[blank_mask, "Plant"] = (
        errors_df.loc[blank_mask, "Sales Document"]
        .map(_ssa_norm)
        .map(doc_plant_map)
        .fillna("")
    )
    return errors_df


def _ssa_filter_non_s_items(vbap_df):
    """Exclude materials whose number starts with 'S-' (case-insensitive)."""
    if vbap_df is None or vbap_df.empty:
        return vbap_df
    work = _ssa_load(vbap_df)
    mat_col = _ssa_pick(work, ["Material", "MATNR"])
    if not mat_col:
        return work
    is_s = work[mat_col].apply(lambda value: _ssa_norm(value).upper().startswith('S-'))
    return work[~is_s].copy()


def _ssa_filter_vbpa_non_s_items(vbpa_df, vbap_df):
    """
    Exclude VBPA item rows that belong to S-* VBAP items.

    Header rows are retained only for documents that still have at least one
    non-S item. This allows header/item comparisons for mixed documents while
    excluding documents made exclusively of S-* materials.
    """
    if vbpa_df is None or vbpa_df.empty or vbap_df is None or vbap_df.empty:
        return vbpa_df

    vbpa = _ssa_load(vbpa_df)
    vbap = _ssa_load(vbap_df)
    vp_doc = _ssa_pick(vbpa, ["SD Document", "Sales Document", "VBELN"])
    vp_item = _ssa_pick(vbpa, ["Item (SD)", "Sales Document Item", "POSNR"])
    ap_doc = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    ap_item = _ssa_pick(vbap, ["Sales Document Item", "Item (SD)", "POSNR"])
    ap_mat = _ssa_pick(vbap, ["Material", "MATNR"])
    if not all([vp_doc, vp_item, ap_doc, ap_item, ap_mat]):
        return vbpa

    non_s = vbap[~vbap[ap_mat].apply(
        lambda value: _ssa_norm(value).upper().startswith('S-')
    )].copy()
    allowed_docs = set(non_s[ap_doc].astype(str).str.strip())
    allowed_items = set(zip(
        non_s[ap_doc].astype(str).str.strip(),
        non_s[ap_item].astype(str).str.strip(),
    ))
    header_values = {'0', '', '000000', '0000000000'}

    keep = []
    for _, row in vbpa.iterrows():
        doc = _ssa_norm(row[vp_doc])
        item = _ssa_norm(row[vp_item])
        keep.append(doc in allowed_docs and (item in header_values or (doc, item) in allowed_items))
    return vbpa.loc[keep].copy()


def _ssa_doc_type_map(vbak_df):
    """Build a map: Sales Document -> Sales Document Type (AUART), from VBAK."""
    if vbak_df is None or vbak_df.empty:
        return {}
    vbak = _ssa_load(vbak_df)
    doc_col  = _ssa_pick(vbak, ["Sales Document", "VBELN"])
    type_col = _ssa_pick(vbak, ["Sales Document Type", "Order Type", "AUART"])
    if not doc_col or not type_col:
        return {}
    return dict(zip(
        vbak[doc_col].astype(str).str.strip(),
        vbak[type_col].astype(str).str.strip().str.upper()
    ))


def _ssa_add_order_type_column(df):
    """
    Insert an 'Order Type' column (Sales Document Type / AUART) right after
    'Sales Document', resolved via the active doc-type map built for this
    run (_SSA_ACTIVE_DOC_TYPE_MAP). Rows whose Sales Document is unknown or
    blank get an empty value. No-op if df is empty or lacks the column.
    """
    if df is None or df.empty or "Sales Document" not in df.columns:
        return df
    df = df.copy()
    df["Order Type"] = df["Sales Document"].apply(
        lambda doc: _SSA_ACTIVE_DOC_TYPE_MAP.get(_ssa_norm(doc), "")
    )
    cols = [c for c in df.columns if c != "Order Type"]
    insert_at = cols.index("Sales Document") + 1
    cols.insert(insert_at, "Order Type")
    return df[cols]


def _ssa_item_is_done(row, status_col, rej_col):
    """
    Return True when a VBAP item is 'done' for SSA validation purposes,
    and can therefore be skipped by every rule in the SSA report:
      - Overall Status = 'C' (complete), OR
      - Reason for Rejection (VBAP-ABGRU) is filled.
    Either condition on its own is enough to skip the item.
    """
    if status_col and _ssa_norm(row[status_col]).upper() == "C":
        return True
    if rej_col and _ssa_norm(row[rej_col]) != "":
        return True
    return False


def _ssa_fully_rejected_docs(vbap_df):
    """
    Return documents for which every VBAP item is 'done': Overall Status
    = 'C', or Reason for Rejection (VBAP-ABGRU) filled (see
    _ssa_item_is_done).
    """
    if vbap_df is None or vbap_df.empty:
        return set()
    vbap = _ssa_load(vbap_df)
    doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    status_col = _ssa_pick(vbap, ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"])
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])
    if not doc_col or not status_col:
        return set()
    closed = set()
    for doc, grp in vbap.groupby(vbap[doc_col].map(_ssa_norm)):
        if doc and grp.apply(lambda r: _ssa_item_is_done(r, status_col, rej_col), axis=1).all():
            closed.add(doc)
    return closed

def _ssa_build_obd_set(vbfa_df):
    """
    Build the set of Sales Documents/items for which an Outbound Delivery
    has already been created, using VBFA Sales Document Flow.

    Correct document-flow logic:
      - VBELV / POSNV identify the preceding Sales Document and item.
      - VBELN identifies the Subsequent Document.
      - VBTYP_N must be 'J', meaning the Subsequent Document is a Delivery.

    Returns (obd_map, ready):
      obd_map: {sales_document: set(items)}
      ready: True only when VBFA and all mandatory OBD-gating columns exist.

    If VBTYP_N is unavailable, ready=False. The engine must not infer that a
    preceding-category value of J represents a Sales Order -> Delivery flow.
    """
    if vbfa_df is None or vbfa_df.empty:
        return {}, False

    vbfa = _ssa_load(vbfa_df)

    prec_doc_col = _ssa_pick(
        vbfa,
        ["Preceding Document", "VBELV"],
    )
    prec_item_col = _ssa_pick(
        vbfa,
        ["Item (SD) Preceding Doc.", "Preceding Item", "POSNV"],
    )
    subs_doc_col = _ssa_pick(
        vbfa,
        ["Subsequent Document", "VBELN"],
    )
    subs_cat_col = _ssa_pick(
        vbfa,
        ["Subsequent Doc. Cat.", "Preceding Document Category", "VBTYP_N"],
    )

    if not all([prec_doc_col, subs_doc_col, subs_cat_col]):
        return {}, False

    selected = [prec_doc_col, subs_doc_col, subs_cat_col]
    if prec_item_col:
        selected.append(prec_item_col)
    selected = list(dict.fromkeys(selected))

    work = vbfa.loc[:, selected].copy()
    work["_ssa_prec_doc"] = work[prec_doc_col].map(_ssa_norm)
    work["_ssa_subs_doc"] = work[subs_doc_col].map(_ssa_norm)
    work["_ssa_subs_cat"] = work[subs_cat_col].map(
        lambda value: _ssa_norm(value).upper()
    )

    if prec_item_col:
        work["_ssa_prec_item"] = work[prec_item_col].map(_ssa_norm)
    else:
        work["_ssa_prec_item"] = ""

    work = work.loc[
        (work["_ssa_prec_doc"] != "")
        & (work["_ssa_subs_doc"] != "")
        & (work["_ssa_subs_cat"] == "J")
    ]

    if work.empty:
        return {}, True

    result = (
        work.groupby("_ssa_prec_doc", sort=False)["_ssa_prec_item"]
        .agg(set)
        .to_dict()
    )
    return result, True

def _ssa_build_s_material_mvke_map(mvke_df, material_list):
    """
    Build a Material -> Acct Assmt Grp Mat. (KTGRM) map from MVKE, restricted
    to material_list.

    A material may carry more than one MVKE entry (one per Sales
    Organization / Distribution Channel combination); the first non-blank
    AAG value found for that material is used as the expected value.
    Returns {} if mvke_df is missing/empty or lacks the required columns.
    """
    if mvke_df is None or mvke_df.empty:
        return {}

    mvke_norm, cols = _prepare_mvke(mvke_df, set(material_list))
    if mvke_norm is None or mvke_norm.empty:
        return {}

    result = {}
    for material, grp in mvke_norm.groupby('_mat'):
        value = _ssa_first_nonempty(*grp['_aag'].tolist())
        if value:
            result[material] = value
    return result


def _ssa_has_obd(obd_map, doc, item):
    """True if an OBD (outbound delivery) already exists for this document/item."""
    items = obd_map.get(_ssa_norm(doc))
    if items is None:
        return False
    if '' in items:
        return True
    return _ssa_norm(item) in items


# -----------------------------------------------------------------------------
# Rule 1 — Mandatory fields
# -----------------------------------------------------------------------------
def ssa_validar_mandatory_fields(vbak_df, vbap_df, marc_df=None, mvke_df=None):
    """
    Validate mandatory VBAK and VBAP fields.

    Rule applicability is based on Sales Document Type (AUART) only, through
    the SSA_ORDER_TYPES_ZLZ_GROUP / SSA_ORDER_TYPES_ZTAM_GROUP /
    SSA_ORDER_TYPES_ZDIC_GROUP / SSA_ORDER_TYPES_ZMD_GROUP sets. The Sales
    Document number is never inspected.

    VBAK:
      Customer Group checks are skipped when VBAK Overall Header (UVALL)
      is equal to "C".

      - Customer Group 1:
            Mandatory for Sales Document Types in SSA_ORDER_TYPES_ALL
            (ZLZ/ZLZM, ZTAM/ZOR/ZPRO/ZFD/ZINI/ZORA/ZFD2/ZORM,
            ZDIC/ZMC/ZRE/ZDSC/ZRE2/ZMC2, ZMD/ZDSD/ZDID/ZMD2).

      - Customer Group 2:
            Mandatory only for Sales Document Types ZLZ, ZLZM, ZOR, ZTAM
            (SSA_ORDER_TYPES_CUSTOMER_GROUP_2).

      - Customer Group 3:
            Mandatory for Sales Document Types in SSA_ORDER_TYPES_ZLZ_GROUP,
            with the following refinement:
              * ZLZ:
                    Mandatory only when VBAK-ABDIS ("MRP for DlvSchType") is not "E".
              * ZLZM:
                    Always mandatory.

    VBAP:
      Items where Reason for Rejection (VBAP-ABGRU) is filled are skipped.

      - Route:
            Mandatory for Sales Document Types in SSA_ORDER_TYPES_ZLZ_GROUP.

      - IFS ID:
            Mandatory for Sales Document Types in
            SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP.

      - Supplier_code:
            Mandatory for Sales Document Types in
            SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP when Sales Document Type
            is one of:
                ZOR, ZTAM, ZORA
            (ZPRO is explicitly excluded from this check.)

      - Plant:
            Mandatory for Sales Document Types in
            SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP.

      - Acct Assmt Grp Mat. (KTGRM):

            Applies only to active items, meaning VBAP-ABGRU is blank.

            KTGRM must not be blank for the following Sales Document Types:
                ZLZ, ZLZM, ZRE, ZPRO, ZOR, ZORM, ZFD, ZFD2, ZRE2,
                ZMD, ZMD2, ZMC, ZMC2, ZDID, ZDIC

            Additional validations:

              1. For ZLZ, ZLZM and ZRE:
                   If the Material Number does not start with "S-",
                   KTGRM must be equal to "Z1".

                   If the Material Number starts with "S-", only the
                   mandatory non-blank validation applies.

              2. For ZPRO, ZOR, ZORM, ZFD, ZFD2, ZRE2, ZMD, ZMD2,
                 ZMC, ZMC2, ZDID and ZDIC:

                   If the Material Number starts with "S-" and is one of the
                   S-materials in SSA_S_MATERIAL_LIST, KTGRM must match the
                   value found for that material in MVKE (field
                   "Acct Assmt Grp Mat.").

                   If the Material Number starts with "S-" but is not
                   included in SSA_S_MATERIAL_LIST (or MVKE carries no value
                   for it), KTGRM must only be non-blank.

                   For materials that do not start with "S-", KTGRM must
                   only be non-blank.

    Note:
      marc_df is retained in the function signature for backward compatibility
      with the existing orchestrator, but it is no longer used by the KTGRM
      validation. mvke_df is used only for the S-material comparison
      described above; when it is not supplied, S-materials fall back to
      the plain non-blank check.
    """
    errors = []
    doc_type_map = _ssa_doc_type_map(vbak_df)

    # ── VBAK ──────────────────────────────────────────────────────────────────
    if vbak_df is not None and not vbak_df.empty:
        vbak = _ssa_load(vbak_df)

        doc_col = _ssa_pick(
            vbak,
            ["Sales Document", "VBELN"],
        )

        abdis_col = _ssa_pick(
            vbak,
            ["MRP for DlvSchType", "MRPforDS", "ABDIS"],
        )

        header_status_col = _ssa_pick(
            vbak,
            [
                "Overall Header",
                "UVALL",
            ],
        )

        def is_completed_document(row):
            """Skip Customer Group checks when VBAK Overall Header = 'C'."""
            if header_status_col is None:
                return False

            return _ssa_norm(row[header_status_col]).upper() == "C"

        # ── Customer Group 1 / Customer Group 2 ───────────────────────────────
        field_order_types = {
            "Customer Group 1": SSA_ORDER_TYPES_ALL,
            "Customer Group 2": SSA_ORDER_TYPES_CUSTOMER_GROUP_2,
        }

        for field, allowed_types in field_order_types.items():
            col = _ssa_pick(vbak, [field])

            # If the column does not exist in the input file,
            # this specific validation cannot be performed.
            if col is None:
                continue

            for _, row in vbak.iterrows():
                # Skip completed sales documents.
                if is_completed_document(row):
                    continue

                doc = (
                    _ssa_norm(row[doc_col])
                    if doc_col
                    else ""
                )

                # Skip document types outside the scope of the field.
                if not _ssa_doc_type_in(doc, allowed_types):
                    continue

                # Generate an error when the applicable field is blank.
                if _ssa_norm(row[col]) == "":
                    errors.append({
                        "Table": "VBAK",
                        "Sales Document": doc,
                        "Item": "",
                        "Material": "",
                        "Plant": "",
                        "Field": field,
                        "Value": "",
                        "Error": (
                            f"Mandatory field blank: '{field}'"
                        ),
                    })

        # ── Customer Group 3 ──────────────────────────────────────────────────
        cg3_col = _ssa_pick(
            vbak,
            ["Customer Group 3"],
        )

        if cg3_col:
            for _, row in vbak.iterrows():
                # Skip completed sales documents.
                if is_completed_document(row):
                    continue

                doc = (
                    _ssa_norm(row[doc_col])
                    if doc_col
                    else ""
                )

                # Customer Group 3 applies only to the ZLZ document-type group.
                if not _ssa_doc_type_in(
                    doc,
                    SSA_ORDER_TYPES_ZLZ_GROUP,
                ):
                    continue

                doc_type = doc_type_map.get(doc, "").upper()

                # For ZLZ, Customer Group 3 is not mandatory when
                # MRP for DlvSchType is equal to E.
                if doc_type == "ZLZ":
                    abdis_val = (
                        _ssa_norm(row[abdis_col]).upper()
                        if abdis_col
                        else ""
                    )

                    if abdis_val == "E":
                        continue

                if _ssa_norm(row[cg3_col]) == "":
                    errors.append({
                        "Table": "VBAK",
                        "Sales Document": doc,
                        "Item": "",
                        "Material": "",
                        "Plant": "",
                        "Field": "Customer Group 3",
                        "Value": "",
                        "Error": (
                            "Mandatory field blank: "
                            "'Customer Group 3'"
                        ),
                    })

    # ── VBAP ──────────────────────────────────────────────────────────────────
    if vbap_df is not None and not vbap_df.empty:
        vbap = _ssa_load(vbap_df)

        doc_col = _ssa_pick(
            vbap,
            ["Sales Document", "VBELN"],
        )
        item_col = _ssa_pick(
            vbap,
            ["Sales Document Item", "Item (SD)", "POSNR"],
        )
        mat_col = _ssa_pick(
            vbap,
            ["Material", "MATNR"],
        )
        plt_col = _ssa_pick(
            vbap,
            ["Plant", "WERKS"],
        )
        rej_col = _ssa_pick(
            vbap,
            ["Reason for Rejection", "ABGRU"],
        )
        status_col = _ssa_pick(
            vbap,
            ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"],
        )

        # Items that are complete (Overall Status = 'C') or rejected
        # (VBAP-ABGRU filled) are excluded from all mandatory-field checks.
        def is_rejected(row):
            return _ssa_item_is_done(row, status_col, rej_col)

        # ── Route ─────────────────────────────────────────────────────────────
        route_col = _ssa_pick(
            vbap,
            ["Route", "ROUTE"],
        )

        if route_col:
            for _, row in vbap.iterrows():
                if is_rejected(row):
                    continue

                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_GROUP):
                    continue

                if _ssa_norm(row[route_col]) == "":
                    errors.append({
                        "Table": "VBAP",
                        "Sales Document": doc,
                        "Item": (
                            _ssa_norm(row[item_col])
                            if item_col
                            else ""
                        ),
                        "Material": (
                            _ssa_norm(row[mat_col])
                            if mat_col
                            else ""
                        ),
                        "Plant": (
                            _ssa_norm(row[plt_col])
                            if plt_col
                            else ""
                        ),
                        "Field": "Route",
                        "Value": "",
                        "Error": "Mandatory field blank: 'Route'",
                    })

        # ── IFS ID ────────────────────────────────────────────────────────────
        ifs_col = _ssa_pick(
            vbap,
            ["IFS ID"],
        )

        if ifs_col:
            for _, row in vbap.iterrows():
                if is_rejected(row):
                    continue

                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
                    continue

                if _ssa_norm(row[ifs_col]) == "":
                    errors.append({
                        "Table": "VBAP",
                        "Sales Document": doc,
                        "Item": (
                            _ssa_norm(row[item_col])
                            if item_col
                            else ""
                        ),
                        "Material": (
                            _ssa_norm(row[mat_col])
                            if mat_col
                            else ""
                        ),
                        "Plant": (
                            _ssa_norm(row[plt_col])
                            if plt_col
                            else ""
                        ),
                        "Field": "IFS ID",
                        "Value": "",
                        "Error": "Mandatory field blank: 'IFS ID'",
                    })

        # ── Supplier_code ─────────────────────────────────────────────────────
        sup_col = _ssa_pick(
            vbap,
            ["Supplier_code"],
        )

        if sup_col:
            # ZPRO is excluded from the Supplier_code check.
            supplier_allowed_types = {
                "ZOR",
                "ZTAM",
                "ZORA",
            }

            for _, row in vbap.iterrows():
                if is_rejected(row):
                    continue

                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
                    continue

                doc_type = doc_type_map.get(doc, "").upper()

                if doc_type not in supplier_allowed_types:
                    continue

                if _ssa_norm(row[sup_col]) == "":
                    errors.append({
                        "Table": "VBAP",
                        "Sales Document": doc,
                        "Item": (
                            _ssa_norm(row[item_col])
                            if item_col
                            else ""
                        ),
                        "Material": (
                            _ssa_norm(row[mat_col])
                            if mat_col
                            else ""
                        ),
                        "Plant": (
                            _ssa_norm(row[plt_col])
                            if plt_col
                            else ""
                        ),
                        "Field": "Supplier_code",
                        "Value": "",
                        "Error": (
                            "Mandatory field blank: 'Supplier_code'"
                        ),
                    })

        # ── Plant ─────────────────────────────────────────────────────────────
        if plt_col:
            for _, row in vbap.iterrows():
                if is_rejected(row):
                    continue

                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
                    continue

                if _ssa_norm(row[plt_col]) == "":
                    errors.append({
                        "Table": "VBAP",
                        "Sales Document": doc,
                        "Item": (
                            _ssa_norm(row[item_col])
                            if item_col
                            else ""
                        ),
                        "Material": (
                            _ssa_norm(row[mat_col])
                            if mat_col
                            else ""
                        ),
                        "Plant": "",
                        "Field": "Plant",
                        "Value": "",
                        "Error": "Mandatory field blank: 'Plant'",
                    })

        # ── Acct Assmt Grp Mat. (KTGRM) ──────────────────────────────────────
        aag_col = _ssa_pick(
            vbap,
            [
                "Acct Assmt Grp Mat.",
                "Account Assignment Group",
                "Acct Assignment Grp",
                "Acc.Assignment Grp.",
                "KTGRM",
            ],
        )

        if aag_col:
            # KTGRM must not be blank for active items belonging to any of
            # these Sales Document Types.
            mandatory_ktgrm_doc_types = {
                "ZLZ",
                "ZLZM",
                "ZRE",
                "ZPRO",
                "ZOR",
                "ZORM",
                "ZFD",
                "ZFD2",
                "ZRE2",
                "ZMD",
                "ZMD2",
                "ZMC",
                "ZMC2",
                "ZDID",
                "ZDIC",
            }

            # For these Sales Document Types, non-S materials require Z1.
            non_s_z1_doc_types = {
                "ZLZ",
                "ZLZM",
                "ZRE",
            }

            # For these Sales Document Types, the approved S-material mapping
            # must be validated.
            s_material_mapping_doc_types = {
                "ZPRO",
                "ZOR",
                "ZORM",
                "ZFD",
                "ZFD2",
                "ZRE2",
                "ZMD",
                "ZMD2",
                "ZMC",
                "ZMC2",
                "ZDID",
                "ZDIC",
            }

            # S-materials in scope for the MVKE-based KTGRM comparison.
            # The expected KTGRM per material is no longer hardcoded here —
            # it is looked up live from MVKE (field "Acct Assmt Grp Mat."),
            # see SSA_S_MATERIAL_LIST / _ssa_build_s_material_mvke_map.
            s_material_ktgrm_map = _ssa_build_s_material_mvke_map(
                mvke_df, SSA_S_MATERIAL_LIST
            )

            for _, row in vbap.iterrows():
                # The rule applies only when VBAP-ABGRU is blank.
                if is_rejected(row):
                    continue

                doc = (
                    _ssa_norm(row[doc_col])
                    if doc_col
                    else ""
                )
                item = (
                    _ssa_norm(row[item_col])
                    if item_col
                    else ""
                )
                material = (
                    _ssa_norm(row[mat_col])
                    if mat_col
                    else ""
                )
                plant = (
                    _ssa_norm(row[plt_col])
                    if plt_col
                    else ""
                )

                doc_type = doc_type_map.get(doc, "").upper()
                material_upper = material.upper()
                ktgrm = _ssa_norm(row[aag_col]).upper()

                # The special KTGRM rule applies only to the document types
                # explicitly listed in the business rule.
                if doc_type not in mandatory_ktgrm_doc_types:
                    continue

                # Rule 1:
                # KTGRM must not be blank for all listed document types.
                if ktgrm == "":
                    errors.append({
                        "Table": "VBAP",
                        "Sales Document": doc,
                        "Item": item,
                        "Material": material,
                        "Plant": plant,
                        "Field": "Acct Assmt Grp Mat.",
                        "Value": "",
                        "Error": (
                            "Mandatory field blank: "
                            "'Acct Assmt Grp Mat.' "
                            f"for active item with Sales Document Type "
                            f"'{doc_type}'"
                        ),
                    })
                    continue

                # Rule 2:
                # For ZLZ, ZLZM and ZRE, non-S materials require KTGRM Z1.
                if (
                    doc_type in non_s_z1_doc_types
                    and not material_upper.startswith("S-")
                ):
                    expected_ktgrm = "Z1"

                    if ktgrm != expected_ktgrm:
                        errors.append({
                            "Table": "VBAP",
                            "Sales Document": doc,
                            "Item": item,
                            "Material": material,
                            "Plant": plant,
                            "Field": "Acct Assmt Grp Mat.",
                            "Value": ktgrm,
                            "Error": (
                                f"Invalid KTGRM '{ktgrm}': "
                                f"Sales Document Type '{doc_type}' with "
                                f"non-S material '{material}' requires "
                                f"KTGRM '{expected_ktgrm}'"
                            ),
                        })

                    continue

                # Rule 3:
                # For the configured document types, validate S-materials
                # included in the approved mapping.
                if (
                    doc_type in s_material_mapping_doc_types
                    and material_upper.startswith("S-")
                ):
                    expected_ktgrm = s_material_ktgrm_map.get(
                        material_upper
                    )

                    # If the S-material is not included in the approved map,
                    # only the mandatory-field check applies. Since KTGRM is
                    # already known to be filled, no error is required.
                    if expected_ktgrm is None:
                        continue

                    if ktgrm != expected_ktgrm:
                        errors.append({
                            "Table": "VBAP",
                            "Sales Document": doc,
                            "Item": item,
                            "Material": material,
                            "Plant": plant,
                            "Field": "Acct Assmt Grp Mat.",
                            "Value": ktgrm,
                            "Error": (
                                f"Invalid KTGRM '{ktgrm}': "
                                f"Material '{material_upper}' requires "
                                f"KTGRM '{expected_ktgrm}' (per MVKE) for "
                                f"Sales Document Type '{doc_type}'"
                            ),
                        })

    return pd.DataFrame(errors)

def ssa_validar_terms_of_payment(vbak_df, vbkd_df=None):
    """
    Header-level validation:
      - Terms of Payment must not be empty.

    Source:
      - VBKD~Terms of Payment / ZTERM
      - Header-level VBKD record:
            POSNR blank or numerically equal to zero

    Note:
      VBAK-ZTERM is not used by this check; only VBKD is inspected.
      vbak_df is still required to know which Sales Documents to check.

    Performance:
      - Builds the VBKD lookup once.
      - Does not filter or copy the complete VBKD DataFrame for every
        Sales Document in VBAK.
    """
    errors = []

    empty_cols = [
        "Table",
        "Sales Document",
        "Item",
        "Material",
        "Plant",
        "Field",
        "Value",
        "Error",
    ]

    if vbak_df is None or vbak_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbak = _ssa_load(vbak_df)

    vbak_doc = _ssa_pick(
        vbak,
        [
            "Sales Document",
            "Sales Doc.",
            "VBELN",
        ],
    )

    if vbak_doc is None:
        return pd.DataFrame(columns=empty_cols)

    # Normalize the relevant VBAK values once.
    vbak_docs = vbak[vbak_doc].map(_ssa_norm)

    # -------------------------------------------------------------------------
    # Build the VBKD header-level Terms of Payment lookup once
    # -------------------------------------------------------------------------
    vbkd_terms_map = {}

    if vbkd_df is not None and not vbkd_df.empty:
        vbkd = _ssa_load(vbkd_df)

        vbkd_doc = _ssa_pick(
            vbkd,
            [
                "SD Document",
                "Sales Document",
                "Sales Doc.",
                "VBELN",
            ],
        )

        vbkd_item = _ssa_pick(
            vbkd,
            [
                "Item (SD)",
                "Sales Document Item",
                "POSNR",
                "Item",
            ],
        )

        vbkd_top = _ssa_pick(
            vbkd,
            [
                "Terms of Payment",
                "Terms of Payment Key",
                "ZTERM",
            ],
        )

        if vbkd_doc is not None and vbkd_top is not None:
            required_columns = [vbkd_doc, vbkd_top]

            if vbkd_item is not None:
                required_columns.append(vbkd_item)

            # Prevent duplicated column selection if aliases resolve
            # to the same physical column.
            required_columns = list(dict.fromkeys(required_columns))

            vbkd_work = vbkd.loc[:, required_columns].copy()

            vbkd_work["_ssa_doc"] = vbkd_work[vbkd_doc].map(_ssa_norm)
            vbkd_work["_ssa_terms"] = vbkd_work[vbkd_top].map(_ssa_norm)

            if vbkd_item is not None:
                vbkd_work["_ssa_item"] = vbkd_work[vbkd_item].map(_ssa_norm)

                # SAP item numbers can be imported from Excel in several forms:
                # '', 0, 0.0, 000000, etc.
                header_mask = vbkd_work["_ssa_item"].apply(
                    _ssa_is_zero_or_blank
                )

                vbkd_work = vbkd_work.loc[header_mask]

            # Remove rows that cannot contribute to the fallback lookup.
            vbkd_work = vbkd_work.loc[
                (vbkd_work["_ssa_doc"] != "")
                & (vbkd_work["_ssa_terms"] != "")
            ]

            if not vbkd_work.empty:
                # Preserve the original functional behavior:
                # use the first non-empty Terms of Payment found for
                # each Sales Document at header level.
                vbkd_terms_map = (
                    vbkd_work
                    .drop_duplicates(
                        subset=["_ssa_doc"],
                        keep="first",
                    )
                    .set_index("_ssa_doc")["_ssa_terms"]
                    .to_dict()
                )

    # -------------------------------------------------------------------------
    # Validate — one Sales Document per VBAK row, value sourced from VBKD only
    # -------------------------------------------------------------------------
    for row_index in vbak.index:
        doc = vbak_docs.at[row_index]
        value = vbkd_terms_map.get(doc, "")

        if value == "":
            errors.append({
                "Table": "VBKD",
                "Sales Document": doc,
                "Item": "",
                "Material": "",
                "Plant": "",
                "Field": "Terms of Payment",
                "Value": "",
                "Error": "Mandatory field blank: 'Terms of Payment'",
            })

    if not errors:
        return pd.DataFrame(columns=empty_cols)

    return pd.DataFrame(errors, columns=empty_cols)


def ssa_validar_shipping_point(vbap_df):
    """
    Item-level (only for Sales Document Types in SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
    - Shipping Point must not be empty
    - Shipping Point must be equal to Plant
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbap_df is None or vbap_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbap = _ssa_load(vbap_df)

    doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "POSNR"])
    mat_col  = _ssa_pick(vbap, ["Material", "MATNR"])
    plt_col  = _ssa_pick(vbap, ["Plant", "WERKS"])
    shp_col  = _ssa_pick(vbap, ["Shipping Point/Receiving Pt", "Shipping Point", "VSTEL"])
    status_col = _ssa_pick(vbap, ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"])
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

    if not all([doc_col, item_col, mat_col, plt_col, shp_col, status_col]):
        return pd.DataFrame(columns=empty_cols)

    for _, row in vbap.iterrows():
        if _ssa_item_is_done(row, status_col, rej_col):
            continue

        doc = _ssa_norm(row[doc_col])
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
            continue

        item  = _ssa_norm(row[item_col])
        mat   = _ssa_norm(row[mat_col])
        plant = _ssa_norm(row[plt_col])
        shp   = _ssa_norm(row[shp_col])

        if shp == '':
            errors.append({
                "Table": "VBAP",
                "Sales Document": doc,
                "Item": item,
                "Material": mat,
                "Plant": plant,
                "Field": "Shipping Point",
                "Value": "",
                "Error": "Mandatory field blank: 'Shipping Point'",
            })
        elif plant != '' and shp != plant:
            errors.append({
                "Table": "VBAP",
                "Sales Document": doc,
                "Item": item,
                "Material": mat,
                "Plant": plant,
                "Field": "Shipping Point",
                "Value": shp,
                "Error": f"Shipping Point must be equal to Plant (current Shipping Point='{shp}', Plant='{plant}')",
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


def ssa_validar_unloading_point_consistency(vbpa_df, vbak_df=None):
    """
    Header vs Item (only for Sales Document Types in
    SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP, and further restricted to Sales
    Document Type 'ZTAM' or 'ZORA'):
    - Unloading Point from header must equal Unloading Point from item.
    - Uses VBPA partner function 'SH'.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbpa_df is None or vbpa_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbpa = _ssa_load(vbpa_df)
    doc_type_map = _ssa_doc_type_map(vbak_df)
    allowed_types = {"ZTAM", "ZORA"}

    doc_col = _ssa_pick(vbpa, ["SD Document", "VBELN"])
    item_col = _ssa_pick(vbpa, ["Item (SD)", "POSNR"])
    pf_col   = _ssa_pick(vbpa, ["Partner Function", "PARVW"])
    up_col   = _ssa_pick(vbpa, ["Unloading Point"])

    if not all([doc_col, item_col, pf_col, up_col]):
        return pd.DataFrame(columns=empty_cols)

    sh = vbpa[vbpa[pf_col].astype(str).str.strip().str.upper() == "SH"].copy()
    if sh.empty:
        return pd.DataFrame(columns=empty_cols)

    for doc, grp in sh.groupby(sh[doc_col].astype(str).str.strip()):
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
            continue
        # If VBAK/doc-type info is available, also require ZTAM/ZORA
        if doc_type_map and doc_type_map.get(doc, '') not in allowed_types:
            continue

        hdr_rows = grp[grp[item_col].astype(str).str.strip().isin(['0', '', '000000', '0000000000'])]
        hdr_val = _ssa_first_nonempty(*hdr_rows[up_col].tolist()) if not hdr_rows.empty else ''

        item_rows = grp[~grp[item_col].astype(str).str.strip().isin(['0', '', '000000', '0000000000'])]
        if item_rows.empty:
            continue

        for _, row in item_rows.iterrows():
            item = _ssa_norm(row[item_col])
            item_val = _ssa_norm(row[up_col])

            if hdr_val != '' and item_val != '' and hdr_val != item_val:
                errors.append({
                    "Table": "VBPA",
                    "Sales Document": doc,
                    "Item": item,
                    "Material": "",
                    "Plant": "",
                    "Field": "Unloading Point",
                    "Value": item_val,
                    "Error": f"Unloading Point mismatch: header='{hdr_val}' vs item='{item_val}'",
                })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


def ssa_validar_customer_reference(vbap_df, vbkd_df, vbak_df=None):
    """
    Item-level Customer Reference validation using VBKD-BSTKD.

    Resolution order:
      1. Item-specific VBKD row for the same VBELN/POSNR.
      2. Header-level VBKD row where POSNR is blank or zero.

    Rejected VBAP items are skipped. The rule applies to Sales Document
    Types in SSA_ORDER_TYPES_ZLZ_GROUP (ZLZ, ZLZM), and to those in
    SSA_ORDER_TYPES_ZTAM_GROUP except ZPRO, ZOR, ZFD, ZMC, ZMD, ZDIC, ZDID
    and ZRE2.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbap_df is None or vbap_df.empty or vbkd_df is None or vbkd_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbap = _ssa_load(vbap_df)
    vbkd = _ssa_load(vbkd_df)
    doc_type_map = _ssa_doc_type_map(vbak_df)
    exclude_types = {"ZPRO", "ZOR", "ZFD", "ZMC", "ZMD", "ZDIC", "ZDID", "ZRE2"}

    ap_doc = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    ap_item = _ssa_pick(vbap, ["Sales Document Item", "Item (SD)", "POSNR"])
    ap_mat = _ssa_pick(vbap, ["Material", "MATNR"])
    ap_plant = _ssa_pick(vbap, ["Plant", "WERKS"])
    ap_status = _ssa_pick(vbap, ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"])
    ap_rej = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

    kd_doc = _ssa_pick(vbkd, ["SD Document", "Sales Document", "VBELN"])
    kd_item = _ssa_pick(vbkd, ["Item (SD)", "Sales Document Item", "POSNR"])
    kd_ref = _ssa_pick(vbkd, [
        "Customer Reference", "Customer Purchase Order Number",
        "Customer PO Number", "Cust. Reference", "BSTKD",
    ])

    if not all([ap_doc, ap_item, ap_status, kd_doc, kd_ref]):
        return pd.DataFrame(columns=empty_cols)

    def norm_item(value):
        raw = _ssa_norm(value)
        if _ssa_is_zero_or_blank(raw):
            return ""
        try:
            number = float(raw.replace(',', '.'))
            if number.is_integer():
                return str(int(number))
        except Exception:
            pass
        stripped = raw.lstrip('0')
        return stripped if stripped else ""

    item_map = {}
    header_map = {}
    for _, row in vbkd.iterrows():
        doc = _ssa_norm(row[kd_doc])
        if not doc:
            continue
        item = norm_item(row[kd_item]) if kd_item else ""
        reference = _ssa_norm(row[kd_ref])
        target = header_map if item == "" else item_map
        key = doc if item == "" else (doc, item)
        if key not in target or (target[key] == "" and reference != ""):
            target[key] = reference

    for _, row in vbap.iterrows():
        if _ssa_item_is_done(row, ap_status, ap_rej):
            continue

        doc = _ssa_norm(row[ap_doc])
        doc_type = doc_type_map.get(doc, "").upper()
        applies = (
            _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_GROUP)
            or (
                _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZTAM_GROUP)
                and doc_type not in exclude_types
            )
        )
        if not applies:
            continue

        item_display = _ssa_norm(row[ap_item])
        item_key = norm_item(row[ap_item])
        reference = _ssa_first_nonempty(
            item_map.get((doc, item_key), ""),
            header_map.get(doc, ""),
        )
        if reference == "":
            errors.append({
                "Table": "VBKD",
                "Sales Document": doc,
                "Item": item_display,
                "Material": _ssa_norm(row[ap_mat]) if ap_mat else "",
                "Plant": _ssa_norm(row[ap_plant]) if ap_plant else "",
                "Field": "Customer Reference",
                "Value": "",
                "Error": "Mandatory field blank: 'Customer Reference' (VBKD-BSTKD blank at item and header level)",
            })

    return pd.DataFrame(errors, columns=empty_cols)

def _ssa_item_key(value):
    """Normalize SAP item numbers so 10, 10.0 and 000010 match."""
    raw = _ssa_norm(value)
    if raw == "":
        return ""
    try:
        number = float(raw.replace(",", "."))
        if number.is_integer():
            return str(int(number))
    except (TypeError, ValueError):
        pass
    stripped = raw.lstrip("0")
    return stripped if stripped else "0"


def _ssa_build_active_demand_set(vbep_df):
    """
    Build {(Sales Document, Item)} for VBEP rows with Release Type = '1'.

    VBAP and VBEP are matched by VBELN + POSNR. Item numbers are normalized
    so Excel representations such as 10, 10.0 and 000010 resolve to the same
    key. A VBEP item is considered to have active demand when at least one
    matching row has Release Type/ABART equal to '1'.

    Returns (active_demand, ready):
      active_demand: set of normalized (document, item) pairs.
      ready: True only when VBEP and all required columns are available.
    """
    if vbep_df is None or vbep_df.empty:
        return set(), False

    vbep = _ssa_load(vbep_df)
    doc_col = _ssa_pick(vbep, ["Sales Document", "SD Document", "VBELN"])
    item_col = _ssa_pick(vbep, ["Sales Document Item", "Item (SD)", "POSNR"])
    release_col = _ssa_pick(vbep, ["Release Type", "ABART"])

    if not all([doc_col, item_col, release_col]):
        return set(), False

    active_demand = set()
    for _, row in vbep.iterrows():
        if _ssa_norm(row[release_col]).upper() != "1":
            continue
        doc = _ssa_norm(row[doc_col])
        item = _ssa_item_key(row[item_col])
        if doc and item:
            active_demand.add((doc, item))

    return active_demand, True


def ssa_validar_weights(vbap_df, vbfa_df=None, vbep_df=None):
    """
    Item-level weight validation for document types in
    SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP.

    Weight checks:
      - Gross Weight and Net Weight must not be blank or zero.
      - Gross Weight must equal Net Weight when both values are present.

    Demand gating for ZLZ/ZLZM:
      - Match VBAP to VBEP by Sales Document (VBELN) and Item (POSNR).
      - Validate weights only when at least one matching VBEP row has
        Release Type/ABART = '1'.
      - If no matching VBEP row exists, or all matching rows carry another
        Release Type, active demand is absent and weight validation is skipped.
      - If VBEP is not supplied or its required columns are unavailable, the
        ZLZ/ZLZM weight check is skipped and the dependency is exposed in the
        Field Diagnostics output.

    OBD gating for the remaining applicable document types:
      - If usable VBFA data is supplied, validate weights only after an
        Outbound Delivery exists for the item.
      - VBFA does not add an additional gate to ZLZ/ZLZM because those types
        are gated by active demand in VBEP.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbap_df is None or vbap_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbap = _ssa_load(vbap_df)
    obd_map, gate_on_obd = _ssa_build_obd_set(vbfa_df)
    active_demand, vbep_ready = _ssa_build_active_demand_set(vbep_df)

    doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "Item (SD)", "POSNR"])
    mat_col = _ssa_pick(vbap, ["Material", "MATNR"])
    plt_col = _ssa_pick(vbap, ["Plant", "WERKS"])
    gwt_col = _ssa_pick(vbap, ["Gross Weight", "BRGEW"])
    nwt_col = _ssa_pick(vbap, ["Net Weight", "NTGEW"])
    status_col = _ssa_pick(
        vbap,
        ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"],
    )
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

    if not all([doc_col, item_col, mat_col, plt_col, gwt_col, nwt_col, status_col]):
        return pd.DataFrame(columns=empty_cols)

    for _, row in vbap.iterrows():
        if _ssa_item_is_done(row, status_col, rej_col):
            continue

        doc = _ssa_norm(row[doc_col])
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
            continue

        item = _ssa_norm(row[item_col])
        item_key = _ssa_item_key(row[item_col])
        doc_type = _SSA_ACTIVE_DOC_TYPE_MAP.get(doc, "").upper()

        if doc_type in SSA_ORDER_TYPES_ZLZ_GROUP:
            # ZLZ/ZLZM: weights are mandatory only when active demand exists.
            if not vbep_ready or (doc, item_key) not in active_demand:
                continue
        elif gate_on_obd and not _ssa_has_obd(obd_map, doc, item):
            # Other applicable types retain the existing VBFA/OBD gate.
            continue

        mat = _ssa_norm(row[mat_col])
        plant = _ssa_norm(row[plt_col])
        gwt = row[gwt_col]
        nwt = row[nwt_col]

        if _ssa_is_zero_or_blank(gwt):
            errors.append({
                "Table": "VBAP",
                "Sales Document": doc,
                "Item": item,
                "Material": mat,
                "Plant": plant,
                "Field": "Gross Weight",
                "Value": _ssa_norm(gwt),
                "Error": "Gross Weight must not be blank or zero: active demand exists in VBEP (Release Type='1')"
                    if doc_type in SSA_ORDER_TYPES_ZLZ_GROUP
                    else "Gross Weight must not be blank or zero",
            })

        if _ssa_is_zero_or_blank(nwt):
            errors.append({
                "Table": "VBAP",
                "Sales Document": doc,
                "Item": item,
                "Material": mat,
                "Plant": plant,
                "Field": "Net Weight",
                "Value": _ssa_norm(nwt),
                "Error": "Net Weight must not be blank or zero: active demand exists in VBEP (Release Type='1')"
                    if doc_type in SSA_ORDER_TYPES_ZLZ_GROUP
                    else "Net Weight must not be blank or zero",
            })

        if (
            not _ssa_is_zero_or_blank(gwt)
            and not _ssa_is_zero_or_blank(nwt)
            and not _ssa_float_equal(gwt, nwt)
        ):
            errors.append({
                "Table": "VBAP",
                "Sales Document": doc,
                "Item": item,
                "Material": mat,
                "Plant": plant,
                "Field": "Gross Weight / Net Weight",
                "Value": f"Gross={_ssa_norm(gwt)} | Net={_ssa_norm(nwt)}",
                "Error": "Gross Weight and Net Weight must have the same value",
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


# -----------------------------------------------------------------------------
# Rule 2 — Incoterms (mandatory) + Incoterms DDP + Country GB
# -----------------------------------------------------------------------------
def ssa_validar_incoterms_ddp_gb(vbak_df, vbkd_df, kna1_df, vbpa_df=None, vbap_df=None):
    """
    New check — Incoterms (VBKD) must not be blank, for Sales Document
    Types in SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP.

    Existing check — if Incoterms (VBKD) = 'DDP' AND Customer Country = 'GB':
      - VBAK Alt.Tax Classific. must be '1'
      - VBAK Tax Departure C/R  must be 'GB'

    Country resolved via:
      Primary:  VBAK~Sold-to Party → KNA1~Customer → KNA1~Country/Region Key
      Optional: VBPA (Partner Function = 'SP') → Country/Region Key

    Exclude Rejected Documents — if vbap_df is supplied, Sales Documents
    where EVERY VBAP item is 'done' (Overall Status = C, or Reason for
    Rejection/VBAP-ABGRU filled) are skipped entirely (neither the
    Incoterms-blank check nor the DDP+GB check apply).
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant",
                  "Field", "Value", "Error"]

    if any(df is None or df.empty for df in [vbak_df, vbkd_df, kna1_df]):
        return pd.DataFrame(columns=empty_cols)

    vbak = _ssa_load(vbak_df)
    vbkd = _ssa_load(vbkd_df)
    kna1 = _ssa_load(kna1_df)
    fully_rejected_docs = _ssa_fully_rejected_docs(vbap_df)

    # Column picks
    vbak_doc  = _ssa_pick(vbak, ["Sales Document", "VBELN"])
    vbak_stp  = _ssa_pick(vbak, ["Sold-to Party", "KUNNR"])
    vbak_atc  = _ssa_pick(vbak, ["Alt.Tax Classific.", "TAXK1"])
    vbak_tdc  = _ssa_pick(vbak, ["Tax Departure C/R", "STCEG_L"])
    vbkd_doc  = _ssa_pick(vbkd, ["SD Document", "VBELN"])
    vbkd_inc  = _ssa_pick(vbkd, ["Incoterms", "INCO1"])
    kna1_cust = _ssa_pick(kna1, ["Customer", "KUNNR"])
    kna1_ctry = _ssa_pick(kna1, ["Country/Region Key", "LAND1"])

    if not all([vbak_doc, vbak_stp, vbak_atc, vbak_tdc,
                vbkd_doc, vbkd_inc, kna1_cust, kna1_ctry]):
        return pd.DataFrame(columns=empty_cols)

    # Build KNA1 map: customer -> country
    kna1_map = dict(zip(
        kna1[kna1_cust].astype(str).str.strip(),
        kna1[kna1_ctry].astype(str).str.strip().str.upper()
    ))

    # Optional: VBPA country override (Partner Function = SP)
    vbpa_map = {}
    if vbpa_df is not None and not vbpa_df.empty:
        vbpa = _ssa_load(vbpa_df)
        vp_doc = _ssa_pick(vbpa, ["SD Document", "VBELN"])
        vp_pf  = _ssa_pick(vbpa, ["Partner Function", "PARVW"])
        vp_cty = _ssa_pick(vbpa, ["Country/Region Key", "LAND1"])
        if all([vp_doc, vp_pf, vp_cty]):
            sp_rows = vbpa[vbpa[vp_pf].astype(str).str.strip().str.upper() == 'SP']
            for _, r in sp_rows.iterrows():
                vbpa_map[_ssa_norm(r[vp_doc])] = _ssa_norm(r[vp_cty]).upper()

    # Build VBKD map: sales doc -> incoterms (raw, for the blank check)
    vbkd_inc_map = {}
    for _, r in vbkd.iterrows():
        doc = _ssa_norm(r[vbkd_doc])
        inc = _ssa_norm(r[vbkd_inc])
        vbkd_inc_map[doc] = inc

    # Incoterms must not be blank, for Sales Document Types in
    # SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP.
    for _, row in vbak.iterrows():
        doc = _ssa_norm(row[vbak_doc])
        if doc in fully_rejected_docs:
            continue
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
            continue
        inc_val = vbkd_inc_map.get(doc, '')
        if inc_val == '':
            errors.append({
                "Table": "VBKD", "Sales Document": doc, "Item": '', "Material": '', "Plant": '',
                "Field": "Incoterms",
                "Value": '',
                "Error": "Mandatory field blank: 'Incoterms'",
            })

    # Evaluate each VBAK row — DDP + GB
    for _, row in vbak.iterrows():
        doc      = _ssa_norm(row[vbak_doc])
        if doc in fully_rejected_docs:
            continue
        stp      = _ssa_norm(row[vbak_stp])
        incoterm = vbkd_inc_map.get(doc, '').upper()
        if incoterm != 'DDP':
            continue

        # Resolve country — VBPA override first, then KNA1
        country = vbpa_map.get(doc) or kna1_map.get(stp, '')
        if country != 'GB':
            continue

        # DDP + GB → check Alt.Tax Classific.
        atc_val = _ssa_norm(row[vbak_atc])
        if atc_val != '1':
            errors.append({
                "Table": "VBAK", "Sales Document": doc, "Item": '', "Material": '', "Plant": '',
                "Field": "Alt.Tax Classific.",
                "Value": atc_val,
                "Error": (f"Incoterms=DDP and Country=GB: "
                          f"Alt.Tax Classific. must be '1' (current: '{atc_val}')"),
            })

        # DDP + GB → check Tax Departure C/R
        tdc_val = _ssa_norm(row[vbak_tdc])
        if tdc_val.upper() != 'GB':
            errors.append({
                "Table": "VBAK", "Sales Document": doc, "Item": '', "Material": '', "Plant": '',
                "Field": "Tax Departure C/R",
                "Value": tdc_val,
                "Error": (f"Incoterms=DDP and Country=GB: "
                          f"Tax Departure C/R must be 'GB' (current: '{tdc_val}')"),
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


# -----------------------------------------------------------------------------
# Rule 3 — Incompletion Log
# -----------------------------------------------------------------------------
def ssa_validar_incompletion_log(vbak_df, vbap_df):
    """
    Incompletion Log checks (rejected items are skipped):

    A) VBAK Overall Header (FA) must not be 'A'
    B) VBAP Item (JC) must not be 'A'

    Exception:
    - VBAP items belonging to order type ZTAM are ignored when
      Order Quantity (KWMENG) is zero.
    """
    errors = []
    completed_docs = _ssa_fully_rejected_docs(vbap_df)

    # Store Sales Documents whose order type is ZTAM.
    ztam_docs = set()

    # ── VBAK header check ─────────────────────────────────────────────────────
    if vbak_df is not None and not vbak_df.empty:
        vbak = _ssa_load(vbak_df)

        doc_col = _ssa_pick(
            vbak,
            [
                "Sales Document",
                "VBELN",
            ],
        )

        order_type_col = _ssa_pick(
            vbak,
            [
                "Sales Document Type",
                "Order Type",
                "Document Type",
                "AUART",
            ],
        )

        oh_col = _ssa_pick(
            vbak,
            [
                "Overall Header",
                "UVALL",
            ],
        )

        # Identify all ZTAM Sales Documents.
        if doc_col and order_type_col:
            for _, row in vbak.iterrows():
                doc = _ssa_norm(row[doc_col])
                order_type = _ssa_norm(row[order_type_col]).upper()

                if order_type == "ZTAM":
                    ztam_docs.add(doc)

        if oh_col:
            for _, row in vbak.iterrows():
                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                if doc in completed_docs:
                    continue

                val = _ssa_norm(row[oh_col]).upper()

                if val == "A":
                    errors.append({
                        "Table": "VBAK",
                        "Sales Document": doc,
                        "Item": "",
                        "Material": "",
                        "Plant": "",
                        "Field": "Overall Header",
                        "Value": val,
                        "Error": (
                            "Incompletion Log: Overall Header = 'A' "
                            "(document is incomplete)"
                        ),
                    })

    # ── VBAP item check (rejected items are skipped) ─────────────────────────
    if vbap_df is not None and not vbap_df.empty:
        vbap = _ssa_load(vbap_df)

        doc_col = _ssa_pick(
            vbap,
            [
                "Sales Document",
                "VBELN",
            ],
        )

        item_col = _ssa_pick(
            vbap,
            [
                "Sales Document Item",
                "POSNR",
            ],
        )

        mat_col = _ssa_pick(
            vbap,
            [
                "Material",
                "MATNR",
            ],
        )

        plt_col = _ssa_pick(
            vbap,
            [
                "Plant",
                "WERKS",
            ],
        )

        qty_col = _ssa_pick(
            vbap,
            [
                "Order Quantity",
                "KWMENG",
            ],
        )

        status_col = _ssa_pick(
            vbap,
            [
                "Overall Status",
                "Overall SD Status",
                "Overall Processing Status",
                "GBSTA",
                "GBSTK",
            ],
        )

        rej_col = _ssa_pick(
            vbap,
            [
                "Reason for Rejection",
                "ABGRU",
            ],
        )

        itm_col = _ssa_pick(
            vbap,
            [
                "UVPOS",
                "Item Incompletion Status",
                "Item Status",
            ],
        )

        if itm_col and status_col:
            for _, row in vbap.iterrows():
                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                # Completed or rejected items are skipped.
                if _ssa_item_is_done(row, status_col, rej_col):
                    continue

                # Exception:
                # Ignore ZTAM items when Order Quantity (KWMENG) is zero.
                if doc in ztam_docs and qty_col:
                    raw_qty = row[qty_col]

                    try:
                        if pd.isna(raw_qty):
                            qty = None
                        elif isinstance(raw_qty, str):
                            qty_text = raw_qty.strip().replace(" ", "")

                            # Supports values such as:
                            # 0
                            # 0.00
                            # 0,00
                            # 1,000.00
                            # 1.000,00
                            if "," in qty_text and "." in qty_text:
                                if qty_text.rfind(",") > qty_text.rfind("."):
                                    qty_text = (
                                        qty_text
                                        .replace(".", "")
                                        .replace(",", ".")
                                    )
                                else:
                                    qty_text = qty_text.replace(",", "")
                            elif "," in qty_text:
                                qty_text = qty_text.replace(",", ".")

                            qty = float(qty_text)
                        else:
                            qty = float(raw_qty)

                    except (TypeError, ValueError):
                        qty = None

                    if qty == 0:
                        continue

                val = _ssa_norm(row[itm_col]).upper()

                if val != "A":
                    continue

                errors.append({
                    "Table": "VBAP",
                    "Sales Document": doc,
                    "Item": (
                        _ssa_norm(row[item_col])
                        if item_col else ""
                    ),
                    "Material": (
                        _ssa_norm(row[mat_col])
                        if mat_col else ""
                    ),
                    "Plant": (
                        _ssa_norm(row[plt_col])
                        if plt_col else ""
                    ),
                    "Field": "Item",
                    "Value": val,
                    "Error": (
                        "Incompletion Log: Item status = 'A' "
                        "(item is incomplete)"
                    ),
                })

    return pd.DataFrame(errors)


# -----------------------------------------------------------------------------
# Rule 4 — Rejected items with Material Status 70/75/99
# -----------------------------------------------------------------------------
def ssa_validar_rejected_material_status(vbak_df, vbap_df, marc_df):
    """
    Report status 70/75/99 items whose VBAP Overall Status is not C.
    Items that are complete (Overall Status = 'C') or rejected
    (VBAP-ABGRU filled) are skipped.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]
    if vbap_df is None or vbap_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=empty_cols)
    vbap = _ssa_load(vbap_df)
    marc = _ssa_load(marc_df)
    doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "Item (SD)", "POSNR"])
    mat_col = _ssa_pick(vbap, ["Material", "MATNR"])
    plt_col = _ssa_pick(vbap, ["Plant", "WERKS"])
    overall_status_col = _ssa_pick(vbap, ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"])
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])
    marc_mat = _ssa_pick(marc, ["Material", "MATNR"])
    marc_plt = _ssa_pick(marc, ["Plant", "WERKS"])
    marc_psms = _ssa_pick(marc, ["Plant-Specific Material Status", "Plant Specific Material Status", "MTSTB"])
    if not all([doc_col, mat_col, plt_col, overall_status_col, marc_mat, marc_plt, marc_psms]):
        return pd.DataFrame(columns=empty_cols)
    def norm_psms(value):
        value = _ssa_norm(value)
        if not value:
            return ""
        try:
            return str(int(float(value.replace(",", "."))))
        except (TypeError, ValueError):
            return value.upper()
    marc_map = {}
    for _, r in marc.iterrows():
        marc_map[(_ssa_norm(r[marc_mat]), _ssa_norm(r[marc_plt]).upper())] = norm_psms(r[marc_psms])
    for _, row in vbap.iterrows():
        doc = _ssa_norm(row[doc_col])
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
            continue
        if _ssa_item_is_done(row, overall_status_col, rej_col):
            continue
        overall_status = _ssa_norm(row[overall_status_col]).upper()
        material = _ssa_norm(row[mat_col])
        plant = _ssa_norm(row[plt_col]).upper()
        material_status = marc_map.get((material, plant), "")
        if material_status not in {"70", "75", "99"}:
            continue
        errors.append({
            "Table": "VBAP", "Sales Document": doc,
            "Item": _ssa_norm(row[item_col]) if item_col else "",
            "Material": material, "Plant": plant, "Field": "Overall Status",
            "Value": overall_status,
            "Error": (f"Order is active: Material Status {material_status} found for "
                      f"Material={material}, Plant={plant}, but Overall Status is "
                      f"'{overall_status or '[blank]'}' instead of 'C'"),
        })
    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)

# -----------------------------------------------------------------------------
# Rule 10 — Sales Document without items in VBAP
# -----------------------------------------------------------------------------
def ssa_validar_order_without_items(vbak_df, vbap_df):
    """
    Flags any Sales Document (VBAK) that has no corresponding line items in VBAP.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbak_df is None or vbak_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbak = _ssa_load(vbak_df)
    vbak_doc_col = _ssa_pick(vbak, ["Sales Document", "VBELN"])
    if vbak_doc_col is None:
        return pd.DataFrame(columns=empty_cols)

    vbap_docs = set()
    if vbap_df is not None and not vbap_df.empty:
        vbap = _ssa_load(vbap_df)
        vbap_doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
        if vbap_doc_col:
            vbap_docs = set(vbap[vbap_doc_col].astype(str).str.strip())

    for _, row in vbak.iterrows():
        doc = _ssa_norm(row[vbak_doc_col])
        if doc == '':
            continue
        if doc not in vbap_docs:
            errors.append({
                "Table": "VBAK", "Sales Document": doc, "Item": '', "Material": '', "Plant": '',
                "Field": "Items",
                "Value": '',
                "Error": "Sales Document has no line items in VBAP",
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


# -----------------------------------------------------------------------------
# Rule 11 — Lines without Customer Material Number (KDMAT)
# -----------------------------------------------------------------------------
def ssa_validar_customer_material(vbap_df):
    """
    Item-level:
    - Customer Material Number (KDMAT) must not be empty (rejected items are skipped).
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbap_df is None or vbap_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbap = _ssa_load(vbap_df)

    doc_col  = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "POSNR"])
    mat_col  = _ssa_pick(vbap, ["Material", "MATNR"])
    plt_col  = _ssa_pick(vbap, ["Plant", "WERKS"])
    cmat_col = _ssa_pick(vbap, ["Customer Material Number", "Material by Customer", "KDMAT"])
    status_col = _ssa_pick(vbap, ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"])
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

    if not all([doc_col, item_col, mat_col, plt_col, cmat_col, status_col]):
        return pd.DataFrame(columns=empty_cols)

    for _, row in vbap.iterrows():
        if _ssa_item_is_done(row, status_col, rej_col):
            continue

        if _ssa_norm(row[cmat_col]) == '':
            errors.append({
                "Table": "VBAP",
                "Sales Document": _ssa_norm(row[doc_col]),
                "Item": _ssa_norm(row[item_col]),
                "Material": _ssa_norm(row[mat_col]),
                "Plant": _ssa_norm(row[plt_col]),
                "Field": "Customer Material Number",
                "Value": "",
                "Error": "Mandatory field blank: 'Customer Material Number'",
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


# -----------------------------------------------------------------------------
# Rule 12 — MRP for DlvSchType Type (VBAK~ABDIS)
# -----------------------------------------------------------------------------
def ssa_validar_mrp_for_ds(vbak_df):
    """
    Header-level (only for Sales Document Types in SSA_ORDER_TYPES_ZLZ_GROUP):
    - VBAK~ABDIS ("MRP for DlvSchType") must not be blank.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbak_df is None or vbak_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbak = _ssa_load(vbak_df)
    doc_col   = _ssa_pick(vbak, ["Sales Document", "VBELN"])
    abdis_col = _ssa_pick(vbak, ["MRP for DlvSchType", "MRPforDS", "ABDIS"])

    if not all([doc_col, abdis_col]):
        return pd.DataFrame(columns=empty_cols)

    for _, row in vbak.iterrows():
        doc = _ssa_norm(row[doc_col])
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_GROUP):
            continue
        if _ssa_norm(row[abdis_col]) == '':
            errors.append({
                "Table": "VBAK", "Sales Document": doc, "Item": '', "Material": '', "Plant": '',
                "Field": "MRP for DlvSchType",
                "Value": "",
                "Error": "Mandatory field blank: 'MRP for DlvSchType' (ABDIS)",
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


# -----------------------------------------------------------------------------
# Orchestrator — run all SSA validations
# -----------------------------------------------------------------------------
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
    global _SSA_ACTIVE_DOC_TYPE_MAP
    _SSA_ACTIVE_DOC_TYPE_MAP = _ssa_doc_type_map(vbak)

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

    out2 = io.BytesIO()
    wb.save(out2)
    out2.seek(0)
    return out2, results, summary_df
