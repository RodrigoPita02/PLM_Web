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
from rolesviews_data import get_rolesviews_bytesio


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
    """PP -> 'Yes' (1); all other Plant Types -> 'No' (0)."""
    if 'Plant Type' not in df.columns or 'Automatic Purchase Order Allowed' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'Automatic Purchase Order Allowed' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        plant_type = str(row['Plant Type']).strip()
        auto_order = str(row['Automatic Purchase Order Allowed']).strip()

        if pd.isna(row['Automatic Purchase Order Allowed']) or auto_order == '':
            erros.append('Mandatory field blank')
        else:
            if plant_type == 'PP' and auto_order != '1':
                erros.append(f"For Plant Type 'PP', should be 'Yes' (Current: '{auto_order}')")
            elif plant_type != 'PP' and auto_order != '0':
                erros.append(f"For Plant Type '{plant_type}', should be 'No' (Current: '{auto_order}')")

        if erros:
            resultados.append({
                'Line':     idx + 2,
                'Field':    'Automatic Purchase Order Allowed',
                'Error':    ' | '.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_batch_management(df, resultados):
    """PP / COI -> 'Yes' (1); all other Plant Types -> 'No' (0)."""
    if 'Plant Type' not in df.columns or 'Batch Management Required Ind.' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'Batch Management Required Ind.' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        plant_type = str(row['Plant Type']).strip()
        batch_ind  = str(row['Batch Management Required Ind.']).strip()

        if pd.isna(row['Batch Management Required Ind.']) or batch_ind == '':
            erros.append('Mandatory field blank')
        else:
            if plant_type in ('PP', 'COI') and batch_ind != '1':
                erros.append(f"For Plant Type 'PP', should be 'Yes' (Current: '{batch_ind}')")
            elif plant_type not in ('PP', 'COI') and batch_ind != '0':
                erros.append(f"For Plant Type '{plant_type}', should be 'No' (Current: '{batch_ind}')")

        if erros:
            resultados.append({
                'Line':     idx + 2,
                'Field':    'Batch Management Required Ind.',
                'Error':    ' | '.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant':    str(row.get('Plant', '')).strip(),
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
    Lot Sizing Procedure rules:
    - Mandatory.
    - EX: allowed only for FP or MP.
    - ZK: allowed only for PP.
    - Exception: MP material containing token '14K450' -> allows ZK or EX.
    """
    if 'Material Type' not in df.columns or 'Lot Sizing Procedure' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Lot Sizing Procedure' not found.")
        return

    token_pattern = re.compile(r'(?i)(?:^|[^A-Z0-9])14K450(?:[^A-Z0-9]|$)')

    def has_14k450(mat):
        return bool(token_pattern.search(str(mat).strip()))

    for idx, row in df.iterrows():
        erros = []
        plant_type  = str(row['Plant Type']).strip()  if 'Plant Type' in df.columns else ''
        lot_sizing  = '' if pd.isna(row['Lot Sizing Procedure']) else str(row['Lot Sizing Procedure']).strip()
        material    = str(row.get('Material', '')).strip()

        if lot_sizing == '':
            erros.append('Mandatory field blank')
        else:
            is_exception = (plant_type == 'MP' and has_14k450(material))
            if is_exception:
                if lot_sizing not in {'EX', 'ZK'}:
                    erros.append(
                        f"Invalid value: '{lot_sizing}' (for Plant Type MP with '14K450', allowed: ZK or EX)"
                    )
            else:
                if lot_sizing == 'EX' and plant_type not in ['FP', 'MP']:
                    erros.append(f"Value 'EX' only allowed for Plant Type FP or MP (current: '{plant_type}')")
                if lot_sizing == 'ZK' and plant_type != 'PP':
                    erros.append(f"Value 'ZK' only allowed for Plant Type PP (current: '{plant_type}')")

        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'Lot Sizing Procedure',
                'Error': ' \n'.join(erros),
                'Material': material,
                'Plant':    str(row.get('Plant', '')).strip(),
            })


def validar_per_lot_sz_plng_cal(marc_df, resultados, mara_df, matplsta_df=None):
    """
    Per.Lot Sz.Plng Cal. validation — sourced from MARC.

    Rules (scoped to FERT, HALB, ROH only):
      - FERT or HALB -> field is mandatory (must not be empty).
      - ROH         -> field must be empty
                       EXCEPTION: allowed when MATPLSTA has Planning Status = 9
                                  for the same Material + Plant.
    """
    field = 'Per.Lot Sz.Plng Cal.'

    if marc_df is None or marc_df.empty:
        print("⚠️ MARC DataFrame is required for Per.Lot Sz.Plng Cal. validation but was not provided.")
        return

    if field not in marc_df.columns:
        print(f"⚠️ Column '{field}' not found in the MARC sheet.")
        return

    if mara_df is None or mara_df.empty:
        print("⚠️ MARA DataFrame is required for Per.Lot Sz.Plng Cal. validation but was not provided.")
        return

    if 'Material' not in mara_df.columns or 'Material Type' not in mara_df.columns:
        print("⚠️ MARA is missing 'Material' or 'Material Type' columns required for Per.Lot Sz.Plng Cal. validation.")
        return

    # --- Material → Material Type map (igual ao original)
    mt_map = dict(zip(
        mara_df['Material'].astype(str).str.strip(),
        mara_df['Material Type'].astype(str).str.strip().str.upper()
    ))

    # --- NOVO: construir set de exceções ROH com Planning Status = 9
    roh_status_9 = set()
    if matplsta_df is not None and not matplsta_df.empty:
        matplsta_df = matplsta_df.copy()
        matplsta_df.columns = matplsta_df.columns.astype(str).str.strip()

        required_cols = {'Product Number', 'Location', 'Planning Status'}
        if required_cols.issubset(matplsta_df.columns):
            for _, r in matplsta_df.iterrows():
                if str(r['Planning Status']).strip() == '9':
                    roh_status_9.add((
                        str(r['Product Number']).strip(),
                        str(r['Location']).strip()
                    ))

    in_scope_types  = {'FERT', 'HALB', 'ROH'}
    mandatory_types = {'FERT', 'HALB'}

    for idx, row in marc_df.iterrows():
        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()
        raw      = row[field]
        val      = '' if pd.isna(raw) else str(raw).strip()

        mt = mt_map.get(material, '')

        if mt not in in_scope_types:
            continue

        # --- FERT / HALB (IGUAL ao original)
        if mt in mandatory_types:
            if val == '':
                resultados.append({
                    'Line':     idx + 2,
                    'Field':    field,
                    'Error':    f"Mandatory field blank for Material Type '{mt}' (FERT and HALB require a value)",
                    'Material': material,
                    'Plant':    plant,
                })

        # --- ROH (alteração mínima e correta)
        else:  # ROH
            if val != '':
                # 👉 SÓ AQUI se consulta a MATPLSTA
                if (material, plant) not in roh_status_9:
                    resultados.append({
                        'Line':     idx + 2,
                        'Field':    field,
                        'Error':    f"Field must be empty for Material Type 'ROH' (current: '{val}')",
                        'Material': material,
                        'Plant':    plant,
                    })


def validar_material_group(df, resultados):
    """Material Group is mandatory. Certain values are forbidden per Material Type."""
    if 'Material Type' not in df.columns or 'Material Group' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Material Group' not found.")
        return

    for idx, row in df.iterrows():
        erros = []

        material = str(row.get('Material', '')).strip()
        real_mt  = str(row.get('Material Type', '')).strip()

        # ✅ Material Type efetivo por regra
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
            if mt == 'FERT' and mg == 'L004':
                erros.append("For FERT, cannot be 'L004'")
            elif mt == 'HALB' and mg == 'L003':
                erros.append("For HALB, cannot be 'L003'")
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

    CSR exemption:
      If (Material, Plant) has an active CSR (status I1/I2/I4 in
      ZPLM_A_CSR_ITEM, matched via UUID to a Plant in ZPLM_A_CSR_HEAD),
      MRP = 'ND' is accepted for FP without transfer — no error raised.
      All other rules apply normally.

    PSMS precedence:
      10       -> skip entirely
      25       -> ND mandatory, TF=0/empty, Planning Calendar consistent
      70/75    -> ND mandatory
      65       -> ND allowed (if ND, skip base rules)

    Base rules (when no special PSMS):
      FP, no transfer, no CSR  -> M1, TF > 0
      FP, no transfer, CSR     -> M1 or ND accepted; TF check skipped when ND
      FP/MP, with transfer     -> M1, P1 or PD
      MP, no transfer          -> PD, TF = 0/empty
      PP + TF > 0              -> P1
      PP + TF = 0/empty        -> PD
    """
    needed = {'MRP Type', 'Plant Type', 'Planning Time Fence', 'Plant'}
    if not needed.issubset(df.columns):
        return

    has_psms = 'Plant Specific Material Status' in df.columns
    has_pc   = 'Planning Calendar' in df.columns
    valores_permitidos = {'M1', 'PD', 'P1', 'ND', 'V1'}

    transfer_map = {}
    if zplm_marc_df is not None:
        try:
            transfer_map = build_material_transfer_map(zplm_marc_df)
        except Exception:
            transfer_map = {}

    csr_exempt = build_csr_exempt_set(csr_item_df, csr_head_df)

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
        mrp   = '' if pd.isna(row['MRP Type'])    else str(row['MRP Type']).strip().upper()
        pt    = '' if pd.isna(row['Plant Type'])   else str(row['Plant Type']).strip().upper()
        plant = '' if pd.isna(row['Plant'])        else str(row['Plant']).strip()
        tf_raw = row.get('Planning Time Fence', None)
        tf_val = parse_tf(tf_raw)
        psms   = parse_status(row.get('Plant Specific Material Status', None)) if has_psms else None
        pc     = '' if (not has_pc or pd.isna(row.get('Planning Calendar'))) else str(row['Planning Calendar']).strip().upper()
        material    = str(row.get('Material', '')).strip()
        is_transfer = transfer_map.get((material, plant), False)
        has_csr     = (material, plant) in csr_exempt

        if psms == 10:
            continue

        if psms == 25:
            if mrp != 'ND':
                erros.append("Mandatory field blank (MRP Type)" if mrp == '' else
                             f"For Part Status = 25, MRP Type must be 'ND' (Current: '{mrp}')")
            if tf_val == 'INVALID':
                erros.append("Planning Time Fence invalid for Part Status = 25 (must be 0 or empty)")
            elif tf_val not in (None, 0):
                erros.append(f"For Part Status = 25, Planning Time Fence must be 0 or empty (Current: '{tf_raw}')")
            if has_pc:
                if plant == 'PT01' and pc != 'PT1':
                    erros.append(f"For Part Status = 25 & PT01, Planning Calendar must be 'PT1' (Current: '{pc}')")
                elif plant == 'PT02' and pc != 'PT2':
                    erros.append(f"For Part Status = 25 & PT02, Planning Calendar must be 'PT2' (Current: '{pc}')")
            if erros:
                resultados.append({'Line': idx + 2, 'Field': 'MRP Type',
                                   'Error': ' \n'.join(erros), 'Material': material, 'Plant': plant})
            continue

        if psms in (70, 75, 99):
            if mrp != 'ND':
                erros.append("Mandatory field blank (MRP Type)" if mrp == '' else
                            f"For Part Status = {psms}, MRP Type must be 'ND' (Current: '{mrp}')")
            if erros:
                resultados.append({'Line': idx + 2, 'Field': 'MRP Type',
                                'Error': ' \n'.join(erros), 'Material': material, 'Plant': plant})
            continue

        if psms == 65 and mrp == 'ND':
            continue

        # Base rules
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
                    pass  # ND accepted — active CSR exemption, TF check also skipped
                else:
                    # FIX: MRP and TF checks are now cleanly separated under the same else branch.
                    # Previously, the TF check used `not (has_csr and mrp == 'ND')` which was True
                    # for any ND without CSR, causing a redundant TF error on top of the MRP error.
                    if mrp != 'M1':
                        erros.append(
                            f"For FP, MRP must be 'M1' (Current: '{mrp}')"
                            + (" — note: ND would be accepted if an active CSR "
                               "(status I1/I2/I4) exists for this material/plant"
                               if mrp == 'ND' else "")
                        )
                    if not has_tf:
                        erros.append(f"For FP, Time Fence must be > 0 (Current: '{tf_raw}')")

            elif pt in {'FP', 'MP'} and is_transfer:
                if mrp not in {'M1', 'P1', 'PD'}:
                    erros.append(f"For {pt} with transfer, allowed MRP Types: M1, P1, PD (Current: '{mrp}')")
            elif pt == 'MP' and not is_transfer:
                if mrp != 'PD':
                    erros.append(f"For MP without transfer, MRP must be 'PD' (Current: '{mrp}')")
                if not no_tf:
                    erros.append(f"For MP without transfer, Time Fence must be 0 or empty (Current: '{tf_raw}')")
            elif pt == 'PP':
                if has_tf and mrp != 'P1':
                    erros.append(f"For PP with Time Fence > 0, MRP must be 'P1' (Current: '{mrp}')")
                elif no_tf and mrp != 'PD':
                    erros.append(f"For PP with Time Fence = 0/empty, MRP must be 'PD' (Current: '{mrp}')")

        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'MRP Type',
                               'Error': ' \n'.join(erros), 'Material': material, 'Plant': plant})


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


def validar_repetitive_manufacturing(df, resultados):
    """
    Repetitive Manufacturing Allow rules by Material Type and Plant Type:
    - ROH + MP/FP -> Yes; ROH + other -> No
    - HALB + PP   -> must not be Yes; HALB + other -> Yes
    - FERT + PP   -> must not be Yes; FERT + other -> Yes
    """
    if 'Material Type' not in df.columns or 'Indicator: Repetitive Manufa. Allow' not in df.columns or 'Plant Type' not in df.columns:
        print("⚠️ Required columns not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        mt    = str(row['Material Type']).strip()
        pt    = str(row['Plant Type']).strip()
        repet = '' if pd.isna(row['Indicator: Repetitive Manufa. Allow']) else str(row['Indicator: Repetitive Manufa. Allow']).strip()

        if repet == '':
            erros.append('Mandatory field blank')
        else:
            if mt == 'ROH':
                if pt in ['MP', 'FP'] and repet != '1':
                    erros.append(f"For ROH with Plant Type '{pt}', should be 'Yes' (Current: '{repet}')")
                elif pt not in ['MP', 'FP', ''] and repet != '0':
                    erros.append(f"For ROH with Plant Type '{pt}', should be 'No' (Current: '{repet}')")
            elif mt == 'HALB':
                if pt == 'PP' and repet == '1':
                    erros.append("For HALB with Plant Type 'PP', shouldn't be 'Yes'")
                elif pt not in ['PP', ''] and repet != '1':
                    erros.append(f"For HALB with Plant Type '{pt}', should be 'Yes' (Current: '{repet}')")
            elif mt == 'FERT':
                if pt == 'PP' and repet == '1':
                    erros.append("For FERT with Plant Type 'PP', shouldn't be 'Yes'")
                elif pt not in ['PP', ''] and repet != '1':
                    erros.append(f"For FERT with Plant Type '{pt}', should be 'Yes' (Current: '{repet}')")

        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Indicator: Repetitive Manufa. Allow',
                               'Error': ' \n '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def validar_repetitive_manufacturing_profile(df, resultados):
    """FERT and HALB -> 'Z002'; ROH -> must NOT be 'Z002'."""
    if 'Material Type' not in df.columns or 'Repetitive Manufacturing Profile' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Repetitive Manufacturing Profile' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        mt  = str(row['Material Type']).strip()
        val = str(row['Repetitive Manufacturing Profile']).strip()
        if mt in ['FERT', 'HALB']:
            if pd.isna(row['Repetitive Manufacturing Profile']) or val == '':
                erros.append("Mandatory field blank For Material Type FERT or HALB")
            elif val != 'Z002':
                erros.append(f"Invalid value: '{val}' (should be 'Z002' For FERT or HALB)")
        if mt == 'ROH' and val == 'Z002':
            erros.append("Material Type ROH cannot have value 'Z002'")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Repetitive Manufacturing Profile',
                               'Error': ' | '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def validar_source_list_required(df, resultados):
    """PP -> 'Yes' (1); all other Plant Types -> 'No' (0)."""
    if 'Plant Type' not in df.columns or 'Source List Required' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'Source List Required' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        pt  = str(row['Plant Type']).strip()
        val = str(row['Source List Required']).strip()
        if pd.isna(row['Source List Required']) or val == '':
            erros.append('Mandatory field blank')
        else:
            if pt == 'PP' and val != '1':
                erros.append(f"For Plant Type PP, should be 'Yes' (Current: '{val}')")
            elif pt != 'PP' and val != '0':
                erros.append(f"For Plant Type {pt}, should be 'No' (Current: '{val}')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Source List Required',
                               'Error': ' | '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


def validar_storage_location(df, resultados):
    """
    Issue Storage Location validation:

    PT01:
      - HALB / ROH -> must be 'P105'
      - FERT       -> must be 'P120'

    PT02:
      - Any Material Type -> must be 'P205' or 'P220'
    """
    required = {'Plant', 'Issue Storage Location', 'Material Type'}
    if not required.issubset(df.columns):
        missing = sorted(required - set(df.columns))
        print(f"⚠️ Missing columns for Issue Storage Location validation: {missing}")
        return

    for idx, row in df.iterrows():
        erros = []
        plant = str(row['Plant']).strip().upper()
        mt    = str(row['Material Type']).strip().upper()
        raw   = row['Issue Storage Location']
        val   = '' if pd.isna(raw) else str(raw).strip()
        material = str(row.get('Material', '')).strip()

        if val == '':
            erros.append('Mandatory field blank')
        elif plant == 'PT01':
            if mt in ('HALB', 'ROH'):
                if val != 'P105':
                    erros.append(
                        f"For PT01 + {mt}, Issue Storage Location must be 'P105' "
                        f"(Current: '{val}')"
                    )
            elif mt == 'FERT':
                if val != 'P120':
                    erros.append(
                        f"For PT01 + FERT, Issue Storage Location must be 'P120' "
                        f"(Current: '{val}')"
                    )
        elif plant == 'PT02':
            if val not in ('P205', 'P220'):
                erros.append(
                    f"For PT02, Issue Storage Location must be 'P205' or 'P220' "
                    f"(Current: '{val}')"
                )
            elif val == 'P220':
                # P220 only allowed for PSMS = 25 or 75
                psms_raw = row.get('Plant Specific Material Status', '')
                psms = '' if pd.isna(psms_raw) else str(psms_raw).strip()
                try:
                    psms = str(int(float(psms)))
                except Exception:
                    pass
                if psms not in ('25', '75'):
                    erros.append(
                        f"Value 'P220' only allowed for PT02 when Plant Specific Material Status "
                        f"is 25 or 75 (Current PSMS: '{psms}')"
                    )

        if erros:
            resultados.append({
                'Line':     idx + 2,
                'Field':    'Issue Storage Location',
                'Error':    ' | '.join(erros),
                'Material': material,
                'Plant':    plant,
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
    """With Qty Structure must always be '1' (Yes)."""
    if 'With Qty Structure' not in df.columns:
        print("⚠️ Column 'With Qty Structure' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        val = str(row['With Qty Structure']).strip()
        if pd.isna(row['With Qty Structure']) or val == '':
            erros.append('Mandatory field blank')
        elif val != '1':
            erros.append(f"Invalid value: '{val}' (should be 'Yes')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'With Qty Structure',
                               'Error': ' | '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


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


def validar_mrp_controller(df, resultados):
    """MRP Controller is mandatory.
       Z01 is not allowed when PSMS is different from 00, 10, 70, 75 or 99.
    """
    if 'MRP Controller' not in df.columns:
        return

    has_psms = 'Plant Specific Material Status' in df.columns
    allowed_z01_status = {0, 10, 70, 75, 99}

    def parse_status(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return int(float(str(val).strip()))
        except Exception:
            return None

    for idx, row in df.iterrows():
        erros = []

        mrp_ctrl = '' if pd.isna(row['MRP Controller']) else str(row['MRP Controller']).strip()
        psms = parse_status(row.get('Plant Specific Material Status')) if has_psms else None

        # Mandatory
        if mrp_ctrl == '':
            erros.append('Mandatory field blank')

        # Z01 restriction
        elif mrp_ctrl == 'Z01' and psms not in allowed_z01_status:
            erros.append(
                "MRP Controller 'Z01' not allowed when Material Status "
                "is different from 00, 10, 70, 75 or 99"
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
            for sorg, dchan in sorted(combos - FERT_REQUIRED):
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

        for _, r in pc_rows.iterrows():
            try:
                num = int(float(r.get('Numerator')))
                den = int(float(r.get('Denominator')))
            except Exception:
                num = den = None

            if num == 1 and den == 1:
                valid_pc_found = True
                break

            if num != 1:
                num_invalid = True
            if den != 1:
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

def validar_country_region_origin_fp(marc_df, zplm_marc_df, resultados):
    """
    SAP Material Validation:
    MARC - Country/Region of Origin is mandatory for:
    - Plant Type = FP (from ZPLM MARC)
    - PSMS in {25,30,35,40,45,50,55,65}
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

    if (
        marc_df is None or zplm_marc_df is None
        or marc_df.empty or zplm_marc_df.empty
        or not required_marc.issubset(marc_df.columns)
        or not required_zplm.issubset(zplm_marc_df.columns)
    ):
        return

    # ✅ status atualizados
    valid_status = {'25', '30', '35', '40', '45', '50', '55', '65'}

    # Join MARC with ZPLM MARC to get Plant Type
    df = marc_df.merge(
        zplm_marc_df[['Material', 'Plant', 'Plant Type']],
        on=['Material', 'Plant'],
        how='left'
    )

    for idx, row in df.iterrows():
        plant_type = str(row.get('Plant Type', '')).strip().upper()
        psms = normalize_code(
            row.get('Plant Specific Material Status', ''),
            as_int=True
        )
        origin = str(row.get('Cntry/Reg of Origin', '')).strip()

        if plant_type == 'FP' and psms in valid_status and origin == '':
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

def validar_storage_location_ext_procurement(marc_df, zplm_marc_df, resultados):
    """
    MARC:
    Storage Location for Ext. Procurement must be filled
    when Plant Type = PP.
    Plant Type is sourced from ZPLM MARC (source of truth).
    Validation per (Material, Plant).
    """
    required_marc = {
        "Material",
        "Plant",
        "Storage Location for Ext. Procurement",
    }
    required_zplm = {
        "Material",
        "Plant",
        "Plant Type",
    }

    if (
        marc_df is None
        or zplm_marc_df is None
        or not required_marc.issubset(marc_df.columns)
        or not required_zplm.issubset(zplm_marc_df.columns)
    ):
        return

    # Join MARC with ZPLM MARC to get Plant Type
    df = marc_df.merge(
        zplm_marc_df[["Material", "Plant", "Plant Type"]],
        on=["Material", "Plant"],
        how="left",
    )

    for idx, row in df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant = str(row.get("Plant", "")).strip()
        plant_type = str(row.get("Plant Type", "")).strip()
        sloc_ext = str(row.get("Storage Location for Ext. Procurement", "")).strip()

        if plant_type == "PP" and sloc_ext == "":
            resultados.append({
                "Line": idx + 2,
                "Field": "Storage Location for Ext. Procurement",
                "Error": "Mandatory field blank for Plant Type PP",
                "Material": material,
                "Plant": plant,
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

def validar_batch_management_plant_marc(marc_df, zplm_marc_df, resultados):
    """
    MARC:
    Batch Management (Plant) rules:
    - If Plant Type ∈ {PP, COI} → must be 'X'
    - It may be 'X' ONLY if Batch Management = 'X'
    Validation per (Material, Plant).
    """
    required_marc = {
        "Material",
        "Plant",
        "Batch Management(Plant)",
        "Batch Management",
    }
    required_zplm = {
        "Material",
        "Plant",
        "Plant Type",
    }

    if (
        marc_df is None
        or zplm_marc_df is None
        or not required_marc.issubset(marc_df.columns)
        or not required_zplm.issubset(zplm_marc_df.columns)
    ):
        return

    # Join MARC + ZPLM MARC to get Plant Type
    df = marc_df.merge(
        zplm_marc_df[["Material", "Plant", "Plant Type"]],
        on=["Material", "Plant"],
        how="left",
    )

    for idx, row in df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant = str(row.get("Plant", "")).strip()
        plant_type = str(row.get("Plant Type", "")).strip().upper()

        bmp = _norm_str(row.get("Batch Management(Plant)")).upper()
        bm  = _norm_str(row.get("Batch Management")).upper()

        erros = []

        if plant_type in {"PP", "COI"}:
            if bmp != "X":
                erros.append(
                    "Mandatory for Plant Type PP/COI (should be 'X')"
                )
        else:
            if bmp == "X":
                erros.append(
                    f"Not allowed for Plant Type '{plant_type}'"
                )

        # Dependency rule
        if bmp == "X" and bm != "X":
            erros.append(
                "Cannot be 'X' when Batch Management is not 'X'"
            )

        if erros:
            resultados.append({
                "Line": idx + 2,
                "Field": "Batch Management(Plant)",
                "Error": " | ".join(erros),
                "Material": material,
                "Plant": plant,
            })

def validar_version_indicator_marc(
    marc_df,
    zplm_marc_df,
    resultados,
):
    """
    MARC:
    Version Indicator rules:
    - Plant Type FP or MP  → must be 'X'
    - Any other Plant Type → must be EMPTY
    Validation per (Material, Plant).
    """
    required_marc = {
        "Material",
        "Plant",
        "Version Indicator",
    }
    required_zplm = {
        "Material",
        "Plant",
        "Plant Type",
    }

    if (
        marc_df is None
        or zplm_marc_df is None
        or not required_marc.issubset(marc_df.columns)
        or not required_zplm.issubset(zplm_marc_df.columns)
    ):
        return

    # Join MARC + ZPLM MARC to get Plant Type
    df = marc_df.merge(
        zplm_marc_df[["Material", "Plant", "Plant Type"]],
        on=["Material", "Plant"],
        how="left",
    )

    for idx, row in df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant = str(row.get("Plant", "")).strip()
        plant_type = str(row.get("Plant Type", "")).strip().upper()

        version_ind = _norm_str(row.get("Version Indicator")).upper()

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
    LTB fields validation (Data Cleanups):
    Only FP + MT∈{FERT,HALB,ROH} + Status∈{55,70,99} may have LTB values.
    Any other material with LTB values -> Type A error.
    """
    z = zplm_marc.copy()
    m = marc.copy()
    a = mara.copy()

    req_z = {'Material', 'Plant', 'Material Plant Type', 'LTB Quantity', 'LTB Start Date'}
    req_m = {'Material', 'Plant', 'Plant-Specific Material Status'}
    req_a = {'Material', 'Material Type'}

    missing = []
    if not req_z.issubset(z.columns): missing.append(f"ZPLM_MARC: {sorted(req_z - set(z.columns))}")
    if not req_m.issubset(m.columns): missing.append(f"MARC: {sorted(req_m - set(m.columns))}")
    if not req_a.issubset(a.columns): missing.append(f"MARA: {sorted(req_a - set(a.columns))}")

    cols = ['Material', 'Plant', 'Material Plant Type', 'Material Type',
            'Plant-Specific Material Status', 'LTB Quantity', 'LTB Start Date', 'Field', 'Error']

    if missing:
        return pd.DataFrame(columns=cols)

    a_use = a[['Material', 'Material Type']].copy()
    a_use['Material'] = a_use['Material'].astype(str).str.strip()
    a_use['Material Type'] = a_use['Material Type'].astype(str).str.strip().str.upper()

    m_use = m[['Material', 'Plant', 'Plant-Specific Material Status']].copy()
    m_use['Material'] = m_use['Material'].astype(str).str.strip()
    m_use['Plant'] = m_use['Plant'].astype(str).str.strip()
    status_num = pd.to_numeric(m_use['Plant-Specific Material Status'], errors='coerce')
    m_use['Plant-Specific Material Status'] = (
        status_num.astype('Int64').astype(str).str.replace('<NA>', '', regex=False).str.strip()
    )

    z['Material'] = z['Material'].astype(str).str.strip()
    z['Plant'] = z['Plant'].astype(str).str.strip()
    z = z.merge(a_use, on='Material', how='left').merge(m_use, on=['Material', 'Plant'], how='left')

    valid_mtypes = {'FERT', 'HALB', 'ROH'}
    valid_status = {'55', '70', '99'}

    is_fp       = z['Material Plant Type'].astype(str).str.upper().eq('FP')
    is_valid_mt = z['Material Type'].astype(str).str.upper().isin(valid_mtypes)
    is_valid_st = z['Plant-Specific Material Status'].astype(str).isin(valid_status)

    ltb_qty_raw = z['LTB Quantity']
    ltb_qty_num = pd.to_numeric(ltb_qty_raw, errors='coerce')
    is_qty_empty = ltb_qty_raw.isna() | ltb_qty_raw.astype(str).str.strip().isin(["", "0", "0.0"]) | (ltb_qty_num.fillna(0) == 0)

    ltb_date_raw = z['LTB Start Date']
    ltb_date_dt  = pd.to_datetime(ltb_date_raw, errors='coerce')
    is_date_empty = ltb_date_raw.isna() | ltb_date_raw.astype(str).str.strip().isin(["", "0", "0.0"]) | ltb_date_dt.isna()

    authorized = is_fp & is_valid_mt & is_valid_st
    mask_A = (~authorized) & (~is_qty_empty | ~is_date_empty)

    df_A = z.loc[mask_A, ['Material', 'Plant', 'Material Plant Type', 'Material Type',
                           'Plant-Specific Material Status', 'LTB Quantity', 'LTB Start Date']].copy()
    df_A['Field'] = 'LTB Quantity / LTB Start Date'
    df_A['Error'] = (
        "Material not authorized to have LTB values "
        "(requires FP + MT∈{FERT/HALB/ROH} + Status∈{55,70,99})"
    )

    if not df_A.empty:
        df_A.sort_values(by=['Plant', 'Material'], inplace=True)

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

    Rule:
    - Must be marked ONLY for Material Types: ROH and HALB
    - Must NOT be marked for any other Material Type
    - If blank for ROH/HALB -> mandatory blank error
    """
    if marc_df is None or marc_df.empty or mara_df is None or mara_df.empty:
        return []

    needed_marc = {'Material', 'Plant', 'Neg. stocks in plant'}
    needed_mara = {'Material', 'Material Type'}

    if not needed_marc.issubset(marc_df.columns) or not needed_mara.issubset(mara_df.columns):
        return []

    # Map Material -> Material Type
    mt_map = dict(
        zip(
            mara_df['Material'].astype(str).str.strip(),
            mara_df['Material Type'].astype(str).str.strip().str.upper()
        )
    )

    def _s(x):
        return '' if pd.isna(x) else str(x).strip()

    def _is_blank(x):
        return _s(x) == ''

    def _is_marked(x):
        return _s(x).upper() in {'X', '1', 'YES', 'Y', 'TRUE'}

    erros = []

    for idx, row in marc_df.iterrows():
        material = _s(row.get('Material'))
        plant    = _s(row.get('Plant'))
        mt       = mt_map.get(material, '')

        neg_raw = row.get('Neg. stocks in plant')
        blank   = _is_blank(neg_raw)
        marked  = _is_marked(neg_raw)

        must_mark = mt in {'ROH', 'HALB'}

        if must_mark and blank:
            erros.append({
                'Line': idx + 2,
                'Field': 'Neg. stocks in plant',
                'Error': f"Mandatory field blank — must be marked for Material Type {mt}",
                'Material': material,
                'Plant': plant
            })

        elif must_mark and not marked:
            erros.append({
                'Line': idx + 2,
                'Field': 'Neg. stocks in plant',
                'Error': f"Must be marked for Material Type {mt}",
                'Material': material,
                'Plant': plant
            })

        elif not must_mark and marked:
            erros.append({
                'Line': idx + 2,
                'Field': 'Neg. stocks in plant',
                'Error': "Invalid marking — allowed only for Material Types ROH or HALB",
                'Material': material,
                'Plant': plant
            })

    return erros

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
        'Repetitive Manufacturing Profile':      lambda mt, pt: mt in ['FERT', 'HALB'],
        'Indicator: Repetitive Manufa. Allow':   lambda mt, pt: True,
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
        'LIFO/FIFO-relevant':                    lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Stock Determination Group':             lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Storage Location for Ext. Procurement': lambda mt, pt: pt == 'PP',
        'Individual/ Collective Requirements':   lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Packaging Mat. Type':                   lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],   
        'Production Supervisor':                 lambda mt, pt: mt in ['FERT', 'HALB'],   
        'Batch Management(Plant)':               lambda mt, pt: pt in ['PP', 'COI'],
        'Advanced Planning':                     lambda mt, pt: True,
        'Heuristic':                             lambda mt, pt: True,
        'PP Plng Procedure':                     lambda mt, pt: True,
        'Requirement Strategy':                  lambda mt, pt: True,
        'Version Indicator':                     lambda mt, pt: pt in ['FP', 'MP'],
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
            view_tabs_dict[v] = pd.DataFrame(rows)

    return errors_df, errors_filtered_df, views_missing_output_df, contagem_final_df, coverage_df, view_tabs_dict


# =============================================================================
# 2) RUN FULL VALIDATION — Main SAPUI5 Pipeline
# =============================================================================

def run_full_validation(
    sap_file,
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
):
    """
    Runs the complete SAPUI5 validation pipeline:
      1. Load files
      2. Pre-normalisation
      3. All field validations
      4. Post-processing (Views / Roles / Summary)
      5. Excel generation in-memory
    Returns: (excel_bytes, errors_df, summary_df)
    """
    def _cb(pct, msg):
        if progress_callback:
            progress_callback(pct, msg)

    # ------------------------------------------------------------------
    # 1) Load DataFrames
    # ------------------------------------------------------------------
    _cb(5,  "Loading SAPUI5 Export...")
    sap_df = load_xlsx_fast(sap_file, sheet_name="SAPUI5 Export")

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

    # ------------------------------------------------------------------
    # 2) Basic sanity checks
    # ------------------------------------------------------------------
    if sap_df is None:
        raise ValueError("SAPUI5 Export.xlsx not loaded or missing sheet 'SAPUI5 Export'")
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

    # Apply plant filter to SAPUI5 Export
    if plant_filter is not None and 'Plant' in sap_df.columns:
        pf = str(plant_filter).strip().upper()
        sap_df = sap_df[sap_df['Plant'].astype(str).str.strip().str.upper() == pf].copy()

    # ------------------------------------------------------------------
    # 3) Pre-normalisation
    # ------------------------------------------------------------------
    marc_df = ensure_basic_data_k(marc_df)

    # ------------------------------------------------------------------
    # 4) Run all field validations
    # ------------------------------------------------------------------
    resultados = []

    # Mandatory fields
    for coluna in ['Material', 'Material Type', 'Description', 'Base UOM', 'Plant', 'Plant Type']:
        _cb(20, f"Validating mandatory field: {coluna}...")
        validar_campo_obrigatorio(sap_df, coluna, resultados)

    validations = [
        (22, lambda: validar_procurement_type(sap_df, resultados)),
        (24, lambda: validar_account_assignment_group(sap_df, resultados, mvke_df)),
        (26, lambda: validar_automatic_purchase_order(sap_df, resultados)),
        (28, lambda: validar_batch_management(sap_df, resultados)),
        (30, lambda: validar_consumption_mode(sap_df, resultados)),
        (32, lambda: validar_consumption_per(marc_df, mara_df, resultados)),
        (33, lambda: validar_division(sap_df, resultados)),
        (34, lambda: validar_export_import_group(sap_df, resultados)),
        (36, lambda: validar_general_item_category_group(sap_df, resultados)),
        (38, lambda: validar_gross_weight(sap_df, resultados)),
        (40, lambda: validar_item_category_group(sap_df, resultados, mvke_df)),
        (42, lambda: validar_jit_delivery_sched(sap_df, resultados)),
        (44, lambda: validar_loading_group(sap_df, resultados)),
        (46, lambda: validar_lot_sizing_procedure(sap_df, resultados)),
        (47, lambda: validar_per_lot_sz_plng_cal(marc_df, resultados, mara_df, matplsta_df)),
        (48, lambda: validar_material_group(sap_df, resultados)),
        (50, lambda: validar_mrp_type(sap_df, resultados, zplm_marc_df, csr_item_df, csr_head_df)),
        (51, lambda: validar_individual_collective_requirements_marc(marc_df, mara_df, resultados)),
        (52, lambda: validar_net_weight(sap_df, resultados)),
        (54, lambda: validar_period_type(sap_df, resultados)),
        (55, lambda: validar_storage_location_ext_procurement(marc_df, zplm_marc_df, resultados)),
        (56, lambda: validar_planning_calendar(sap_df, resultados)),
        (57, lambda: validar_stock_determination_group_marc(marc_df, mara_df, resultados)),
        (58, lambda: validar_planning_time_fence(sap_df, resultados)),
        (59, lambda: validar_version_indicator_marc(marc_df, zplm_marc_df, resultados)),
        (60, lambda: validar_price_control(sap_df, resultados)),
        (61, lambda: validar_batch_management_plant_marc(marc_df, zplm_marc_df, resultados)),
        (62, lambda: validar_price_unit(sap_df, resultados)),
        (63, lambda: validar_product_hierarchy(sap_df, resultados, mvke_df, mara_df)),
        (64, lambda: validar_production_supervisor_marc(marc_df, mara_df, resultados)),
        (66, lambda: validar_product_hierarchy_vs_enterproj(mara_df, zplm_mara_df, resultados)),
        (67, lambda: validar_packaging_material_type_mara(mara_df, resultados)),
        (68, lambda: validar_production_scheduling_profile(sap_df, resultados)),
        (69, lambda: validar_country_region_origin_fp(marc_df, zplm_marc_df, resultados)),
        (70, lambda: validar_profit_center(sap_df, resultados)),
        (71, lambda: validar_lifo_fifo_relevant_mbew(mbew_df, mara_df, resultados)),
        (72, lambda: validar_purchasing_group(sap_df, resultados)),
        (74, lambda: validar_repetitive_manufacturing(sap_df, resultados)),
        (76, lambda: validar_repetitive_manufacturing_profile(sap_df, resultados)),
        (78, lambda: validar_source_list_required(sap_df, resultados)),
        (80, lambda: validar_storage_location(sap_df, resultados)),
        (82, lambda: validar_valuation_class(sap_df, resultados)),
        (84, lambda: validar_with_qty_structure(sap_df, resultados)),
        (86, lambda: validar_enterproj_number(sap_df, resultados)),
        (87, lambda: validar_mrp_controller(sap_df, resultados)),
        (88, lambda: validar_plant_specific_material_status(sap_df, resultados)),
        (89, lambda: validar_backflush(sap_df, resultados)),
        (90, lambda: validar_strategy_group(sap_df, resultados)),
        (91, lambda: validar_product_related_origin(sap_df, resultados)),
        (92, lambda: validar_availability_check(sap_df, resultados)),
        (93, lambda: validar_material_group_appoint(sap_df, resultados)),
        (94, lambda: validar_purchasing_group_dxx(sap_df, resultados)),
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
        marc_df=marc_df,
        mara_df=mara_df,
        resultados=resultados,
    )

    # ------------------------------------------------------------------
    # 5) Build errors DataFrame
    # ------------------------------------------------------------------
    errors_df = _coerce_df_for_display(pd.DataFrame(resultados))

    # Enterproj enrichment (optional — needs previous report)
    if previous_report:
        errors_df = fill_plant_and_validate_enterproj_from_zplm(
            errors_df, latest_report=previous_report,
            zplm_marc_file=zplm_marc_file, zplm_mara_file=zplm_mara_file
        )

    # Additional enrichments
    if not errors_df.empty:
        errors_df = add_old_material_number_errors(mara_df, errors_df)
        neg_errors = validar_neg_stocks_in_plant(marc_df, mara_df)
        if neg_errors:
            errors_df = pd.concat([errors_df, pd.DataFrame(neg_errors)], ignore_index=True)

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

        for view_code, df_view in view_tabs_dict.items():
            df_view.to_excel(writer, sheet_name=f"Missing_{view_code}"[:31], index=False)

        sap_df.to_excel(writer, sheet_name="Original Data", index=False)

        # Assignments sheet — enriched with owner/status from master.json
        try:
            _enriched = merge_assignments(errors_df)
            _asgn_df  = build_assignments_sheet(_enriched)
            _asgn_df.to_excel(writer, sheet_name="Assignments", index=False)
            save_open_assignments(_enriched)
        except Exception:
            pass

    # Apply full professional styling
    file_bytes.seek(0)
    file_bytes = _style_sap_report(
        file_bytes,
        errors_df=errors_df,
        views_missing_output_df=views_missing_output_df,
        contagem_final_df=contagem_final_df,
        coverage_df=coverage_df,
        view_tabs_dict=view_tabs_dict,
        sap_df=sap_df,
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
    mbew_file, ekpo_file, mard_file, plant_filter=None,
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
    mbew      = load_xlsx_fast(mbew_file,      sheet_name=0)
    ekpo      = load_xlsx_fast(ekpo_file,      sheet_name=0)
    mard      = load_xlsx_fast(mard_file,      sheet_name=0)

    if mara is None or marc is None or zplm_marc is None:
        raise ValueError("Missing required files for Data Cleanups (MARA, MARC, ZPLM MARC).")

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
    # VALIDATIONS
    # ===============================================================
    results = {}

    # FP Status
    df_fp, kpi_fp = run_fp_status_validation(
        mara_file=mara_file, marc_file=marc_file,
        zplm_marc_file=zplm_marc_file, zplm_mara_file=zplm_mara_file,
        plant_filter=plant_filter
    )
    results["FP Level Status"] = df_fp

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
):
    """
    Validates Advanced Planning fields (new rules):

    - Advanced Planning:
        Must be 'X' for:
          * PT01 + FERT
          * PT02 + FERT
          * PT02 + HALB

    - Product Heuristic:
        Must NOT be empty for FERT and HALB

    - PP Planning Procedure:
        Must be '4' for FERT and HALB

    - Proposed Strategy (Requirement Strategy):
        Must be '20' for FERT and HALB
    """

    if marc_df is None or marc_df.empty or mara_df is None or mara_df.empty:
        return

    # Helper to pick first existing column
    def _pick(df, candidates):
        return next((c for c in candidates if c in df.columns), None)

    col_adv  = _pick(marc_df, ['Advanced Planning'])
    col_heur = _pick(marc_df, ['Heuristic'])
    col_pp   = _pick(marc_df, ['PP Plng Procedure'])
    col_req  = _pick(marc_df, ['Requirement Strategy'])

    # If any required column is missing, skip silently
    if any(c is None for c in [col_adv, col_heur, col_pp, col_req]):
        return

    if not {'Material', 'Material Type'}.issubset(mara_df.columns):
        return

    # Build material → material type map
    mt_map = dict(zip(
        mara_df['Material'].astype(str).str.strip(),
        mara_df['Material Type'].astype(str).str.strip().str.upper()
    ))

    def _s(x):
        return '' if pd.isna(x) else str(x).strip()

    for idx, row in marc_df.iterrows():
        material = _s(row.get('Material'))
        plant    = _s(row.get('Plant'))
        real_mt  = mt_map.get(material, '').upper()

        # ✅ Apply rule-specific material type override logic
        mt = get_effective_material_type(
            material,
            real_mt,
            'validar_advanced_planning_fields',
        )

        if mt not in {'FERT', 'HALB'}:
            continue

        adv  = _s(row.get(col_adv)).upper()
        heur = _s(row.get(col_heur))
        pp   = normalize_code(row.get(col_pp), as_int=True)
        req  = normalize_code(row.get(col_req), as_int=True)

        # ── Advanced Planning ──────────────────────────────────────────
        if (
            (plant == 'PT01' and mt == 'FERT') or
            (plant == 'PT02' and mt in {'FERT', 'HALB'})
        ):
            if adv != 'X':
                resultados.append({
                    'Line': idx + 2,
                    'Field': 'Advanced Planning',
                    'Error': (
                        "Mandatory: must be 'X' for "
                        "PT01 FERT or PT02 FERT/HALB"
                    ),
                    'Material': material,
                    'Plant': plant,
                })

        # ── Product Heuristic ──────────────────────────────────────────
        if heur == '':
            resultados.append({
                'Line': idx + 2,
                'Field': 'Heuristic',
                'Error': "Mandatory: cannot be empty for FERT or HALB",
                'Material': material,
                'Plant': plant,
            })

        # ── PP Planning Procedure ──────────────────────────────────────
        if pp != '4':
            resultados.append({
                'Line': idx + 2,
                'Field': 'PP Plng Procedure',
                'Error': "Invalid value: should be '4'",
                'Material': material,
                'Plant': plant,
            })

        # ── Proposed Strategy (Requirement Strategy) ───────────────────
        if req != '20':
            resultados.append({
                'Line': idx + 2,
                'Field': 'Requirement Strategy',
                'Error': "Invalid value: should be '20'",
                'Material': material,
                'Plant': plant,
            })


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


def build_sap_log_payload(*, user, plant_filter, sap_file, mara_file, marc_file,
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
        "files": {"sap": _name(sap_file, "Item Browse Report.xlsx"),
                  "mara": _name(mara_file, "MARA.xlsx"),
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
                 "DATA_CLEANUPS": "dataclean"}.get(run_type, "misc")

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
                  "mrp": "mrp", "dataclean": "dataclean"}
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


def get_owner(field: str, plant: str = "", master: dict | None = None) -> str:
    """
    Returns the owner for a (field, plant) combination.
    Plant-level override takes priority over field-level default.
    Returns "" if not assigned.
    """
    if master is None:
        master = _load_master()
    plant = (plant or "").strip().upper()
    # Plant override
    override = master.get("plant_overrides", {}).get(plant, {}).get(field, {})
    if isinstance(override, dict) and override.get("owner"):
        return override["owner"]
    # Field default
    default = master.get("field_assignments", {}).get(field, {})
    if isinstance(default, dict):
        return default.get("owner", "")
    return ""


def set_owner(field: str, owner: str, *,
              plant: str = "",
              team: str = "",
              master: dict | None = None) -> dict:
    """
    Set (or update) the owner for a field.
    If plant is given, sets a plant-level override; otherwise sets the field default.
    Returns the updated master dict.
    """
    if master is None:
        master = _load_master()
    plant = (plant or "").strip().upper()
    owner = (owner or "").strip()
    now   = datetime.now().strftime("%Y-%m-%d")

    if plant:
        master.setdefault("plant_overrides", {}).setdefault(plant, {})[field] = {
            "owner": owner, "team": team, "since": now
        }
    else:
        master.setdefault("field_assignments", {})[field] = {
            "owner": owner, "team": team, "since": now
        }
    _save_master(master)
    return master


def remove_owner(field: str, *, plant: str = "") -> None:
    """Remove an assignment (field default or plant override)."""
    master = _load_master()
    plant  = (plant or "").strip().upper()
    if plant:
        master.get("plant_overrides", {}).get(plant, {}).pop(field, None)
    else:
        master.get("field_assignments", {}).pop(field, None)
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


def merge_assignments(errors_df: pd.DataFrame,
                      run_date: str | None = None) -> pd.DataFrame:
    """
    Enriches errors_df with assignment data from master.json.

    Adds columns:
      Owner         — person responsible for this field
      Team          — team of the owner
      Status        — "New" | "Open" | "Overdue" (no due date = "New")
      Days_Open     — days since first seen (from previous_assignments if available)
      Due_Date      — ISO date string or ""
      Notes         — free-text notes (blank by default)

    The function also reads previous assignment rows from the
    .history/assignments/open_items.json file (written by
    save_open_assignments) to carry forward Days_Open and Due_Date.
    """
    if errors_df is None or errors_df.empty:
        return errors_df

    master    = _load_master()
    today_str = run_date or datetime.now().strftime("%Y-%m-%d")
    today     = datetime.strptime(today_str, "%Y-%m-%d")

    # Load previously open items for carry-forward
    open_path = os.path.join(".history", "assignments", "open_items.json")
    open_map: dict[tuple, dict] = {}
    if os.path.exists(open_path):
        try:
            with open(open_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            # key: "Material||Plant||Field" → dict with due_date, notes, first_seen
            for k, v in raw.items():
                parts = k.split("||")
                if len(parts) == 3:
                    open_map[tuple(parts)] = v
        except Exception:
            pass

    df = errors_df.copy()

    owners, teams, statuses, days_open_list, due_dates, notes_list = [], [], [], [], [], []

    for _, row in df.iterrows():
        field    = str(row.get("Field",    "")).strip()
        plant    = str(row.get("Plant",    "")).strip().upper()
        material = str(row.get("Material", "")).strip()

        owner = get_owner(field, plant, master=master)
        team  = ""
        if owner:
            # Get team from master
            override = master.get("plant_overrides", {}).get(plant, {}).get(field, {})
            default  = master.get("field_assignments", {}).get(field, {})
            team = (override if isinstance(override, dict) and override.get("owner") == owner
                    else default).get("team", "")

        # Carry-forward from previous open items
        item_key = (material, plant, field)
        prev     = open_map.get(item_key, {})
        due_str  = prev.get("due_date", "")
        notes    = prev.get("notes", "")
        first_seen_str = prev.get("first_seen", today_str)

        try:
            first_seen = datetime.strptime(first_seen_str, "%Y-%m-%d")
            d_open = (today - first_seen).days
        except Exception:
            d_open = 0

        # Status
        if not owner:
            status = "Unassigned"
        elif due_str:
            try:
                due_dt = datetime.strptime(due_str, "%Y-%m-%d")
                status = "Overdue" if today > due_dt else "Open"
            except Exception:
                status = "Open"
        else:
            status = "New" if d_open == 0 else "Open"

        owners.append(owner)
        teams.append(team)
        statuses.append(status)
        days_open_list.append(d_open)
        due_dates.append(due_str)
        notes_list.append(notes)

    df["Owner"]     = owners
    df["Team"]      = teams
    df["Status"]    = statuses
    df["Days_Open"] = days_open_list
    df["Due_Date"]  = due_dates
    df["Notes"]     = notes_list

    return df


def save_open_assignments(enriched_df: pd.DataFrame) -> None:
    """
    Persists the current open assignment items to
    .history/assignments/open_items.json for carry-forward in the next run.

    Only saves rows that have an Owner (assigned items).
    Preserves first_seen date from existing entries.
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
            "due_date":   str(row.get("Due_Date",  "")).strip(),
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
             Status | Days_Open | Due_Date | Notes
    Sorted: Overdue first, then Open, then New, then Unassigned.
    """
    cols = ["Material", "Plant", "Field", "Error",
            "Owner", "Team", "Status", "Days_Open", "Due_Date", "Notes"]

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
    Overdue = red, Open = orange, New = blue, Unassigned = grey.
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    STATUS_FILLS = {
        "Overdue":    PatternFill(start_color="FFD7D7", fill_type="solid"),
        "Open":       PatternFill(start_color="FFE5B4", fill_type="solid"),
        "New":        PatternFill(start_color="DBEAFE", fill_type="solid"),
        "Unassigned": PatternFill(start_color="F3F4F6", fill_type="solid"),
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
# SSA VALIDATION ENGINE
# Schedule Agreements (SSA) validation module
# Tables: VBAK, VBAP, VBKD, VBPA, KNA1, MARC
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
# Rule 1 — Mandatory fields
# -----------------------------------------------------------------------------
def ssa_validar_mandatory_fields(vbak_df, vbap_df):
    """
    Fields that must not be empty:
      VBAK: Customer Group 1, Customer Group 2, Customer Group 3
      VBAP: Plant, Route, IFS ID, Supplier_code,
            Acct Assmt Grp Mat. (must be 'Z1')
    """
    errors = []

    # ── VBAK ──────────────────────────────────────────────────────────────────
    if vbak_df is not None and not vbak_df.empty:
        vbak = _ssa_load(vbak_df)
        doc_col = _ssa_pick(vbak, ["Sales Document", "VBELN"])

        for field in ["Customer Group 1", "Customer Group 2", "Customer Group 3"]:
            col = _ssa_pick(vbak, [field])
            if col is None:
                continue
            for idx, row in vbak.iterrows():
                if _ssa_norm(row[col]) == '':
                    errors.append({
                        "Table":           "VBAK",
                        "Sales Document":  _ssa_norm(row[doc_col]) if doc_col else '',
                        "Item":            '',
                        "Material":        '',
                        "Plant":           '',
                        "Field":           field,
                        "Value":           '',
                        "Error":           f"Mandatory field blank: '{field}'",
                    })

    # ── VBAP ──────────────────────────────────────────────────────────────────
    if vbap_df is not None and not vbap_df.empty:
        vbap = _ssa_load(vbap_df)
        doc_col  = _ssa_pick(vbap, ["Sales Document", "VBELN"])
        item_col = _ssa_pick(vbap, ["Sales Document Item", "POSNR"])
        mat_col  = _ssa_pick(vbap, ["Material", "MATNR"])
        plt_col  = _ssa_pick(vbap, ["Plant", "WERKS"])
        rej_col  = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

        # Skip rejected items for mandatory checks
        def is_rejected(row):
            if rej_col is None:
                return False
            return _ssa_norm(row[rej_col]) != ''

        simple_mandatory = ["Route", "IFS ID", "Supplier_code"]
        for field in simple_mandatory:
            col = _ssa_pick(vbap, [field])
            if col is None:
                continue
            for idx, row in vbap.iterrows():
                if is_rejected(row):
                    continue
                if _ssa_norm(row[col]) == '':
                    errors.append({
                        "Table":          "VBAP",
                        "Sales Document": _ssa_norm(row[doc_col])  if doc_col  else '',
                        "Item":           _ssa_norm(row[item_col]) if item_col else '',
                        "Material":       _ssa_norm(row[mat_col])  if mat_col  else '',
                        "Plant":          _ssa_norm(row[plt_col])  if plt_col  else '',
                        "Field":          field,
                        "Value":          '',
                        "Error":          f"Mandatory field blank: '{field}'",
                    })

        # Plant — mandatory
        if plt_col:
            for idx, row in vbap.iterrows():
                if is_rejected(row):
                    continue
                if _ssa_norm(row[plt_col]) == '':
                    errors.append({
                        "Table": "VBAP", "Sales Document": _ssa_norm(row[doc_col]) if doc_col else '',
                        "Item": _ssa_norm(row[item_col]) if item_col else '',
                        "Material": _ssa_norm(row[mat_col]) if mat_col else '',
                        "Plant": '', "Field": "Plant",
                        "Value": '', "Error": "Mandatory field blank: 'Plant'",
                    })

        # Acct Assmt Grp Mat. — mandatory AND must be 'Z1'
        aag_col = _ssa_pick(vbap, ["Acct Assmt Grp Mat.", "KTGRM"])
        if aag_col:
            for idx, row in vbap.iterrows():
                if is_rejected(row):
                    continue
                val = _ssa_norm(row[aag_col])
                if val == '':
                    errors.append({
                        "Table": "VBAP", "Sales Document": _ssa_norm(row[doc_col]) if doc_col else '',
                        "Item": _ssa_norm(row[item_col]) if item_col else '',
                        "Material": _ssa_norm(row[mat_col]) if mat_col else '',
                        "Plant": _ssa_norm(row[plt_col]) if plt_col else '',
                        "Field": "Acct Assmt Grp Mat.",
                        "Value": '', "Error": "Mandatory field blank: 'Acct Assmt Grp Mat.'",
                    })
                elif val.upper() != 'Z1':
                    errors.append({
                        "Table": "VBAP", "Sales Document": _ssa_norm(row[doc_col]) if doc_col else '',
                        "Item": _ssa_norm(row[item_col]) if item_col else '',
                        "Material": _ssa_norm(row[mat_col]) if mat_col else '',
                        "Plant": _ssa_norm(row[plt_col]) if plt_col else '',
                        "Field": "Acct Assmt Grp Mat.",
                        "Value": val,
                        "Error": f"Invalid value '{val}' — must always be 'Z1'",
                    })

    return pd.DataFrame(errors)


# -----------------------------------------------------------------------------
# Rule 2 — Incoterms DDP + Country GB
# -----------------------------------------------------------------------------
def ssa_validar_incoterms_ddp_gb(vbak_df, vbkd_df, kna1_df, vbpa_df=None):
    """
    If Incoterms (VBKD) = 'DDP' AND Customer Country = 'GB':
      - VBAK Alt.Tax Classific. must be '1'
      - VBAK Tax Departure C/R  must be 'GB'

    Country resolved via:
      Primary:  VBAK~Sold-to Party → KNA1~Customer → KNA1~Country/Region Key
      Optional: VBPA (Partner Function = 'SP') → Country/Region Key
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant",
                  "Field", "Value", "Error"]

    if any(df is None or df.empty for df in [vbak_df, vbkd_df, kna1_df]):
        return pd.DataFrame(columns=empty_cols)

    vbak = _ssa_load(vbak_df)
    vbkd = _ssa_load(vbkd_df)
    kna1 = _ssa_load(kna1_df)

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

    # Build VBKD map: sales doc -> incoterms
    vbkd_inc_map = {}
    for _, r in vbkd.iterrows():
        doc = _ssa_norm(r[vbkd_doc])
        inc = _ssa_norm(r[vbkd_inc]).upper()
        vbkd_inc_map[doc] = inc

    # Evaluate each VBAK row
    for _, row in vbak.iterrows():
        doc      = _ssa_norm(row[vbak_doc])
        stp      = _ssa_norm(row[vbak_stp])
        incoterm = vbkd_inc_map.get(doc, '')
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
    Incompletion Log checks:
    A) VBAK Overall Header (FA) must not be 'A'
    B) VBAP Item (JC) must not be 'A'
       Exception: if Reason for Rejection is filled AND Item = 'A'
                  → report as 'Rejected item with incompletion log' (not blocked)
    C) Skip if all items in the SSA are rejected
    """
    errors = []

    # ── VBAK header check ─────────────────────────────────────────────────────
    if vbak_df is not None and not vbak_df.empty:
        vbak    = _ssa_load(vbak_df)
        doc_col = _ssa_pick(vbak, ["Sales Document", "VBELN"])
        oh_col  = _ssa_pick(vbak, ["Overall Header", "UVALL"])

        if oh_col:
            for _, row in vbak.iterrows():
                val = _ssa_norm(row[oh_col]).upper()
                if val == 'A':
                    errors.append({
                        "Table":          "VBAK",
                        "Sales Document": _ssa_norm(row[doc_col]) if doc_col else '',
                        "Item": '', "Material": '', "Plant": '',
                        "Field":  "Overall Header",
                        "Value":  val,
                        "Error":  "Incompletion Log: Overall Header = 'A' (document is incomplete)",
                    })

    # ── VBAP item check ───────────────────────────────────────────────────────
    if vbap_df is not None and not vbap_df.empty:
        vbap     = _ssa_load(vbap_df)
        doc_col  = _ssa_pick(vbap, ["Sales Document", "VBELN"])
        item_col = _ssa_pick(vbap, ["Sales Document Item", "POSNR"])
        mat_col  = _ssa_pick(vbap, ["Material", "MATNR"])
        plt_col  = _ssa_pick(vbap, ["Plant", "WERKS"])
        rej_col  = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])
        itm_col  = _ssa_pick(vbap, ["Item", "UVPOS"])   # JC = Item status

        if itm_col:
            # Group by Sales Document to check if ALL items are rejected
            if doc_col and rej_col:
                all_rejected_docs = set()
                for doc, grp in vbap.groupby(vbap[doc_col].astype(str).str.strip()):
                    if grp[rej_col].astype(str).str.strip().replace('', pd.NA).notna().all():
                        all_rejected_docs.add(doc)
            else:
                all_rejected_docs = set()

            for _, row in vbap.iterrows():
                doc = _ssa_norm(row[doc_col]) if doc_col else ''
                if doc in all_rejected_docs:
                    continue  # Skip entirely-rejected SSAs

                val     = _ssa_norm(row[itm_col]).upper()
                rej_val = _ssa_norm(row[rej_col]) if rej_col else ''

                if val == 'A':
                    if rej_val != '':
                        # Rejected item with incompletion — informational
                        errors.append({
                            "Table": "VBAP", "Sales Document": doc,
                            "Item":     _ssa_norm(row[item_col]) if item_col else '',
                            "Material": _ssa_norm(row[mat_col])  if mat_col  else '',
                            "Plant":    _ssa_norm(row[plt_col])  if plt_col  else '',
                            "Field": "Item",
                            "Value": val,
                            "Error": (f"Rejected item with Incompletion Log "
                                      f"(Reason for Rejection: '{rej_val}', Item status: 'A')"),
                        })
                    else:
                        errors.append({
                            "Table": "VBAP", "Sales Document": doc,
                            "Item":     _ssa_norm(row[item_col]) if item_col else '',
                            "Material": _ssa_norm(row[mat_col])  if mat_col  else '',
                            "Plant":    _ssa_norm(row[plt_col])  if plt_col  else '',
                            "Field": "Item",
                            "Value": val,
                            "Error": "Incompletion Log: Item status = 'A' (item is incomplete)",
                        })

    return pd.DataFrame(errors)


# -----------------------------------------------------------------------------
# Rule 4 — Rejected items with Material Status 70/75
# -----------------------------------------------------------------------------
def ssa_validar_rejected_material_status(vbap_df, marc_df):
    """
    If Reason for Rejection is filled AND Material+Plant has
    Material Status (MARC) = 70 or 75 → report as error.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material",
                  "Plant", "Field", "Value", "Error"]

    if vbap_df is None or vbap_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbap = _ssa_load(vbap_df)
    marc = _ssa_load(marc_df)

    doc_col  = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "POSNR"])
    mat_col  = _ssa_pick(vbap, ["Material", "MATNR"])
    plt_col  = _ssa_pick(vbap, ["Plant", "WERKS"])
    rej_col  = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

    marc_mat  = _ssa_pick(marc, ["Material", "MATNR"])
    marc_plt  = _ssa_pick(marc, ["Plant", "WERKS"])
    marc_psms = _ssa_pick(marc, ["Plant-Specific Material Status",
                                  "Plant Specific Material Status", "MTSTB"])

    if not all([doc_col, mat_col, plt_col, rej_col,
                marc_mat, marc_plt, marc_psms]):
        return pd.DataFrame(columns=empty_cols)

    # Build MARC status map: (material, plant) -> status
    def _norm_status(x):
        if pd.isna(x) or str(x).strip() == '':
            return ''
        try:
            return str(int(float(str(x).strip())))
        except Exception:
            return str(x).strip()

    marc_map = {}
    for _, r in marc.iterrows():
        key = (_ssa_norm(r[marc_mat]), _ssa_norm(r[marc_plt]).upper())
        marc_map[key] = _norm_status(r[marc_psms])

    for _, row in vbap.iterrows():
        rej = _ssa_norm(row[rej_col])
        if rej == '':
            continue  # Only process rejected items

        mat   = _ssa_norm(row[mat_col])
        plant = _ssa_norm(row[plt_col]).upper()
        status = marc_map.get((mat, plant), '')

        if status in {'70', '75'}:
            errors.append({
                "Table":          "VBAP",
                "Sales Document": _ssa_norm(row[doc_col])  if doc_col  else '',
                "Item":           _ssa_norm(row[item_col]) if item_col else '',
                "Material":       mat,
                "Plant":          plant,
                "Field":          "Reason for Rejection",
                "Value":          rej,
                "Error": (f"Rejected item with Material Status {status} "
                          f"(Material={mat}, Plant={plant}) — "
                          f"status {status} indicates obsolete/discontinued material"),
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


# -----------------------------------------------------------------------------
# Rule 5 — Country of Origin check
# -----------------------------------------------------------------------------
def ssa_validar_country_of_origin(vbap_df, marc_df):
    """
    For every SSA line item, join Material+Plant with MARC to get
    MARC~HERKL (Country of Origin).
    Error if HERKL is empty for non-rejected items.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material",
                  "Plant", "Field", "Value", "Error"]

    if vbap_df is None or vbap_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbap = _ssa_load(vbap_df)
    marc = _ssa_load(marc_df)

    doc_col  = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "POSNR"])
    mat_col  = _ssa_pick(vbap, ["Material", "MATNR"])
    plt_col  = _ssa_pick(vbap, ["Plant", "WERKS"])
    rej_col  = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

    marc_mat  = _ssa_pick(marc, ["Material", "MATNR"])
    marc_plt  = _ssa_pick(marc, ["Plant", "WERKS"])
    marc_coo  = _ssa_pick(marc, ["Country of Origin", "HERKL"])

    if not all([doc_col, mat_col, plt_col, marc_mat, marc_plt]):
        return pd.DataFrame(columns=empty_cols)

    if marc_coo is None:
        # HERKL not in MARC export — skip silently
        return pd.DataFrame(columns=empty_cols)

    # Build COO map
    coo_map = {}
    for _, r in marc.iterrows():
        key = (_ssa_norm(r[marc_mat]), _ssa_norm(r[marc_plt]).upper())
        coo_map[key] = _ssa_norm(r[marc_coo])

    for _, row in vbap.iterrows():
        rej = _ssa_norm(row[rej_col]) if rej_col else ''
        if rej != '':
            continue  # Skip rejected items

        mat   = _ssa_norm(row[mat_col])
        plant = _ssa_norm(row[plt_col]).upper()
        coo   = coo_map.get((mat, plant), None)

        if coo is None:
            errors.append({
                "Table":          "VBAP",
                "Sales Document": _ssa_norm(row[doc_col])  if doc_col  else '',
                "Item":           _ssa_norm(row[item_col]) if item_col else '',
                "Material":       mat, "Plant": plant,
                "Field":          "Country of Origin (HERKL)",
                "Value":          '',
                "Error": (f"Material+Plant not found in MARC "
                          f"(Material={mat}, Plant={plant})"),
            })
        elif coo == '':
            errors.append({
                "Table":          "VBAP",
                "Sales Document": _ssa_norm(row[doc_col])  if doc_col  else '',
                "Item":           _ssa_norm(row[item_col]) if item_col else '',
                "Material":       mat, "Plant": plant,
                "Field":          "Country of Origin (HERKL)",
                "Value":          '',
                "Error": (f"Country of Origin (MARC~HERKL) is empty "
                          f"for Material={mat}, Plant={plant}"),
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


# -----------------------------------------------------------------------------
# Orchestrator — run all SSA validations
# -----------------------------------------------------------------------------
def run_ssa_validation(
    vbak_file, vbap_file, vbkd_file, kna1_file,
    marc_file=None, vbpa_file=None,
    plant_filter=None,
):
    """
    Runs all SSA (Schedule Agreement) validations.
    Returns: (excel_bytes, results_dict, summary_df)

    Excel report structure (same style as SAP Material Validation):
      - Adjusted Summary   : error count per Field, sorted descending
      - Field With Error   : all actionable errors
      - All Errors         : same as above (full detail)
      - Original Data VBAK : raw VBAK input
      - Original Data VBAP : raw VBAP input
    """
    from openpyxl.utils import get_column_letter as _gcl

    # ── Load ──────────────────────────────────────────────────────────────────
    vbak = load_xlsx_fast(vbak_file, sheet_name=0) if vbak_file else None
    vbap = load_xlsx_fast(vbap_file, sheet_name=0) if vbap_file else None
    vbkd = load_xlsx_fast(vbkd_file, sheet_name=0) if vbkd_file else None
    kna1 = load_xlsx_fast(kna1_file, sheet_name=0) if kna1_file else None
    marc = load_xlsx_fast(marc_file, sheet_name=0) if marc_file else None
    vbpa = load_xlsx_fast(vbpa_file, sheet_name=0) if vbpa_file else None

    if any(df is None or df.empty for df in [vbak, vbap, vbkd, kna1]):
        raise ValueError("Missing required files: VBAK, VBAP, VBKD and KNA1 are mandatory.")

    # ── Plant filter ──────────────────────────────────────────────────────────
    if plant_filter and vbap is not None:
        plt_col = _ssa_pick(_ssa_load(vbap), ["Plant", "WERKS"])
        if plt_col:
            vbap = vbap[vbap[plt_col].astype(str).str.strip().str.upper()
                        == str(plant_filter).upper()].copy()

    # ── Run all validations ───────────────────────────────────────────────────
    all_frames = [
        ssa_validar_mandatory_fields(vbak, vbap),
        ssa_validar_incoterms_ddp_gb(vbak, vbkd, kna1, vbpa),
        ssa_validar_incompletion_log(vbak, vbap),
        ssa_validar_rejected_material_status(vbap, marc),
        ssa_validar_country_of_origin(vbap, marc),
    ]

    # Combine all errors into one DataFrame
    all_errors_df = pd.concat(
        [f for f in all_frames if f is not None and not f.empty],
        ignore_index=True
    ) if any(f is not None and not f.empty for f in all_frames) else pd.DataFrame(
        columns=["Table","Sales Document","Item","Material","Plant","Field","Value","Error"]
    )

    # results_dict — one entry per validation rule (for the app tabs)
    rule_names = [
        "Mandatory Fields",
        "Incoterms DDP + GB",
        "Incompletion Log",
        "Rejected + Status 70-75",
        "Country of Origin",
    ]
    results = {
        name: df for name, df in zip(rule_names, all_frames)
        if df is not None and not df.empty
    }

    # ── Adjusted Summary — errors per Field, sorted descending ───────────────
    if not all_errors_df.empty and "Field" in all_errors_df.columns:
        summary_df = (
            all_errors_df["Field"]
            .value_counts()
            .rename_axis("Field")
            .reset_index(name="Nº of Errors")
        )
        # Add Rule column
        field_to_rule = {}
        for rule_name, df in results.items():
            if "Field" in df.columns:
                for field in df["Field"].unique():
                    field_to_rule[field] = rule_name
        summary_df["Validation Rule"] = summary_df["Field"].map(field_to_rule).fillna("—")
        summary_df = summary_df[["Field", "Validation Rule", "Nº of Errors"]]
    else:
        summary_df = pd.DataFrame(columns=["Field", "Validation Rule", "Nº of Errors"])

    # ── Write Excel ───────────────────────────────────────────────────────────
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # 1) Adjusted Summary
        summary_df.to_excel(writer, sheet_name="Adjusted Summary", index=False)
        # 2) Field With Error (actionable errors only — same as All Errors for SSA)
        all_errors_df.to_excel(writer, sheet_name="Field With Error", index=False)
        # 3) All Errors (full)
        all_errors_df.to_excel(writer, sheet_name="All Errors", index=False)
        # 4) Original Data
        vbak.to_excel(writer, sheet_name="Original Data VBAK", index=False)
        vbap.to_excel(writer, sheet_name="Original Data VBAP", index=False)

    # ── Style ─────────────────────────────────────────────────────────────────
    out.seek(0)
    wb = load_workbook(out)

    thin   = Side(style="thin", color="B8CCE4")
    bdr    = Border(left=thin, right=thin, top=thin, bottom=thin)
    zebra  = PatternFill(start_color="F5F9FF", fill_type="solid")

    TAB_COLORS = {
        "Adjusted Summary": "2E75B6",
        "Field With Error": "FF7200",
        "All Errors":       "C00000",
        "Original Data VBAK": "70AD47",
        "Original Data VBAP": "70AD47",
    }
    HEADER_COLORS = {
        "Adjusted Summary": "BDD7EE",
        "Field With Error": "FFE0CC",
        "All Errors":       "FFD7D7",
        "Original Data VBAK": "E2EFDA",
        "Original Data VBAP": "E2EFDA",
    }

    for sn in wb.sheetnames:
        ws = wb[sn]
        ws.sheet_properties.tabColor = TAB_COLORS.get(sn, "4F81BD")
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 26

        hfill = PatternFill(
            start_color=HEADER_COLORS.get(sn, "BDD7EE"), fill_type="solid"
        )
        hfont = Font(bold=True, color="1F4E79")
        halign = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Header row
        for cell in ws[1]:
            cell.fill    = hfill
            cell.font    = hfont
            cell.alignment = halign
            cell.border  = bdr

        # Data rows — zebra + border + wrap Error column
        err_col_idx = next(
            (i for i, cell in enumerate(ws[1], start=1)
             if str(cell.value) == "Error"), None
        )
        for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            for cell in row:
                cell.border = bdr
                cell.alignment = Alignment(vertical="center", wrap_text=(cell.column == err_col_idx))
                if r_idx % 2 == 0:
                    if cell.fill.fgColor.rgb in ("00000000", "FFFFFFFF", "00FFFFFF"):
                        cell.fill = zebra

        # Auto-width
        for col_idx, col in enumerate(ws.columns, start=1):
            max_len = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[_gcl(col_idx)].width = min(max_len + 3, 55)

        # Error column wider
        if err_col_idx:
            ws.column_dimensions[_gcl(err_col_idx)].width = 60

    # Reorder sheets: Summary first, Original Data last
    desired_order = [
        "Adjusted Summary", "Field With Error", "All Errors",
        "Original Data VBAK", "Original Data VBAP",
    ]
    existing = [s for s in desired_order if s in wb.sheetnames]
    wb._sheets.sort(key=lambda ws: existing.index(ws.title) if ws.title in existing else 99)

    out2 = io.BytesIO()
    wb.save(out2)
    out2.seek(0)
    return out2, results, summary_df