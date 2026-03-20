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
#   B) Cross-file validations (SAPUI5 + MVKE, QMAT, MARC, MARA, ZPLM)
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


def validar_per_lot_sz_plng_cal(marc_df, resultados, mara_df):
    """
    Per.Lot Sz.Plng Cal. validation — sourced from MARC (not SAPUI5 Export).

    Rules (scoped to FERT, HALB, ROH only — other Material Types are ignored):
      - FERT or HALB -> field is mandatory (must not be empty).
      - ROH           -> field must be empty.

    Parameters
    ----------
    marc_df    : MARC DataFrame — contains the Per.Lot Sz.Plng Cal. column
    resultados : list to append errors to
    mara_df    : MARA DataFrame — used to look up Material Type per material
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

    mt_map = dict(zip(
        mara_df['Material'].astype(str).str.strip(),
        mara_df['Material Type'].astype(str).str.strip().str.upper()
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

        if mt in mandatory_types:
            if val == '':
                resultados.append({
                    'Line':     idx + 2,
                    'Field':    field,
                    'Error':    f"Mandatory field blank for Material Type '{mt}' (FERT and HALB require a value)",
                    'Material': material,
                    'Plant':    plant,
                })
        else:
            if val != '':
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
        mt  = str(row['Material Type']).strip()
        mg  = str(row['Material Group']).strip()

        if pd.isna(row['Material Group']) or mg == '':
            erros.append('Mandatory field blank')
        else:
            if mt == 'HALB' and mg in ['L004', 'APPOINT']:
                erros.append(f"For HALB, cannot be '{mg}'")
            elif mt == 'FERT' and mg in ['L003', 'APPOINT']:
                erros.append(f"For FERT, cannot be '{mg}'")
            elif mt == 'ROH' and mg in ['L003', 'L004']:
                erros.append(f"For ROH, cannot be '{mg}'")

        if erros:
            resultados.append({
                'Line': idx + 2, 'Field': 'Material Group',
                'Error': ' \n'.join(erros),
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

        if psms in (70, 75):
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

    MVKE Product Hierarchy (PRODH) consistency:
      If MARA 'Product Hierarchy' is filled for a material, every existing
      MVKE entry for that material must also have PRODH filled.
      Each offending SalesOrg/DC combination is reported individually.
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

    # -- MVKE Product Hierarchy (PRODH) consistency ---------------------------
    if mvke_df is None or mvke_df.empty:
        return

    # Build material -> MARA Product Hierarchy map
    prdha_map = {}
    if mara_df is not None and not mara_df.empty:
        m = mara_df.copy()
        m.columns = [c.strip() for c in m.columns]
        _col_mat = next((c for c in ['Material', 'MATNR'] if c in m.columns), None)
        _col_ph  = next((c for c in ['Product Hierarchy', 'PRDHA', 'Prod. Hierarchy']
                         if c in m.columns), None)
        if _col_mat and _col_ph:
            for _, r in m.iterrows():
                val = '' if pd.isna(r[_col_ph]) else str(r[_col_ph]).strip()
                if val:
                    prdha_map[str(r[_col_mat]).strip()] = val
    else:
        # Fallback: read from SAPUI5 Export
        for _, r in df.iterrows():
            val = '' if pd.isna(r.get('Product Hierarchy')) else str(r['Product Hierarchy']).strip()
            if val:
                prdha_map[str(r.get('Material', '')).strip()] = val

    if not prdha_map:
        return

    mvke_norm, _ = _prepare_mvke(mvke_df, set(prdha_map.keys()))
    if mvke_norm is None or mvke_norm.empty:
        return

    for _, r in mvke_norm.iterrows():
        mat = r['_mat']
        if mat not in prdha_map:
            continue
        prodh = r.get('_prodh', '')
        if prodh == '' or str(prodh).upper() in ('NAN', 'NONE'):
            combo = f"SalesOrg='{r['_sorg']}' / DC='{r['_dchan']}'"
            resultados.append({
                'Line': '', 'Field': 'Product Hierarchy',
                'Error': (f"MARA Product Hierarchy is '{prdha_map[mat]}' "
                          f"but MVKE Product Hierarchy (PRODH) is blank for {combo}"),
                'Material': mat, 'Plant': '',
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
    """Purchasing Group is mandatory only for Plant Type = PP."""
    if 'Plant Type' not in df.columns or 'Purchasing Group' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'Purchasing Group' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        if str(row['Plant Type']).strip() == 'PP':
            val = str(row['Purchasing Group']).strip()
            if pd.isna(row['Purchasing Group']) or val == '':
                erros.append("Mandatory field blank For Plant Type 'PP'")
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
    - 'P105' -> only PT01
    - 'P205' / 'P220' -> only PT02
    """
    if 'Plant' not in df.columns or 'Issue Storage Location' not in df.columns:
        print("⚠️ Columns 'Plant' or 'Issue Storage Location' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        plant = str(row['Plant']).strip()
        val   = str(row['Issue Storage Location']).strip()
        if pd.isna(row['Issue Storage Location']) or val == '':
            erros.append('Mandatory field blank')
        else:
            if val == 'P105' and plant != 'PT01':
                erros.append(f"Value 'P105' only allowed for Plant = PT01 (Current: '{plant}')")
            elif val in ['P205', 'P220'] and plant != 'PT02':
                erros.append(f"Value '{val}' only allowed for Plant = PT02 (Current: '{plant}')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Issue Storage Location',
                               'Error': ' | '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant': plant})


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
        mt  = str(row['Material Type']).strip()
        val = str(row['Valuation Class']).strip()
        if pd.isna(row['Valuation Class']) or val == '':
            erros.append('Mandatory field blank')
        else:
            if val == '3000' and mt != 'ROH':
                erros.append(f"Value '3000' only allowed for Material Type = ROH (Current: '{mt}')")
            elif val == '7900' and mt != 'HALB':
                erros.append(f"Value '7900' only allowed for Material Type = HALB (Current: '{mt}')")
            elif val == '7920' and mt != 'FERT':
                erros.append(f"Value '7920' only allowed for Material Type = FERT (Current: '{mt}')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Valuation Class',
                               'Error': ' | '.join(erros),
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


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
    """MRP Controller is mandatory."""
    if 'MRP Controller' not in df.columns:
        return
    for idx, row in df.iterrows():
        if pd.isna(row['MRP Controller']) or str(row['MRP Controller']).strip() == '':
            resultados.append({'Line': idx + 2, 'Field': 'MRP Controller',
                               'Error': 'Mandatory field blank',
                               'Material': str(row.get('Material', '')).strip(),
                               'Plant':    str(row.get('Plant', '')).strip()})


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
# B) CROSS-FILE VALIDATIONS (SAPUI5 + MVKE / QMAT / MARC / MARA)
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
        mt         = str(row['Material Type']).strip()
        enterproj  = '' if pd.isna(row['Enterproj Number']) else str(row['Enterproj Number']).strip()
        material   = str(row.get('Material', '')).strip()

        if mt == 'FERT':
            if enterproj == '':
                erros.append("For Material Type 'FERT', the field is required and cannot be empty.")
        elif mt in ['ROH', 'HALB']:
            if enterproj != '':
                display = str(int(float(enterproj))) if enterproj.replace('.', '', 1).isdigit() else enterproj
                erros.append(f"For Material Type '{mt}', the field must be empty (current: '{display}')")

        if erros and not any(r['Material'] == material and r['Field'] == 'Enterproj Number' for r in resultados):
            resultados.append({'Line': idx + 2, 'Field': 'Enterproj Number',
                               'Error': ' | '.join(erros), 'Material': material, 'Plant': ''})


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


def validar_inspection_type_qmat(sap_df, marc_df, qmat_df, resultados):
    """
    QMAT — Inspection Type validation:
    - Rule 2 (ROH/HALB): Inspection Type '08' must NOT have Active flag = 'X'.
    - Rule 3 (FERT/HALB/ROH): Inspection Type '1708' must NOT have Active flag = 'X'.
    """
    if qmat_df is None or qmat_df.empty:
        print("⚠️ QMAT.xlsx missing, QMAT validation skipped.")
        return

    qmat = qmat_df.copy()
    qmat.columns = [c.strip().upper() for c in qmat.columns]

    col_mat    = next((c for c in ["MATERIAL", "MATNR"] if c in qmat.columns), None)
    col_plant  = next((c for c in ["PLANT", "WERKS"] if c in qmat.columns), None)
    col_insp   = next((c for c in ["INSPECTION TYPE", "QMTXT", "ART"] if c in qmat.columns), None)
    col_active = next((c for c in ["INSPTYPE-MATL ACTIVE", "ACTIVE", "STATUS"] if c in qmat.columns), None)

    if not all([col_mat, col_plant, col_insp, col_active]):
        print("⚠️ QMAT has insufficient columns. Skipped.")
        return

    qmat["MAT"]  = qmat[col_mat].astype(str).str.strip()
    qmat["PLT"]  = qmat[col_plant].astype(str).str.strip()
    qmat["INSP"] = qmat[col_insp].astype(str).str.strip().str.upper()
    qmat["ACT"]  = qmat[col_active].astype(str).str.strip().str.upper()

    mt_map = dict(zip(
        sap_df['Material'].astype(str).str.strip(),
        sap_df['Material Type'].astype(str).str.strip().str.upper()
    ))

    marc = marc_df.copy()
    marc.columns = [c.strip() for c in marc.columns]

    for idx, row in marc.iterrows():
        material = str(row.get("Material", "")).strip()
        plant    = str(row.get("Plant", "")).strip()
        mt       = mt_map.get(material, "")

        for insp_type, valid_types in [("08", ["HALB", "ROH"]), ("1708", ["FERT", "HALB", "ROH"])]:
            if mt in valid_types:
                subset = qmat[(qmat["MAT"] == material) & (qmat["PLT"] == plant) & (qmat["INSP"] == insp_type)]
                for _, r in subset.iterrows():
                    if r["ACT"] == "X":
                        resultados.append({
                            'Line': idx + 2, 'Field': f'Inspection Type (QMAT {insp_type})',
                            'Error': f"Inspection Type '{insp_type}' has Active flag = 'X' in QMAT",
                            'Material': material, 'Plant': plant
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
    """FERT/HALB/ROH (Material, Plant) in MARC that are missing in ZPLM MARC."""
    if marc_df is None or marc_df.empty or zplm_marc_df is None or zplm_marc_df.empty:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    marc = marc_df.copy(); marc.columns = [c.strip().upper() for c in marc.columns]
    zplm = zplm_marc_df.copy(); zplm.columns = [c.strip().upper() for c in zplm.columns]
    mara = mara_df.copy(); mara.columns = [c.strip().upper() for c in mara.columns]

    m_mat   = next((c for c in ["MATERIAL", "MATNR"] if c in marc.columns), None)
    m_plant = next((c for c in ["PLANT", "WERKS"] if c in marc.columns), None)
    z_mat   = next((c for c in ["MATERIAL", "MATNR"] if c in zplm.columns), None)
    z_plant = next((c for c in ["PLANT", "WERKS"] if c in zplm.columns), None)
    mt_col  = next((c for c in ["MATERIAL TYPE", "MTART"] if c in mara.columns), None)
    mara_mat = next((c for c in ["MATERIAL", "MATNR"] if c in mara.columns), None)

    if not all([m_mat, m_plant, z_mat, z_plant, mt_col, mara_mat]):
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    marc[m_mat]   = marc[m_mat].astype(str).str.strip()
    marc[m_plant] = marc[m_plant].astype(str).str.strip().str.upper()
    zplm[z_mat]   = zplm[z_mat].astype(str).str.strip()
    zplm[z_plant] = zplm[z_plant].astype(str).str.strip().str.upper()
    mara[mara_mat] = mara[mara_mat].astype(str).str.strip()
    mara[mt_col]   = mara[mt_col].astype(str).str.strip().str.upper()

    valid_mats = set(mara[mara[mt_col].isin({"FERT", "HALB", "ROH"})][mara_mat])
    marc_keys = set(zip(marc[marc[m_mat].isin(valid_mats)][m_mat], marc[marc[m_mat].isin(valid_mats)][m_plant]))
    zplm_keys = set(zip(zplm[z_mat], zplm[z_plant]))
    missing = sorted(marc_keys - zplm_keys)

    return pd.DataFrame([{"Material": m, "Plant": p, "Error": "Missing in ZPLM MARC"} for m, p in missing],
                        columns=["Material", "Plant", "Error"])


def validar_zplm_marc_vs_marc_existence(zplm_marc_df, marc_df, mara_df):
    """Entries in ZPLM MARC that are not in MARC (for FERT/HALB/ROH materials)."""
    if zplm_marc_df is None or zplm_marc_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    zplm = zplm_marc_df.copy(); zplm.columns = [c.strip().upper() for c in zplm.columns]
    marc = marc_df.copy();      marc.columns = [c.strip().upper() for c in marc.columns]
    mara = mara_df.copy();      mara.columns = [c.strip().upper() for c in mara.columns]

    z_mat   = next((c for c in ["MATERIAL", "MATNR"] if c in zplm.columns), None)
    z_plant = next((c for c in ["PLANT", "WERKS"] if c in zplm.columns), None)
    m_mat   = next((c for c in ["MATERIAL", "MATNR"] if c in marc.columns), None)
    m_plant = next((c for c in ["PLANT", "WERKS"] if c in marc.columns), None)
    mara_mat = next((c for c in ["MATERIAL", "MATNR"] if c in mara.columns), None)
    mt_col   = next((c for c in ["MATERIAL TYPE", "MTART"] if c in mara.columns), None)

    if not all([z_mat, z_plant, m_mat, m_plant, mara_mat, mt_col]):
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    zplm[z_mat]  = zplm[z_mat].astype(str).str.strip()
    zplm[z_plant]= zplm[z_plant].astype(str).str.strip().str.upper()
    marc[m_mat]  = marc[m_mat].astype(str).str.strip()
    marc[m_plant]= marc[m_plant].astype(str).str.strip().str.upper()
    mara[mara_mat] = mara[mara_mat].astype(str).str.strip()
    mara[mt_col]   = mara[mt_col].astype(str).str.strip().str.upper()

    valid_mats = set(mara[mara[mt_col].isin({"FERT", "HALB", "ROH"})][mara_mat])
    marc_filtered = marc[marc[m_mat].isin(valid_mats)]
    extra = sorted(set(zip(zplm[z_mat], zplm[z_plant])) - set(zip(marc_filtered[m_mat], marc_filtered[m_plant])))

    return pd.DataFrame(
        [{"Material": m, "Plant": p, "Error": "Exists in ZPLM MARC but missing in MARC (FERT/HALB/ROH)"} for m, p in extra],
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
    MARC + MARA: Neg. stocks in plant validation for HALB and ROH.
    Must be marked for:
      1) PT01 + PSL=P105
      2) PT02 + PSL=P205 + PSMS not in {25,75}
      3) PT02 + PSL=P220 + PSMS in {25,75}
    Must NOT be marked in any other scenario.
    If the field is blank in a scenario where it must be marked -> specific blank error.
    """
    if marc_df is None or marc_df.empty or mara_df is None or mara_df.empty:
        return []

    needed = {'Material', 'Plant', 'Neg. stocks in plant',
              'Production Storage Location', 'Plant Specific Material Status'}
    if not needed.issubset(marc_df.columns) or not {'Material', 'Material Type'}.issubset(mara_df.columns):
        return []

    mt_map = dict(zip(mara_df['Material'].astype(str).str.strip(),
                      mara_df['Material Type'].astype(str).str.strip().upper()))

    def _s(x):
        return '' if pd.isna(x) else str(x).strip()

    def _is_blank(x):
        return _s(x) == ''

    def _is_marked(x):
        return _s(x).upper() in {'X', '1', 'YES', 'Y', 'TRUE'}

    def _psms_int(x):
        sx = _s(x)
        if sx == '':
            return None
        try:
            return int(float(sx))
        except Exception:
            return None

    erros = []
    for idx, row in marc_df.iterrows():
        material = _s(row.get('Material'))
        plant    = _s(row.get('Plant'))
        mt       = mt_map.get(material, '')

        if mt not in {'HALB', 'ROH'}:
            continue

        neg_raw  = row.get('Neg. stocks in plant')
        blank    = _is_blank(neg_raw)
        marked   = _is_marked(neg_raw)
        psl      = _s(row.get('Production Storage Location'))
        psms     = _psms_int(row.get('Plant Specific Material Status'))

        must_mark = (
            (plant == 'PT01' and psl == 'P105') or
            (plant == 'PT02' and psl == 'P205' and psms not in {25, 75}) or
            (plant == 'PT02' and psms in {25, 75} and psl == 'P220')
        )

        if must_mark and blank:
            erros.append({'Line': idx + 2, 'Field': 'Neg. stocks in plant',
                          'Error': f"Mandatory field blank — must be marked for this combination "
                                   f"(Plant={plant}, PSL={psl}, PSMS={psms})",
                          'Material': material, 'Plant': plant})
        elif must_mark and not marked:
            erros.append({'Line': idx + 2, 'Field': 'Neg. stocks in plant',
                          'Error': f"Must be marked for this combination (Plant={plant}, PSL={psl}, PSMS={psms})",
                          'Material': material, 'Plant': plant})
        elif not must_mark and marked:
            erros.append({'Line': idx + 2, 'Field': 'Neg. stocks in plant',
                          'Error': "Invalid marking: allowed only for PT01+P105, PT02+P205 (PSMS≠25/75), or PT02+P220 (PSMS=25/75)",
                          'Material': material, 'Plant': plant})

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
        errors_df['View'] = errors_df['Field'].map(field_to_view).fillna('Not aplicable')
    else:
        errors_df['View'] = 'Not aplicable'

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
        if view not in ('', 'Not aplicable'):
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
        contagem_final_df['View']             = contagem_final_df['Field'].map(field_to_view).fillna('Not aplicable')
        contagem_final_df['View_Description'] = contagem_final_df['View'].map(view_to_desc).fillna('Not aplicable')

        if 'Role' in roles_df.columns:
            roles_df['Role']  = roles_df['Role'].astype(str).str.strip()
            field_to_role     = roles_df.set_index('Field')['Role'].to_dict()
            contagem_final_df['Role'] = contagem_final_df['Field'].map(field_to_role).fillna('Not aplicable')

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
                {'nan': 'Not aplicable', 'None': 'Not aplicable', '': 'Not aplicable'}
            )

    # -----------------------------------------------
    # 9) Coverage per View
    # -----------------------------------------------
    field_rules = {
        'Production Scheduling Profile':     lambda mt, pt: mt in ['FERT', 'HALB'],
        'Repetitive Manufacturing Profile':  lambda mt, pt: mt in ['FERT', 'HALB'],
        'Indicator: Repetitive Manufa. Allow': lambda mt, pt: True,
        'Enterproj Number':                  lambda mt, pt: mt == 'FERT',
        'Procurement Type':                  lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'] and pt in ['PP', 'MP'],
        'Account Assignment Group':          lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Automatic Purchase Order Allowed':  lambda mt, pt: pt == 'PP',
        'Batch Management Required Ind.':    lambda mt, pt: True,
        'Source List Required':              lambda mt, pt: pt == 'PP',
        'Purchasing Group':                  lambda mt, pt: pt == 'PP',
        'MRP Type':                          lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Planning Time Fence':               lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Lot Sizing Procedure':              lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Per.Lot Sz.Plng Cal.':              lambda mt, pt: mt in ['FERT', 'HALB'],
        'With Qty Structure':                lambda mt, pt: mt in ['FERT', 'HALB'],
        'Item Category Group':               lambda mt, pt: True,
        'General Item Category Group':       lambda mt, pt: True,
        'Item Category Group (MVKE)':        lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Loading Group':                     lambda mt, pt: True,
        'Division':                          lambda mt, pt: True,
        'Export/Import Group':               lambda mt, pt: True,
        'Product Hierarchy':                 lambda mt, pt: mt == 'FERT',
        'Price Control':                     lambda mt, pt: True,
        'Price Unit':                        lambda mt, pt: True,
        'Valuation Class':                   lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Gross Weight':                      lambda mt, pt: pt == 'FP',
        'Net Weight':                        lambda mt, pt: pt == 'FP',
        'Planning Calendar':                 lambda mt, pt: True,
        'Period Type':                       lambda mt, pt: True,
        'Strategy Group':                    lambda mt, pt: True,
        'Inspection Type (QMAT 08)':         lambda mt, pt: mt in ['HALB', 'ROH'],
        'MRP Controller':                    lambda mt, pt: True,
        'Plant Specific Material Status':    lambda mt, pt: True,
        'Backflush':                         lambda mt, pt: True,
        'Product-Related Origin':            lambda mt, pt: True,
        'Availability Check':                lambda mt, pt: True,
        'Neg. stocks in plant':              lambda mt, pt: mt in ['ROH', 'HALB'],
        'Consumption Mode':                  lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Bwd consumption per.':              lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Fwd consumption per.':              lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Profit Center':                     lambda mt, pt: True,
        'Material Group':                    lambda mt, pt: True,
        'JIT Delivery Sched.':               lambda mt, pt: pt == 'PP',
        'Issue Storage Location':            lambda mt, pt: True,
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
    qmat_file,
    roles_file=None,
    zplm_marc_file=None,
    zplm_mara_file=None,
    previous_report=None,
    plant_filter: str | None = None,
    progress_callback=None,
    csr_item_file=None,
    csr_head_file=None,
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

    _cb(14, "Loading QMAT...")
    qmat_df = load_xlsx_fast(qmat_file, sheet_name=0)

    if roles_file is None:
        roles_file = get_rolesviews_bytesio()
    rolesviews_dict = load_xlsx_fast(roles_file, sheet_name=None)

    _cb(16, "Loading ZPLM data...")
    zplm_marc_df = load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None

    _cb(18, "Preparing ZPLM data...")
    zplm_mara_df = load_xlsx_fast(zplm_mara_file, sheet_name=0) if zplm_mara_file else None

    _cb(19, "Loading CSR data...")
    csr_item_df = load_xlsx_fast(csr_item_file, sheet_name=0) if csr_item_file else None
    csr_head_df = load_xlsx_fast(csr_head_file, sheet_name=0) if csr_head_file else None

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
    if qmat_df is None:
        raise ValueError("QMAT.xlsx not loaded.")
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
        (47, lambda: validar_per_lot_sz_plng_cal(marc_df, resultados, mara_df)),
        (48, lambda: validar_material_group(sap_df, resultados)),
        (50, lambda: validar_mrp_type(sap_df, resultados, zplm_marc_df, csr_item_df, csr_head_df)),
        (52, lambda: validar_net_weight(sap_df, resultados)),
        (54, lambda: validar_period_type(sap_df, resultados)),
        (56, lambda: validar_planning_calendar(sap_df, resultados)),
        (58, lambda: validar_planning_time_fence(sap_df, resultados)),
        (60, lambda: validar_price_control(sap_df, resultados)),
        (62, lambda: validar_price_unit(sap_df, resultados)),
        (64, lambda: validar_product_hierarchy(sap_df, resultados, mvke_df, mara_df)),
        (66, lambda: validar_product_hierarchy_vs_enterproj(mara_df, zplm_mara_df, resultados)),
        (68, lambda: validar_production_scheduling_profile(sap_df, resultados)),
        (70, lambda: validar_profit_center(sap_df, resultados)),
        (72, lambda: validar_purchasing_group(sap_df, resultados)),
        (74, lambda: validar_repetitive_manufacturing(sap_df, resultados)),
        (76, lambda: validar_repetitive_manufacturing_profile(sap_df, resultados)),
        (78, lambda: validar_source_list_required(sap_df, resultados)),
        (80, lambda: validar_storage_location(sap_df, resultados)),
        (82, lambda: validar_valuation_class(sap_df, resultados)),
        (84, lambda: validar_with_qty_structure(sap_df, resultados)),
        (86, lambda: validar_enterproj_number(sap_df, resultados)),
        (88, lambda: validar_inspection_type_qmat(sap_df, marc_df, qmat_df, resultados)),
        (90, lambda: validar_mrp_controller(sap_df, resultados)),
        (92, lambda: validar_plant_specific_material_status(sap_df, resultados)),
        (94, lambda: validar_backflush(sap_df, resultados)),
        (96, lambda: validar_strategy_group(sap_df, resultados)),
        (98, lambda: validar_product_related_origin(sap_df, resultados)),
        (99, lambda: validar_availability_check(sap_df, resultados)),
    ]

    for pct, fn in validations:
        _cb(pct, "Running field validations...")
        fn()

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
            ("Field Errors", fe_cnt,                        "FCE4D6"),
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

    PSMS_NAME = {"40": "PreProd", "45": "Prod", "50": "Balout", "55": "Service"}
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

        dup50 = g[g["PSMS"] == "50"]
        if len(dup50) >= 2:
            g.loc[dup50.index[0], "Status"] = "Outdated"

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

    # Apply plant filter consistently to all DataFrames before passing to validators
    if plant_filter and 'Plant' in marc.columns:
        _pf = str(plant_filter).strip().upper()
        marc      = marc[marc['Plant'].astype(str).str.strip().str.upper() == _pf].copy()
        if zplm_marc is not None and 'Plant' in zplm_marc.columns:
            zplm_marc = zplm_marc[zplm_marc['Plant'].astype(str).str.strip().str.upper() == _pf].copy()
        if mbew is not None and 'Plant' in mbew.columns:
            mbew      = mbew[mbew['Plant'].astype(str).str.strip().str.upper() == _pf].copy()
        if ekpo is not None and 'Plant' in ekpo.columns:
            ekpo      = ekpo[ekpo['Plant'].astype(str).str.strip().str.upper() == _pf].copy()
        if mard is not None and 'Plant' in mard.columns:
            mard      = mard[mard['Plant'].astype(str).str.strip().str.upper() == _pf].copy()

    results = {}

    # FP Status
    df_fp, kpi_fp = run_fp_status_validation(
        mara_file=mara_file, marc_file=marc_file,
        zplm_marc_file=zplm_marc_file, zplm_mara_file=zplm_mara_file,
        plant_filter=plant_filter
    )
    results["FP Level Status"] = df_fp

    # Other validations
    results["FP in PT02"]                   = validar_fp_pt01_pt02_status55(zplm_marc, marc)
    results["EOPvsSL (Blank or Swapped)"]   = validar_fp55_eop_vs_service_liability(zplm_marc, marc)
    results["Past Service Liability"]        = validar_fp55_service_liability_expired(zplm_marc, marc)
    results["No Standard Cost"]             = validar_fppp70_standard_price_mbew(zplm_marc, marc, mbew)
    results["Missing in ZPLM_MARC"]         = validar_marc_vs_zplm_marc_existence(marc, zplm_marc, mara)
    results["Missing in ZPLM_MARA"]         = validar_mara_vs_zplm_mara_existence(mara, zplm_mara)
    results["Not in MARC"]                  = validar_zplm_marc_vs_marc_existence(zplm_marc, marc, mara)
    results["Not in MARA"]                  = validar_zplm_mara_vs_mara_existence(zplm_mara, mara)
    results["FP LTB Fields"]               = validar_ltb_fields(zplm_marc, marc, mara)
    results["PO's to Close"]               = validar_ekpo_delivery_completion(marc, mara, zplm_marc, ekpo)
    results["OBSOLETE with No Stock"]       = validar_mard_total_stock(mard, marc, mara)

    # Strip datetime (remove time component from date columns)
    def strip_datetime(df):
        if df is None or df.empty:
            return df
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.date
        return df

    for k in list(results.keys()):
        results[k] = strip_datetime(results[k])

    # Remove empty results
    cleaned = {}
    for name, df in results.items():
        if df is None or df.empty:
            continue
        if "Plant" in df.columns and plant_filter in ("PT01", "PT02"):
            dfp = df[df["Plant"] == plant_filter]
            if not dfp.empty:
                cleaned[name] = dfp
        else:
            cleaned[name] = df
    results = cleaned

    # ------------------------------------------------------------------
    # Write Excel
    # ------------------------------------------------------------------
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:

        # SUMMARY sheet
        summary_rows = []
        for sheet_name, df in results.items():
            if "Plant" not in df.columns:
                summary_rows.append({"Validation": sheet_name, "Plant": "ALL", "Issues": len(df)})
                continue
            for plant in ["PT01", "PT02"]:
                dfp = df[df["Plant"] == plant]
                if not dfp.empty:
                    summary_rows.append({"Validation": sheet_name, "Plant": plant, "Issues": len(dfp)})

        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="SUMMARY", index=False)

        # Detail sheets (split by plant if applicable)
        for sheet_name, df in results.items():
            if "Plant" not in df.columns:
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                continue
            for plant in ["PT01", "PT02"]:
                dfp = df[df["Plant"] == plant]
                if not dfp.empty:
                    dfp.to_excel(writer, sheet_name=f"{sheet_name}_{plant}"[:31], index=False)

    # Apply styling to SUMMARY
    out.seek(0)
    wb = load_workbook(out)
    if "SUMMARY" in wb.sheetnames:
        ws = wb["SUMMARY"]
        link_col   = ws.max_column + 1
        ws.cell(row=1, column=link_col, value="Link")

        header_fill = PatternFill(start_color="D9D9D9", fill_type="solid")
        zebra       = PatternFill(start_color="F7F7F7", fill_type="solid")
        border      = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"),  bottom=Side(style="thin")
        )

        for cell in ws[1]:
            cell.font      = Font(bold=True)
            cell.fill      = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border    = border

        for r in range(2, ws.max_row + 1):
            validation = str(ws.cell(row=r, column=1).value)
            plant      = str(ws.cell(row=r, column=2).value)
            sn         = f"{validation}_{plant}" if plant in ("PT01", "PT02") else validation
            safe       = sn.replace("'", "''")
            c          = ws.cell(row=r, column=link_col)
            if sn in wb.sheetnames:
                c.value = f'=HYPERLINK("#\'{safe}\'!A1","Open")'
                c.font  = Font(color="0563C1", underline="single")
            else:
                c.value = "Not found"
            c.alignment = Alignment(horizontal="center")
            c.border    = border

        for r in range(2, ws.max_row + 1):
            for c in ws[r]:
                if r % 2 == 0:
                    c.fill = zebra
                c.border = border

        for col_idx, col in enumerate(ws.columns, start=1):
            from openpyxl.utils import get_column_letter as _gcl
            max_len = max((len(str(cell.value or "")) for cell in col), default=8)
            ws.column_dimensions[_gcl(col_idx)].width = max_len + 2

    out2 = io.BytesIO()
    wb.save(out2)
    out2.seek(0)

    return out2, results, kpi_fp


# =============================================================================
# 5) MRP VALIDATION
# =============================================================================

def run_mrp_validation(mara_file, marc_file, zplm_marc_file=None,
                        *, plant_filter: str | None = None):
    """
    Validates MRP-related MARC fields:
    - Advanced Planning: HALB+Transfer / FERT -> 'true'; HALB internal / ROH -> 'false'
    - Heuristic / PP Plng Procedure / Requirement Strategy:
        ROH / HALB internal -> empty
        HALB+Transfer / FERT -> 'SAP_PP_002' / 4 / 20
    Returns: (excel_bytes, errors_df, summary_df)
    """
    mara_df     = load_xlsx_fast(mara_file, sheet_name=0)
    marc_df     = load_xlsx_fast(marc_file, sheet_name=0)
    zplm_marc_df = load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None

    if mara_df is None or marc_df is None:
        raise ValueError("MARA.xlsx / MARC.xlsx not loaded.")

    for df in (mara_df, marc_df):
        df.columns = [str(c).strip() for c in df.columns]

    if plant_filter and 'Plant' in marc_df.columns:
        pf = str(plant_filter).strip().upper()
        marc_df = marc_df[marc_df['Plant'].astype(str).str.strip().str.upper() == pf].copy()

    if not {'Material', 'Material Type'}.issubset(mara_df.columns):
        raise ValueError("MARA.xlsx must contain 'Material' and 'Material Type'.")

    mt_map = dict(zip(
        mara_df['Material'].astype(str).str.strip(),
        mara_df['Material Type'].astype(str).str.strip().str.upper()
    ))

    transfer_map = {}
    try:
        transfer_map = build_material_transfer_map(zplm_marc_df)
    except Exception:
        pass

    def _s(x):
        return '' if pd.isna(x) else str(x).strip()

    def pick_col(df, candidates):
        return next((c for c in candidates if c in df.columns), None)

    col_adv  = pick_col(marc_df, ['Advanced Planning'])
    col_heur = pick_col(marc_df, ['Heuristic'])
    col_pp   = pick_col(marc_df, ['PP Plng Procedure'])
    col_req  = pick_col(marc_df, ['Requirement Strategy'])

    for col, name in [(col_adv, 'Advanced Planning'), (col_heur, 'Heuristic'),
                      (col_pp, 'PP Plng Procedure'), (col_req, 'Requirement Strategy')]:
        if col is None:
            raise ValueError(f"MARC.xlsx must contain column '{name}'.")

    resultados = []
    for idx, row in marc_df.iterrows():
        material = _s(row.get('Material'))
        plant    = _s(row.get('Plant'))
        mt       = mt_map.get(material, '').upper()
        if mt not in {'FERT', 'HALB', 'ROH'}:
            continue

        is_transfer = transfer_map.get((material, plant), False)
        adv  = _s(row.get(col_adv)).lower()
        heur = _s(row.get(col_heur))
        pp   = normalize_code(row.get(col_pp), as_int=True) if col_pp else ''
        req  = normalize_code(row.get(col_req), as_int=True) if col_req else ''

        # Advanced Planning
        if (mt == 'HALB' and is_transfer) or mt == 'FERT':
            if adv != 'true':
                resultados.append({'Line': idx + 2, 'Field': 'Advanced Planning',
                                   'Error': "Invalid: should be 'true' (HALB with Transfer or FERT)",
                                   'Material': material, 'Plant': plant})
        elif (mt == 'HALB' and not is_transfer) or mt == 'ROH':
            if adv != 'false':
                resultados.append({'Line': idx + 2, 'Field': 'Advanced Planning',
                                   'Error': "Invalid: should be 'false' (HALB internal or ROH)",
                                   'Material': material, 'Plant': plant})

        # Heuristic / PP Plng Procedure / Requirement Strategy
        if mt == 'ROH' or (mt == 'HALB' and not is_transfer):
            for field, val, label in [(col_heur, heur, 'Heuristic'),
                                       (col_pp,   _s(row.get(col_pp)), 'PP Plng Procedure'),
                                       (col_req,  _s(row.get(col_req)), 'Requirement Strategy')]:
                if val != '':
                    resultados.append({'Line': idx + 2, 'Field': label,
                                       'Error': "Should be empty (ROH or HALB internal)",
                                       'Material': material, 'Plant': plant})
        else:
            if heur != 'SAP_PP_002':
                resultados.append({'Line': idx + 2, 'Field': 'Heuristic',
                                   'Error': "Invalid: should be 'SAP_PP_002'",
                                   'Material': material, 'Plant': plant})
            if pp != '4':
                resultados.append({'Line': idx + 2, 'Field': 'PP Plng Procedure',
                                   'Error': "Invalid: should be 4", 'Material': material, 'Plant': plant})
            if req != '20':
                resultados.append({'Line': idx + 2, 'Field': 'Requirement Strategy',
                                   'Error': "Invalid: should be 20", 'Material': material, 'Plant': plant})

    errors_df = _coerce_df_for_display(pd.DataFrame(resultados))

    # Post-processing with Views/Roles
    rolesviews_dict = load_xlsx_fast(get_rolesviews_bytesio(), sheet_name=None)
    roles_df = rolesviews_dict.get("Roles")
    views_df = rolesviews_dict.get("Views")

    (errors_df, _, _, contagem_final_df, coverage_df, view_tabs_dict) = \
        postprocess_with_views_roles_and_classification(
            errors_df=errors_df, mara_df=mara_df, marc_df=marc_df,
            roles_df=roles_df, views_df=views_df, plant_filter=plant_filter,
        )

    # Build enriched Original Data
    marc_enriched = marc_df.copy()
    marc_enriched['Material Type'] = marc_enriched['Material'].astype(str).str.strip().map(mt_map).fillna('')
    out_cols = ['Material', 'Plant', 'Material Type',
                'Advanced Planning', 'Heuristic', 'PP Plng Procedure', 'Requirement Strategy']
    original_df = marc_enriched[[c for c in out_cols if c in marc_enriched.columns]].copy()

    file_bytes = io.BytesIO()
    with pd.ExcelWriter(file_bytes, engine="openpyxl") as writer:
        original_df.to_excel(writer, sheet_name="Original Data", index=False)
        errors_df.to_excel(writer, sheet_name="All Errors", index=False)
        base_row = 1
        if not contagem_final_df.empty:
            contagem_final_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=1)
            base_row = len(contagem_final_df) + 5
        if not coverage_df.empty:
            coverage_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=base_row)
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


def run_plm_workflow_combined(auto_file, plm_file, timing_file):
    """
    ATL/PLM workflow: combines Automotive + TimingLetter validation reports.
    Maps Trigger ID -> Change Number via Timing Letters file.
    Returns: (excel_bytes, df_pt01, df_pt02)
    """
    df_auto   = load_xlsx_fast(auto_file,   sheet_name=0)
    df_plm    = load_xlsx_fast(plm_file,    sheet_name=0)
    df_timing = load_xlsx_fast(timing_file, sheet_name=0)

    if df_auto is None or df_plm is None or df_timing is None:
        raise ValueError("ATL files missing: Automotive / TimingLetter / Timing Letters")

    required_base = ['Plant', 'Approval Status', 'Trigger ID', 'Creation Date']
    for col in required_base:
        if col not in df_auto.columns:
            raise ValueError(f"Column '{col}' missing in Automotive file.")
        if col not in df_plm.columns:
            raise ValueError(f"Column '{col}' missing in TimingLetter PLM file.")
    if 'Timing Letter Number' not in df_timing.columns or 'Change Number' not in df_timing.columns:
        raise ValueError("Timing Letters file must contain 'Timing Letter Number' and 'Change Number'.")

    # Statuses that disqualify a Timing Letter from appearing in the report
    _EXCLUDED_STATUSES = {'Complete', 'Rejected'}
    _status_col = next(
        (col for col in df_timing.columns
         if col.strip().lower() in ('timing letter status', 'status')),
        None
    )
    if _status_col:
        _active_timing = df_timing[
            ~df_timing[_status_col].astype(str).str.strip().isin(_EXCLUDED_STATUSES)
        ].copy()
    else:
        _active_timing = df_timing.copy()

    # Map Trigger ID -> Change Number  (active Timing Letters only)
    map_trigger_to_change = dict(zip(
        _active_timing['Timing Letter Number'].astype(str).str.strip(),
        _active_timing['Change Number'].astype(str).str.strip()
    ))
    # Reverse map: Change Number -> Timing Letter Number (active only)
    map_change_to_tln = dict(zip(
        _active_timing['Change Number'].astype(str).str.strip(),
        _active_timing['Timing Letter Number'].astype(str).str.strip()
    ))

    # Keep original Trigger ID before replacing with Change Number
    df_plm['_original_trigger'] = df_plm['Trigger ID'].astype(str).str.strip()
    df_plm['Trigger ID'] = df_plm['_original_trigger'].apply(
        lambda x: map_trigger_to_change.get(x, x)
    )
    # Populate Timing Letter Number — active Timing Letters only
    df_plm['Timing Letter Number'] = df_plm['_original_trigger'].apply(
        lambda x: x if x in map_trigger_to_change else map_change_to_tln.get(x, "")
    )
    df_plm.drop(columns=['_original_trigger'], inplace=True)

    def add_pending_claimed(df):
        actual_end_cols = [c for c in df.columns if c.startswith("Actual End Date of Work Item")]
        df['Pending(Claimed)'] = ""
        for idx, row in df.iterrows():
            pending = []
            for col in actual_end_cols:
                date_val = row.get(col)
                if pd.notna(date_val):
                    s = str(date_val)
                    if s.startswith("9999-12-31") or s.startswith("12/31/9999"):
                        col_idx = df.columns.get_loc(col)
                        if col_idx > 0:
                            approver = row[df.columns[col_idx - 1]]
                            if pd.notna(approver):
                                pending.append(str(approver))
            if pending:
                df.at[idx, 'Pending(Claimed)'] = ", ".join(pending)
        return df

    df_auto = add_pending_claimed(df_auto)
    df_plm  = add_pending_claimed(df_plm)

    # Automotive
    categories_auto = ["Pricing", "Purchasing", "MPL", "Customs", "Sales", "Manufacturing", "Costing", "MDS"]
    approver_cols_auto = ['Approver1'] + [f'approver{i}' for i in range(2, 8)] + ['Approver8']
    df_auto_f = df_auto[df_auto['Approval Status'].isin(['INAPPROVAL', 'REJECTED'])].copy()
    df_auto_f['Creation Date'] = pd.to_datetime(df_auto_f['Creation Date'], errors='coerce')
    df_auto_f.sort_values('Creation Date', ascending=False, inplace=True)
    df_auto_f.drop_duplicates(subset=['Trigger ID'], keep='first', inplace=True)
    for cat in categories_auto:
        df_auto_f[cat] = ""
    for idx, row in df_auto_f.iterrows():
        for col in approver_cols_auto:
            if col in df_auto_f.columns:
                clean = str(row[col]).split("(")[0].strip() if pd.notna(row[col]) else ""
                if clean in categories_auto:
                    df_auto_f.at[idx, clean] = "X"
    df_auto_f['Creation Date'] = df_auto_f['Creation Date'].dt.strftime('%Y-%m-%d')
    df_auto_f['Source'] = "Automotive"

    # TimingLetter
    df_tl_f = df_plm[df_plm['Approval Status'].isin(['INAPPROVAL', 'REJECTED'])].copy()
    df_tl_f['Creation Date'] = pd.to_datetime(df_tl_f['Creation Date'], errors='coerce')
    df_tl_f.sort_values('Creation Date', ascending=False, inplace=True)
    df_tl_f.drop_duplicates(subset=['Trigger ID'], keep='first', inplace=True)
    df_tl_f['MPL - TL'] = ""; df_tl_f['MDS - TL'] = ""
    for idx, row in df_tl_f.iterrows():
        for col in [c for c in df_plm.columns if "approver" in c.lower()]:
            val = str(row[col]) if pd.notna(row.get(col)) else ""
            if val.startswith('MPL'): df_tl_f.at[idx, 'MPL - TL'] = "X"
            if val.startswith('MDS'): df_tl_f.at[idx, 'MDS - TL'] = "X"
    df_tl_f['Creation Date'] = df_tl_f['Creation Date'].dt.strftime('%Y-%m-%d')
    df_tl_f['Source'] = "TimingLetter"

    all_cols = ['Source', 'Trigger ID', 'Timing Letter Number', 'Plant', 'Creation Date',
                'Approval Status', 'Pending(Claimed)',
                'MPL - TL', 'MDS - TL', 'Pricing', 'Purchasing', 'MPL', 'Customs',
                'Sales', 'Manufacturing', 'Costing', 'MDS']
    for col in all_cols:
        if col not in df_auto_f.columns: df_auto_f[col] = ""
        if col not in df_tl_f.columns:   df_tl_f[col]   = ""

    df_final = pd.concat([df_auto_f[all_cols], df_tl_f[all_cols]], ignore_index=True)

    # Propagate Timing Letter Number from TimingLetter rows to Automotive rows
    # with the same Change Number (Trigger ID)
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

    df_final['Creation Date'] = pd.to_datetime(df_final['Creation Date'], errors='coerce')
    df_final['Source'] = pd.Categorical(df_final['Source'], categories=["Automotive", "TimingLetter"], ordered=True)
    df_final.sort_values(["Source", "Creation Date"], ascending=[True, False], inplace=True)

    today = datetime.now()
    df_final['Aging'] = (today - df_final['Creation Date']).dt.days
    df_final['Creation Date'] = df_final['Creation Date'].dt.strftime('%Y-%m-%d')

    final_cols = ["Source", "Trigger ID", "Timing Letter Number", "Plant", "Creation Date",
                  "Approval Status", "Pending(Claimed)",
                  "MPL - TL", "MDS - TL", "Pricing", "Purchasing", "MPL", "Customs",
                  "Sales", "Manufacturing", "Costing", "MDS", "Aging"]
    df_final   = df_final[final_cols]
    df_pt01    = df_final[df_final['Plant'] == 'PT01'].copy()
    df_pt02    = df_final[df_final['Plant'] == 'PT02'].copy()

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
                if mat and view and view not in ('', 'Not aplicable'):
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


def save_run_log(payload: dict, *, base_dir: str = ".history") -> str | None:
    """Saves a log payload as a JSON file under .history/<run_type>/."""
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
        return path
    except Exception:
        return None


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
    "validar_inspection_type_qmat",

    # --- Part 2: Data Cleanup validations ---
    "validar_fp_pt01_pt02_status55",
    "validar_fp55_eop_vs_service_liability",
    "validar_fp55_service_liability_expired",
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
]