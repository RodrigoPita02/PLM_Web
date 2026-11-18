"""Regras A) — validações ao nível do campo (export SAPUI5)."""

import re
import pandas as pd

from engines.common.helpers import _norm_str, normalize_code
from engines.material_rules.batch_utils import get_batch_management_effective_material_type
from engines.material_rules.sap_like import (
    build_csr_exempt_set,
    build_material_transfer_map,
    get_effective_material_type,
)

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
          - ROH with Plant Type PP / COI -> must be marked ('Yes' / 1)
          - FERT and HALB (any Plant Type) -> must NOT be marked ('No' / 0)
          - ROH with any other Plant Type -> must NOT be marked ('No' / 0)

    Exception: the materials in BATCH_MANAGEMENT_HALB_LIKE_MATERIALS, when
    ROH with Plant Type PP / COI, are treated as HALB (must NOT be marked).

    The TN02 rule takes precedence over the Material Type / Plant Type rules.
    """
    col = 'Batch Management Required Ind.'

    required_columns = {
        'Plant',
        'Plant Type',
        'Material Type',
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

        plant = str(row.get('Plant', '')).strip().upper()
        plant_type = str(row.get('Plant Type', '')).strip().upper()

        raw_batch_ind = row.get(col)
        batch_ind = (
            ''
            if pd.isna(raw_batch_ind)
            else str(raw_batch_ind).strip().upper()
        )

        # Normalise: '1' / 'X' = marked, '0' = not marked
        is_marked = batch_ind in ('1', 'X', '1.0')
        is_unmarked = batch_ind in ('0', '0.0')

        real_mt = str(row.get('Material Type', '')).strip().upper()
        material_type = get_batch_management_effective_material_type(
            row.get('Material'),
            real_mt,
            plant_type,
        )
        exception_applied = (material_type != real_mt)

        if plant == 'TN02':
            # TN02 must always be empty (takes precedence)
            if batch_ind != '':
                erros.append(
                    "For Plant 'TN02', should be blank "
                    f"(Current: '{batch_ind}')"
                )

        elif material_type == 'ROH' and plant_type in ('PP', 'COI'):
            # ROH + PP/COI -> must be marked
            if batch_ind == '':
                erros.append('Mandatory field blank')
            elif not is_marked:
                erros.append(
                    f"For Material Type 'ROH' with Plant Type '{plant_type}', "
                    f"should be 'Yes' (Current: '{batch_ind}')"
                )

        else:
            # FERT / HALB (and ROH outside PP/COI) -> must NOT be marked
            mt_label = "ROH (treated as HALB)" if exception_applied else material_type
            if batch_ind == '':
                erros.append('Mandatory field blank')
            elif not is_unmarked:
                erros.append(
                    f"For Material Type '{mt_label}' "
                    f"with Plant Type '{plant_type}', "
                    f"should be 'No' (Current: '{batch_ind}')"
                )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' | '.join(erros),
                'Material': str(row.get('Material', '')).strip(),
                'Plant': plant,
            })


def validar_batch_management_mara_global(sap_df, mara_df, resultados):
    """
    MARA (global) 'Batch Management' — material level, same rule as plant level.

    A material must have the global flag marked when at least one of its plants
    requires it at plant level (ROH with Plant Type PP / COI, TN02 excluded).
    Otherwise (FERT, HALB, ROH without PP/COI plants, and the materials in
    BATCH_MANAGEMENT_HALB_LIKE_MATERIALS when ROH with PP/COI) it must NOT be marked.

    Reported once per material (Plant blank), Field 'Batch Management (Global)',
    so it never collides with the plant-level error of the same material.
    """
    field = 'Batch Management (Global)'

    if sap_df is None or sap_df.empty or mara_df is None or mara_df.empty:
        return
    if 'Batch Management' not in mara_df.columns or 'Material' not in mara_df.columns:
        print("⚠️ MARA has no 'Batch Management' column - global Batch Management validation skipped.")
        return
    if not {'Material', 'Plant', 'Plant Type', 'Material Type'}.issubset(sap_df.columns):
        return

    def _flag(v):
        v = '' if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip().upper()
        return 'X' if v in ('X', '1', '1.0') else ('' if v in ('', '0', '0.0', 'NAN', 'NONE') else v)

    mara = mara_df.copy()
    mara['Material'] = mara['Material'].fillna('').astype(str).str.strip()
    mara = mara.drop_duplicates(subset=['Material'], keep='first')
    mara_flag = dict(zip(mara['Material'], mara['Batch Management'].map(_flag)))
    mara_type = (
        dict(zip(mara['Material'], mara['Material Type'].fillna('').astype(str).str.strip().str.upper()))
        if 'Material Type' in mara.columns else {}
    )
    mara_line = {m: i + 2 for i, m in enumerate(mara['Material'])}

    plants_requiring = {}   # material -> sorted plants that require batch management
    seen_materials = set()
    for _, row in sap_df.iterrows():
        material = _norm_str(row.get('Material'))
        plant = _norm_str(row.get('Plant')).upper()
        if not material or plant == 'TN02':
            continue
        seen_materials.add(material)
        plant_type = _norm_str(row.get('Plant Type')).upper()
        eff_mt = get_batch_management_effective_material_type(
            material, _norm_str(row.get('Material Type')).upper(), plant_type
        )
        if eff_mt == 'ROH' and plant_type in ('PP', 'COI'):
            plants_requiring.setdefault(material, set()).add(plant)

    for material in sorted(seen_materials):
        if material not in mara_flag:
            continue
        current = mara_flag[material]
        required = material in plants_requiring
        mt = mara_type.get(material, '')

        if required and current != 'X':
            plants = ', '.join(sorted(plants_requiring[material]))
            error = (
                "Mandatory field blank " if current == '' else f"Current: '{current}' "
            ) + (
                "- Batch Management (Global) should be 'Yes' "
                f"(ROH with Plant Type PP/COI in {plants})"
            )
        elif (not required) and current == 'X':
            error = (
                "Batch Management (Global) should be 'No' "
                f"(Material Type '{mt}' without any ROH PP/COI plant)"
            )
        else:
            continue

        resultados.append({
            'Line': mara_line.get(material, ''),
            'Field': field,
            'Error': error,
            'Material': material,
            'Plant': '',
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
