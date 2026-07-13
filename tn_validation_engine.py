# =============================================================================
# TN VALIDATION ENGINE
# =============================================================================
# Tunisia validation for plants TN01 and TN02.
#
# DESIGN:
#   1. Build a `sap_like_df` by merging MARC + MARA + ZPLM MARC.
#      This replaces the SAPUI5 Export (which doesn't exist for Tunisia).
#      All column names are mapped to match what the Palmela functions expect.
#
#   2. Run ALL Palmela validation functions using sap_like_df.
#      Rules that reference PT01/PT02 in their logic will simply not fire
#      (no rows match those plants), so they are inherently safe to call.
#
#   3. For rules that have TN-specific EXCEPTIONS, we use dedicated TN
#      replacement functions instead of the Palmela ones for those fields.
#
#   This guarantees Tunisia gets 100% of Palmela rules, plus TN exceptions,
#   with no duplication.
#
# TN-SPECIFIC EXCEPTIONS (from Rule Book):
#   3.1  Backflush          — TN02: must be BLANK (Palmela: mandatory)
#   3.1  Strategy Group     — TN01/TN02: must be 'Z1' FERT/HALB (Palmela: '40')
#   3.1  Availability Check — TN01/TN02: must be 'KP' (Palmela: just mandatory)
#   3.2  Procurement Type   — TN02 FERT/HALB: must be 'F'
#   3.3  Auto Purchase Order— TN01/TN02 FP: must be 'Yes' (Palmela: PP only)
#   3.5  Consumption Mode   — TN02 HALB: must be '0' or empty
#   3.10 JIT Delivery Sched.— TN01/TN02 FP: must be 'Yes' (Palmela: PP only)
#   3.14 MRP Type           — TN02 FERT/HALB: must be 'PD'
#                             PSMS=25 Planning Calendar: TN01->TN1, TN02->TN2
#   3.16 Planning Calendar  — TN01 FP/MP: 'TN1'; TN02 FP/MP: 'TN2'
#   3.17 Planning Time Fence— TN02 FERT: must be 0 or empty
#   3.21 Prod.Sched.Profile — TN02: must be BLANK (Palmela: 'ZREM01')
#   3.22 Profit Center      — TN01/TN02: must be 'TN00413'
#   3.23 Purchasing Group   — TN01/TN02 FP: must be 'STO' (Palmela: FP blank)
#   3.24 Repetitive Mfg.    — TN02: must be BLANK
#   3.25 Repetitive Mfg. Profile — TN02: must be BLANK
#   3.26 Source List        — TN01/TN02 FP: must be 'Yes' (Palmela: PP only)
#   3.27 Issue Storage Loc. — TN01: T105/T120; TN02: T220
#   3.31 Special Proc. Type — TN02 FP: must be 'Z3'
#   4.1  Account Assignment — TN FERT: TN00/DC10 + TN00/DC20
#   4.10 Storage Loc Ext.Pr — TN02 FP: mandatory (Palmela: PP only)
#   4.11 Indiv/Coll.Req.    — TN01/TN02 HALB/ROH: must be '2' (Palmela: '1')
#   4.13 Production Superv. — TN02: must be BLANK (Palmela: 'Z01')
#   3.32 Advanced Planning  — TN01 FERT, TN02 FERT/HALB: must be 'X'
# =============================================================================

import io
import pandas as pd
import numpy as np

from utils.fast_loader import load_xlsx_fast
from rolesviews_data import get_rolesviews_bytesio as _get_rolesviews_bytesio_raw


def get_rolesviews_bytesio():
    """
    Defensive wrapper around rolesviews_data.get_rolesviews_bytesio().
    See sap_validation_engine.py for the full explanation: rewinds the
    stream to position 0 in case it's a cached/shared BytesIO whose
    cursor is already at EOF from an earlier read elsewhere in the
    same process.
    """
    buf = _get_rolesviews_bytesio_raw()
    if hasattr(buf, "seek"):
        try:
            buf.seek(0)
        except Exception:
            pass
    return buf

from sap_validation_engine import (
    normalize_code,
    _norm_str,
    ensure_basic_data_k,
    build_material_transfer_map,
    build_csr_exempt_set,
    get_effective_material_type,
    _prepare_mvke,
    _coerce_df_for_display,
    _style_sap_report,
    postprocess_with_views_roles_and_classification,
    add_old_material_number_errors,
    validar_neg_stocks_in_plant,
    fill_plant_and_validate_enterproj_from_zplm,
    validar_per_lot_sz_plng_cal,
    validar_consumption_per,
    validar_marm_pc_ea_roh_halb,
    validar_lifo_fifo_relevant_mbew,
    validar_stock_determination_group_marc,
    validar_packaging_material_type_mara,
    validar_batch_management,
    validar_batch_management_plant_marc,
    validar_version_indicator_marc,
    validar_country_region_origin_fp,
    validar_product_hierarchy_vs_enterproj,
    validar_campo_obrigatorio,
    validar_division,
    validar_export_import_group,
    validar_general_item_category_group,
    validar_gross_weight,
    validar_net_weight,
    validar_loading_group,
    validar_lot_sizing_procedure,
    validar_material_group,
    validar_material_group_appoint,
    validar_period_type,
    validar_price_control,
    validar_price_unit,
    validar_product_hierarchy,
    validar_valuation_class,
    validar_with_qty_structure,
    validar_product_related_origin,
    validar_mrp_controller,
    validar_plant_specific_material_status,
    validar_item_category_group,
    validar_enterproj_number,
    validar_purchasing_group_dxx,
    merge_assignments,
    build_assignments_sheet,
    save_open_assignments,
    # New functions added to Palmela engine
    validar_costing_lot_size,
    validar_safety_time_marc,
    validar_material_price_determination_control,
    # Shared helpers (moved here so both engines share a single implementation)
    _normalize_columns,
    _normalize_df_columns,
    build_sap_like_df,
    _mt_map_from_mara,
    _pt_map_from_zplm,
    _filter_valid_material_types,
    _get_mps99_material_plant_pairs,
    _exclude_material_plant_status_99,
)

TN_PLANTS = {"TN01", "TN02"}


# =============================================================================
# STEP 2 — TN-specific replacement functions
# =============================================================================

def tn_validar_backflush(df, resultados):
    """TN01: mandatory. TN02: must be BLANK."""
    col = 'Backflush'
    if col not in df.columns:
        return
    for idx, row in df.iterrows():
        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        val   = _norm_str(row.get(col))
        if plant == 'TN01':
            if val == '':
                resultados.append({'Line': idx + 2, 'Field': col,
                                   'Error': 'Mandatory field blank',
                                   'Material': mat, 'Plant': plant})
        elif plant == 'TN02':
            if val != '':
                resultados.append({'Line': idx + 2, 'Field': col,
                                   'Error': f"TN02 should be blank (current: '{val}')",
                                   'Material': mat, 'Plant': plant})


def tn_validar_strategy_group(df, mara_df, resultados):
    """TN01/TN02 FERT and TN02 HALB: must be 'Z1'. Others: must be '40'."""
    col = 'Planning Strategy Group'
    if col not in df.columns:
        return
    mt_map = _mt_map_from_mara(mara_df)
    for idx, row in df.iterrows():
        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        mt    = mt_map.get(mat, '')
        raw   = row.get(col)
        val   = normalize_code(raw, as_int=True)
        tn_exc = (plant in TN_PLANTS and mt == 'FERT') or (plant == 'TN02' and mt == 'HALB')
        if val == '':
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': 'Mandatory field blank',
                               'Material': mat, 'Plant': plant})
        elif tn_exc:
            if val != 'Z1':
                resultados.append({'Line': idx + 2, 'Field': col,
                                   'Error': f"TN01/TN02 FERT and TN02 HALB must be 'Z1' (current: '{raw}')",
                                   'Material': mat, 'Plant': plant})
        else:
            if val != '40':
                resultados.append({'Line': idx + 2, 'Field': col,
                                   'Error': f"Invalid value: '{raw}' (should be '40')",
                                   'Material': mat, 'Plant': plant})


def tn_validar_availability_check(df, resultados):
    """TN01/TN02: mandatory AND must be 'KP'."""
    col = 'Availability Check'
    if col not in df.columns:
        return
    for idx, row in df.iterrows():
        mat = _norm_str(row.get('Material'))
        plant = _norm_str(row.get('Plant')).upper()
        val = _norm_str(row.get(col)).upper()
        if val == '':
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': "Mandatory field blank, must be 'KP'",
                               'Material': mat, 'Plant': plant})
        elif val != 'KP':
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': f"Must be 'KP' (current: '{val}')",
                               'Material': mat, 'Plant': plant})


def tn_validar_procurement_type(df, mara_df, resultados):
    """Palmela base rules + TN02 FERT/HALB must be 'F'."""
    col = 'Procurement Type'
    if col not in df.columns or 'Plant Type' not in df.columns:
        return
    mt_map    = _mt_map_from_mara(mara_df)
    universal = {'E', 'F', 'X'}

    for idx, row in df.iterrows():
        erros = []
        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        mt    = mt_map.get(mat, '')
        raw   = row.get(col)
        pt    = '' if pd.isna(raw) else str(raw).strip().upper()
        ptype = _norm_str(row.get('Plant Type')).upper()

        if pt == '':
            erros.append('Mandatory field blank')
        elif pt not in universal:
            erros.append(f"Invalid value: '{raw}' (allowed: E, F, X)")

        if pt in universal:
            if plant == 'TN02' and mt in {'FERT', 'HALB'}:
                if pt != 'F':
                    erros.append(f"For TN02 FERT/HALB, Procurement Type should be 'F' (current: '{pt}')")
            else:
                if ptype == 'PP':
                    if pt not in {'F', 'X'}:
                        erros.append(f"For Plant Type 'PP', should be 'F' or 'X' (current: '{pt}')")
                elif ptype in {'MP', 'FP'}:
                    if pt not in {'E', 'X'}:
                        erros.append(f"For Plant Type '{ptype}', should be 'E' or 'X' (current: '{pt}')")
                elif ptype == '':
                    if mt in {'FERT', 'HALB'} and pt not in {'E', 'X'}:
                        erros.append(f"For Material Type '{mt}', should be 'E' or 'X' (current: '{pt}')")
                    elif mt == 'ROH' and pt not in {'F', 'X'}:
                        erros.append(f"For Material Type 'ROH', should be 'F' or 'X' (current: '{pt}')")

        if erros:
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': ' \n'.join(erros), 'Material': mat, 'Plant': plant})


def tn_validar_automatic_purchase_order(df, resultados):
    """PP or FP -> 'X'; others -> empty (''), SAP checkbox logic."""
    
    col = 'Automatic Purchase Order Allowed'
    
    if col not in df.columns or 'Plant Type' not in df.columns:
        return

    for idx, row in df.iterrows():
        erros = []
        
        plant = _norm_str(row.get('Plant')).upper()
        ptype = _norm_str(row.get('Plant Type')).upper()
        val   = _norm_str(row.get(col)).upper()
        mat   = _norm_str(row.get('Material'))

        # ✅ Normalizar valores possíveis
        if val in {'X', 'YES', '1'}:
            val_norm = 'X'
        else:
            val_norm = ''

        # ✅ Regra principal
        if ptype in {'PP', 'FP'}:
            if val_norm != 'X':
                erros.append(
                    f"For Plant Type '{ptype}', should be 'X' (current: '{val}')"
                )
        else:
            if val_norm != '':
                erros.append(
                    f"For Plant Type '{ptype}', should be blank (current: '{val}')"
                )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' | '.join(erros),
                'Material': mat,
                'Plant': plant
            })

def tn_validar_consumption_mode(df, mara_df, resultados):
    """FERT -> 0 or empty. TN02 HALB -> 0 or empty. Others -> 2."""
    col = 'Consumption Mode'
    if col not in df.columns:
        return
    mt_map = _mt_map_from_mara(mara_df)
    for idx, row in df.iterrows():
        erros = []
        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        mt    = mt_map.get(mat, '')
        raw   = row.get(col)
        val   = normalize_code(raw, as_int=True)
        if mt == 'FERT':
            if val not in ('', '0'):
                erros.append(f"Invalid value: '{raw}' (FERT requires Consumption Mode = '0' or empty)")
        elif plant == 'TN02' and mt == 'HALB':
            if val not in ('', '0'):
                erros.append(f"Invalid value: '{raw}' (TN02 HALB requires Consumption Mode = '0' or empty)")
        else:
            if val == '':
                erros.append('Mandatory field blank')
            elif val != '2':
                erros.append(f"Invalid value: '{raw}' (should be '2' for non-FERT materials)")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': '\n'.join(erros), 'Material': mat, 'Plant': plant})


def tn_validar_jit_delivery_sched(df, resultados):
    """PP or FP -> '1'; others -> empty."""
    col = 'JIT Delivery Sched.'
    if col not in df.columns or 'Plant Type' not in df.columns:
        return
    for idx, row in df.iterrows():
        erros = []
        plant = _norm_str(row.get('Plant')).upper()
        ptype = _norm_str(row.get('Plant Type')).upper()
        raw   = row.get(col)
        val   = _norm_str(raw)
        mat   = _norm_str(row.get('Material'))
        if ptype in {'PP', 'FP'}:
            if val == '' or pd.isna(raw):
                erros.append(f"Mandatory field blank for Plant Type '{ptype}'")
            elif val not in ('1', '1.0'):
                erros.append(f"For Plant Type '{ptype}', should be 'Yes' (current: '{val}')")
        else:
            if val != '' and not pd.isna(raw):
                erros.append(f"For Plant Type '{ptype}', field must be empty (current: '{val}')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': ' \n '.join(erros), 'Material': mat, 'Plant': plant})


def tn_validar_mrp_type(df, mara_df, resultados, zplm_marc_df=None,
                         csr_item_df=None, csr_head_df=None):
    """Palmela base + TN02 FERT/HALB must be 'PD'. PSMS=25 cal: TN1/TN2."""
    needed = {'MRP Type', 'Plant Type', 'Planning Time Fence', 'Plant'}
    if not needed.issubset(df.columns):
        return

    mt_map       = _mt_map_from_mara(mara_df)
    has_psms     = 'Plant Specific Material Status' in df.columns
    has_pc       = 'Planning Calendar' in df.columns
    transfer_map = build_material_transfer_map(zplm_marc_df) if zplm_marc_df is not None else {}
    csr_exempt   = build_csr_exempt_set(csr_item_df, csr_head_df)
    valores_ok   = {'M1', 'PD', 'P1', 'ND', 'V1'}

    def _tf(v):
        if pd.isna(v) or str(v).strip() == '':
            return None
        try:
            return float(str(v).strip())
        except Exception:
            return 'INVALID'

    def _st(v):
        if pd.isna(v) or str(v).strip() == '':
            return None
        try:
            return int(float(str(v).strip()))
        except Exception:
            return None

    for idx, row in df.iterrows():
        erros   = []
        mrp     = _norm_str(row.get('MRP Type')).upper()
        pt      = _norm_str(row.get('Plant Type')).upper()
        plant   = _norm_str(row.get('Plant'))
        mat     = _norm_str(row.get('Material'))
        mt      = mt_map.get(mat, '').upper()
        tf_raw  = row.get('Planning Time Fence')
        tf_val  = _tf(tf_raw)
        psms    = _st(row.get('Plant Specific Material Status')) if has_psms else None
        pc      = _norm_str(row.get('Planning Calendar')).upper() if has_pc else ''
        is_tr   = transfer_map.get((mat, plant), False)
        has_csr = (mat, plant) in csr_exempt

        if psms == 10:
            continue

        if psms == 25:
            if mrp != 'ND':
                erros.append("Mandatory field blank (MRP Type)" if mrp == ''
                             else f"For Part Status = 25, MRP Type must be 'ND' (current: '{mrp}')")
            if tf_val not in (None, 0):
                erros.append(f"For Part Status = 25, Time Fence must be 0 or empty (current: '{tf_raw}')")
            if has_pc:
                expected = 'TN1' if plant == 'TN01' else 'TN2'
                if pc != expected:
                    erros.append(f"For Part Status = 25 & {plant}, Planning Calendar must be '{expected}' (current: '{pc}')")
            if erros:
                resultados.append({'Line': idx + 2, 'Field': 'MRP Type',
                                   'Error': ' \n'.join(erros), 'Material': mat, 'Plant': plant})
            continue

        if psms in (70, 75, 99):
            if mrp != 'ND':
                erros.append(f"For Part Status = {psms}, MRP Type must be 'ND' (current: '{mrp}')")
            if erros:
                resultados.append({'Line': idx + 2, 'Field': 'MRP Type',
                                   'Error': ' \n'.join(erros), 'Material': mat, 'Plant': plant})
            continue

        if psms == 65 and mrp == 'ND':
            continue

        # TN02 FERT/HALB must be 'PD'
        if plant == 'TN02' and mt in {'FERT', 'HALB'}:
            if mrp != 'PD':
                resultados.append({'Line': idx + 2, 'Field': 'MRP Type',
                                   'Error': f"TN02 FERT/HALB should be 'PD' (current: '{mrp}')",
                                   'Material': mat, 'Plant': plant})
            continue

        if mrp == '':
            erros.append('Mandatory field blank (MRP Type)')
        elif mrp not in valores_ok:
            erros.append(f"Invalid MRP Type: '{mrp}'")

        has_tf = (tf_val is not None) and (tf_val > 0)
        no_tf  = (tf_val is None) or (tf_val == 0)

        if pt == 'FP' and not is_tr:
            if has_csr and mrp == 'ND':
                pass
            else:
                if mrp != 'M1':
                    erros.append(f"For FP, MRP must be 'M1' (current: '{mrp}')"
                                 + (" — note: ND accepted if active CSR exists" if mrp == 'ND' else ""))
                if not has_tf:
                    erros.append(f"For FP, Time Fence must be > 0 (current: '{tf_raw}')")
        elif pt in {'FP', 'MP'} and is_tr:
            if mrp not in {'M1', 'P1', 'PD'}:
                erros.append(f"For {pt} with transfer, allowed MRP Types: M1, P1, PD (current: '{mrp}')")
        elif pt == 'MP' and not is_tr:
            if mrp != 'PD':
                erros.append(f"For MP without transfer, MRP must be 'PD' (current: '{mrp}')")
            if not no_tf:
                erros.append(f"For MP without transfer, Time Fence must be 0 or empty (current: '{tf_raw}')")
        elif pt == 'PP':
            if has_tf and mrp != 'P1':
                erros.append(f"For PP with Time Fence > 0, MRP must be 'P1' (current: '{mrp}')")
            elif no_tf and mrp != 'PD':
                erros.append(f"For PP with Time Fence = 0/empty, MRP must be 'PD' (current: '{mrp}')")

        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'MRP Type',
                               'Error': ' \n'.join(erros), 'Material': mat, 'Plant': plant})


def tn_validar_planning_calendar(df, resultados):
    """TN01 FP/MP -> 'TN1'; TN02 FP/MP -> 'TN2'. PP rules same as Palmela."""
    required = {'Material', 'Plant', 'Plant Type', 'Planning Calendar',
                 'Plant Specific Material Status'}
    if not required.issubset(df.columns):
        return

    def _st(v):
        if pd.isna(v) or str(v).strip() == '':
            return None
        try:
            return int(float(str(v).strip()))
        except Exception:
            return None

    for idx, row in df.iterrows():
        erros  = []
        mat    = _norm_str(row.get('Material'))
        plant  = _norm_str(row.get('Plant'))
        pt     = _norm_str(row.get('Plant Type')).upper()
        pc_raw = row.get('Planning Calendar')
        pc     = '' if pd.isna(pc_raw) else str(pc_raw).strip().upper()
        psms   = _st(row.get('Plant Specific Material Status'))
        has14k = '14K450' in mat.upper()

        if psms is None or not (10 <= psms <= 70):
            continue

        if has14k and pt == 'MP':
            if pc == '':
                erros.append("For Plant Type 'MP' with '14K450', Planning Calendar must not be empty")
            elif pc in ('PT2', 'TN2'):
                erros.append("For Plant Type 'MP' with '14K450', Planning Calendar must not be 'PT2'")
            if erros:
                resultados.append({'Line': idx + 2, 'Field': 'Planning Calendar',
                                   'Error': ' \n'.join(erros), 'Material': mat, 'Plant': plant})
            continue

        if plant == 'TN01':
            if pt in ('FP', 'MP') and pc != 'TN1':
                erros.append(f"For Plant TN01 and Plant Type '{pt}', Planning Calendar must be 'TN1' (current: '{pc_raw}')")
            elif pt == 'PP':
                if pc == '':
                    erros.append("For Plant TN01 and Plant Type 'PP', Planning Calendar must not be empty")
                elif pc == 'TN2':
                    erros.append("For Plant TN01 and Plant Type 'PP', Planning Calendar must not be 'TN2'")
        elif plant == 'TN02':
            if pt in ('FP', 'MP') and pc != 'TN2':
                erros.append(f"For Plant TN02 and Plant Type '{pt}', Planning Calendar must be 'TN2' (current: '{pc_raw}')")
            elif pt == 'PP' and pc == '':
                erros.append("For Plant TN02 and Plant Type 'PP', Planning Calendar must not be empty")

        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Planning Calendar',
                               'Error': ' \n'.join(erros), 'Material': mat, 'Plant': plant})


def tn_validar_planning_time_fence(df, mara_df, resultados):
    """TN02 FERT: must be 0 or empty. Others: Palmela rules (PD->0, P1/M1->0)."""
    required = {'Material', 'Plant', 'MRP Type', 'Planning Time Fence'}
    if not required.issubset(df.columns):
        return
    mt_map = _mt_map_from_mara(mara_df)

    def _tf(v):
        if pd.isna(v) or str(v).strip() == '':
            return None
        try:
            return float(str(v).strip())
        except Exception:
            return 'INVALID'

    def _st(v):
        if pd.isna(v) or str(v).strip() == '':
            return None
        try:
            return int(float(str(v).strip()))
        except Exception:
            return None

    for idx, row in df.iterrows():
        erros  = []
        mat    = _norm_str(row.get('Material'))
        plant  = _norm_str(row.get('Plant'))
        mrp    = _norm_str(row.get('MRP Type')).upper()
        mt     = mt_map.get(mat, '').upper()
        tf_raw = row.get('Planning Time Fence')
        tf     = _tf(tf_raw)
        psms   = _st(row.get('Plant Specific Material Status')) if 'Plant Specific Material Status' in df.columns else None

        if psms == 10:
            continue

        if plant == 'TN02' and mt == 'FERT':
            if tf not in (None, 0):
                erros.append(f"For TN02 FERT, Planning Time Fence should be empty or 0 (current: '{tf_raw}')")
        else:
            if tf == 'INVALID':
                erros.append("Planning Time Fence contains non-numeric value.")
            else:
                if mrp == 'PD' and tf not in (None, 0):
                    erros.append(f"For {mrp}, Planning Time Fence should be empty or 0 (current: '{tf_raw}')")
                elif mrp in ('P1', 'M1') and (tf is None or tf == 0):
                    erros.append(f"For {mrp}, Planning Time Fence should be > 0 (current: '{tf_raw}')")

        if erros:
            resultados.append({'Line': idx + 2, 'Field': 'Planning Time Fence',
                               'Error': ' \n'.join(erros), 'Material': mat, 'Plant': plant})


def tn_validar_production_scheduling_profile(df, mara_df, resultados):
    """TN02: must be BLANK. TN01: FERT/HALB must be 'ZREM01'."""
    col = 'Production Scheduling Profile'
    if col not in df.columns:
        return
    mt_map = _mt_map_from_mara(mara_df)
    for idx, row in df.iterrows():
        erros = []
        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        mt    = mt_map.get(mat, _norm_str(row.get('Material Type', ''))).upper()
        val   = _norm_str(row.get(col))
        if plant == 'TN02':
            if val != '':
                erros.append(f"Field must be blank or empty for TN02 (current: '{val}')")
        else:
            if mt in ('FERT', 'HALB'):
                if val == '':
                    erros.append("Mandatory field blank For Material Type FERT or HALB")
                elif val != 'ZREM01':
                    erros.append(f"Invalid value: '{val}' (should be 'ZREM01')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': ' | '.join(erros), 'Material': mat, 'Plant': plant})


def tn_validar_profit_center(df, resultados):
    """TN01/TN02: must be 'TN00413'."""
    col = 'Profit Center'
    if col not in df.columns:
        return
    for idx, row in df.iterrows():
        erros = []
        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        val   = _norm_str(row.get(col))
        if val == '':
            erros.append('Mandatory field blank')
        elif val != 'TN00413':
            erros.append(f"For Plant TN01/TN02, should be 'TN00413' (current: '{val}')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': ' | '.join(erros), 'Material': mat, 'Plant': plant})


def tn_validar_purchasing_group(df, resultados):
    """PP -> must be 'DXX'; FP -> must be 'STO'; MP -> blank."""

    col = 'Purchasing Group'
    if col not in df.columns or 'Plant Type' not in df.columns:
        return

    for idx, row in df.iterrows():
        erros = []

        plant = _norm_str(row.get('Plant')).upper()
        pt    = _norm_str(row.get('Plant Type')).upper()
        val   = _norm_str(row.get(col))
        mat   = _norm_str(row.get('Material'))

        # -----------------------------
        # PP → obrigatório e = DXX
        # -----------------------------
        if pt == 'PP':
            if val == '':
                erros.append("Mandatory field blank for Plant Type 'PP'")
            elif val.upper() != 'DXX':
                erros.append(f"For Plant Type 'PP', should be 'DXX' (current: '{val}')")

        # -----------------------------
        # FP → obrigatório e = STO
        # -----------------------------
        elif pt == 'FP':
            if val == '':
                erros.append("Mandatory field blank for Plant Type 'FP' in TN01/TN02")
            elif val.upper() != 'STO':
                erros.append(f"For Plant Type 'FP' in TN01/TN02, should be 'STO' (current: '{val}')")

        # -----------------------------
        # MP → deve ser vazio
        # -----------------------------
        elif pt == 'MP':
            if val != '':
                erros.append("Field should be blank for Plant Type 'MP'")

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' | '.join(erros),
                'Material': mat,
                'Plant': plant
            })

def tn_validar_repetitive_manufacturing(df, mara_df, resultados):
    """
    TN rule for Repetitive Manufacturing:
    - Only validates materials with Material Type in MARA = FERT/HALB/ROH
    - TN02 -> must be blank
    - TN01 / others:
        - PP    -> blank
        - MP/FP -> 'X'
    Uses the merged sap_like_df (df), which already contains Plant Type.
    """
    field_candidates = [
        'Repetitive Manufacturing',
        'Indicator: Repetitive Manufa. Allow'
    ]
    col = next((c for c in field_candidates if c in df.columns), None)
    if not col:
        return

    pt_candidates = ['Plant Type', 'Material Plant Type']
    pt_col = next((c for c in pt_candidates if c in df.columns), None)
    if not pt_col:
        return

    if 'Plant' not in df.columns or 'Material' not in df.columns:
        return

    if mara_df is None or mara_df.empty or 'Material' not in mara_df.columns or 'Material Type' not in mara_df.columns:
        return

    # Material -> Material Type (MARA)
    mt_map = (
        mara_df[['Material', 'Material Type']]
        .dropna(subset=['Material'])
        .assign(
            Material=lambda d: d['Material'].fillna('').astype(str).str.strip(),
            Material_Type=lambda d: d['Material Type'].fillna('').astype(str).str.strip().str.upper()
        )
        .set_index('Material')['Material_Type']
        .to_dict()
    )

    for idx, row in df.iterrows():
        erros = []

        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        pt    = _norm_str(row.get(pt_col)).upper()

        # only validate FERT / HALB / ROH
        mt = mt_map.get(mat, '')
        if mt not in {'FERT', 'HALB', 'ROH'}:
            continue

        val_raw = row.get(col)
        val = _norm_str(val_raw).upper()

        # Checkbox normalization
        if val in {'X', '1', 'YES'}:
            val_norm = 'X'
        else:
            val_norm = ''

        if plant == 'TN02':
            if val_norm != '':
                erros.append(f"Must be blank for TN02 (current: '{val}')")
        else:
            if pt == 'PP':
                if val_norm != '':
                    erros.append(
                        f"For Plant Type 'PP', field '{col}' must be blank (current: '{val}')"
                    )
            elif pt in {'MP', 'FP'}:
                if val_norm != 'X':
                    erros.append(
                        f"For Plant Type '{pt}', field '{col}' should be 'X' (current: '{val}')"
                    )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' | '.join(erros),
                'Material': mat,
                'Plant': plant
            })

def tn_validar_repetitive_manufacturing_profile(df, mara_df, resultados):
    """
    TN rule for REM Profile:
    - Only validates materials with Material Type in MARA = FERT/HALB/ROH
    - TN02 -> must be blank
    - TN01 / others:
        - PP    -> blank
        - MP/FP -> 'Z002'
    Uses the merged sap_like_df (df), which already contains Plant Type.
    """
    field_candidates = [
        'REM Profile',
        'Repetitive Manufacturing Profile'
    ]
    col = next((c for c in field_candidates if c in df.columns), None)
    if not col:
        return

    pt_candidates = ['Plant Type', 'Material Plant Type']
    pt_col = next((c for c in pt_candidates if c in df.columns), None)
    if not pt_col:
        return

    if 'Plant' not in df.columns or 'Material' not in df.columns:
        return

    if mara_df is None or mara_df.empty or 'Material' not in mara_df.columns or 'Material Type' not in mara_df.columns:
        return

    # Material -> Material Type (MARA)
    mt_map = (
        mara_df[['Material', 'Material Type']]
        .dropna(subset=['Material'])
        .assign(
            Material=lambda d: d['Material'].fillna('').astype(str).str.strip(),
            Material_Type=lambda d: d['Material Type'].fillna('').astype(str).str.strip().str.upper()
        )
        .set_index('Material')['Material_Type']
        .to_dict()
    )

    for idx, row in df.iterrows():
        erros = []

        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        pt    = _norm_str(row.get(pt_col)).upper()

        # only validate FERT / HALB / ROH
        mt = mt_map.get(mat, '')
        if mt not in {'FERT', 'HALB', 'ROH'}:
            continue

        val = _norm_str(row.get(col)).upper()

        if plant == 'TN02':
            if val != '':
                erros.append(f"Must be blank for TN02 (current: '{val}')")
        else:
            if pt == 'PP':
                if val != '':
                    erros.append(
                        f"For Plant Type 'PP', field '{col}' must be blank (current: '{val}')"
                    )
            elif pt in {'MP', 'FP'}:
                if val != 'Z002':
                    erros.append(
                        f"For Plant Type '{pt}', field '{col}' should be 'Z002' (current: '{val}')"
                    )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' | '.join(erros),
                'Material': mat,
                'Plant': plant
            })

def tn_validar_source_list_required(df, resultados):
    """PP or FP -> Yes (1); others -> No (0).

    Supports both TN source fields: 'Source list' and 'Source list required'."""
    col = next((c for c in ['Source list', 'Source list required'] if c in df.columns), None)
    if col is None or 'Plant Type' not in df.columns:
        return
    for idx, row in df.iterrows():
        erros = []
        plant = _norm_str(row.get('Plant')).upper()
        pt    = _norm_str(row.get('Plant Type')).upper()
        val   = _norm_str(row.get(col))
        mat   = _norm_str(row.get('Material'))

        if pt in ('PP', 'FP'):
            if val == '':
                erros.append('Mandatory field blank')
            elif val not in ('1', '1.0', 'X', 'YES'):
                erros.append(f"For Plant Type '{pt}', should be 'Yes' (current: '{val}')")
        else:
            if val not in ('0', ''):
                erros.append(f"For Plant Type '{pt}', should be 'No' (current: '{val}')")

        if erros:
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': ' | '.join(erros), 'Material': mat, 'Plant': plant})


def tn_validar_storage_location(df, mara_df, resultados):
    """TN01: HALB/ROH -> T105; FERT -> T120. TN02: HALB/FERT -> T220."""
    col = 'Issue Storage Location'
    if col not in df.columns:
        return
    mt_map = _mt_map_from_mara(mara_df)
    for idx, row in df.iterrows():
        erros = []
        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        mt    = mt_map.get(mat, _norm_str(row.get('Material Type', ''))).upper()
        val   = _norm_str(row.get(col))
        if val == '':
            erros.append('Mandatory field blank')
        elif plant == 'TN01':
            if mt in ('HALB', 'ROH') and val != 'T105':
                erros.append(f"For TN01 + {mt}, must be 'T105' (current: '{val}')")
            elif mt == 'FERT' and val != 'T120':
                erros.append(f"For TN01 + FERT, must be 'T120' (current: '{val}')")
        elif plant == 'TN02':
            if mt in ('HALB', 'FERT') and val != 'T220':
                erros.append(f"For TN02 + {mt}, must be 'T220' (current: '{val}')")
        if erros:
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': ' | '.join(erros), 'Material': mat, 'Plant': plant})


def tn_validar_special_procurement_type(marc_df, zplm_marc_df, resultados):
    """TN02 FP -> must be 'Z3'."""
    col = 'Special Procurement Type'
    if marc_df is None or col not in marc_df.columns:
        return
    pt_map = _pt_map_from_zplm(zplm_marc_df)
    for idx, row in marc_df.iterrows():
        plant = _norm_str(row.get('Plant')).upper()
        if plant != 'TN02':
            continue
        mat = _norm_str(row.get('Material'))
        pt  = pt_map.get((mat, plant), '')
        if pt != 'FP':
            continue
        val = _norm_str(row.get(col)).upper()
        if val != 'Z3':
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': f"For Plant Type 'FP' from TN02, should be 'Z3' (current: '{val if val else 'blank'}')",
                               'Material': mat, 'Plant': plant})


def tn_validar_account_assignment_group(mara_df, marc_df, mvke_df, resultados):
    """TN FERT: must have TN00/10 + TN00/20. AAG = Z1 for TN entries."""
    if mara_df is None or marc_df is None or mvke_df is None:
        return
    mt_map = _mt_map_from_mara(mara_df)
    tn_mats = set()
    if 'Plant' in marc_df.columns and 'Material' in marc_df.columns:
        tn_mats = set(
            marc_df[marc_df['Plant'].astype(str).str.strip().str.upper().isin(TN_PLANTS)]
            ['Material'].astype(str).str.strip()
        )
    scope = {m for m in tn_mats if mt_map.get(m, '') in {'FERT', 'HALB', 'ROH'}}
    if not scope:
        return
    mvke_norm, _ = _prepare_mvke(mvke_df, scope)
    if mvke_norm is None or mvke_norm.empty:
        return
    TN_REQ = {('TN00', '10'), ('TN00', '20')}

    def _err(material, msg):
        resultados.append({'Line': '', 'Field': 'Account Assignment Group',
                           'Error': msg, 'Material': material, 'Plant': ''})

    mats_in_mvke = set(mvke_norm['_mat'].unique())
    for material in scope:
        mt      = mt_map.get(material, '')
        tn_rows = mvke_norm[(mvke_norm['_mat'] == material) & (mvke_norm['_sorg'] == 'TN00')]
        tn_combos = set(zip(tn_rows['_sorg'], tn_rows['_dchan']))
        if mt == 'FERT':
            if material not in mats_in_mvke or tn_rows.empty:
                _err(material, "TN FERT: no TN00 entries found — required combinations TN00/10, TN00/20 are missing")
                continue
            for sorg, dchan in sorted(TN_REQ - tn_combos):
                _err(material, f"TN FERT: missing required MVKE combination SalesOrg='{sorg}' / DC='{dchan}'")
        for _, row in tn_rows.iterrows():
            combo = f"SalesOrg='{row['_sorg']}' / DC='{row['_dchan']}'"
            aag   = row['_aag']
            if aag == '':
                _err(row['_mat'], f"Mandatory field blank in MVKE ({combo})")
            elif aag.upper() != 'Z1':
                _err(row['_mat'], f"Must be 'Z1' for {combo} (current: '{aag}')")


def tn_validar_storage_location_ext_procurement(marc_df, zplm_marc_df, mara_df, resultados):
    """
    Storage Location for Ext. Procurement:

    Base:
    - PP → mandatory
    - TN02 + FP → mandatory

    TN-specific values:
    TN01:
        ROH  → T100
        HALB → T115
        FERT → T120

    TN02:
        ROH  → must be empty
        HALB → T220
        FERT → T220
    """

    col = 'Storage Location for Ext. Procurement'

    if marc_df is None or col not in marc_df.columns:
        return

    pt_map = _pt_map_from_zplm(zplm_marc_df)
    mt_map = _mt_map_from_mara(mara_df)

    # Defensive: iterate over deduplicated (Material,Plant) pairs from MARC
    marc = marc_df.copy()
    marc.columns = [str(c).strip() for c in marc.columns]
    if 'Material' in marc.columns and 'Plant' in marc.columns:
        marc['Material'] = marc['Material'].astype(str).str.strip()
        marc['Plant'] = marc['Plant'].astype(str).str.strip().str.upper()
        marc = marc.drop_duplicates(subset=['Material', 'Plant'], keep='first')
    else:
        return

    for idx, row in marc.iterrows():
        erros = []

        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        pt    = pt_map.get((mat, plant), '')
        mt    = mt_map.get(mat, '')
        val   = _norm_str(row.get(col))

        # --------------------------------------------------------
        # ✅ BASE RULES
        # --------------------------------------------------------
        if pt == 'PP' and val == '':
            erros.append("Mandatory field blank for Plant Type PP")

        elif plant == 'TN02' and pt == 'FP' and val == '':
            erros.append("Mandatory field blank for Plant Type FP in TN02")

        # --------------------------------------------------------
        # ✅ TN01 VALUE RULES
        # --------------------------------------------------------
        if plant == 'TN01' and mt in {'FERT', 'HALB', 'ROH'}:
            expected_map = {
                'ROH':  'T100',
                'HALB': 'T115',
                'FERT': 'T120',
            }

            expected = expected_map.get(mt)

            if val == '':
                erros.append(f"Mandatory field blank (should be '{expected}' for TN01 {mt})")
            elif val != expected:
                erros.append(f"Invalid value: '{val}' (should be '{expected}' for TN01 {mt})")

        # --------------------------------------------------------
        # ✅ TN02 VALUE RULES
        # --------------------------------------------------------
        if plant == 'TN02' and mt in {'FERT', 'HALB', 'ROH'}:

            if mt == 'ROH':
                if val != '':
                    erros.append("Must be empty for TN02 ROH")

            else:  # HALB / FERT
                if val == '':
                    erros.append("Mandatory field blank (should be 'T220')")
                elif val != 'T220':
                    erros.append(f"Invalid value: '{val}' (should be 'T220' for TN02 {mt})")

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' | '.join(erros),
                'Material': mat,
                'Plant': plant
            })

def tn_validar_individual_collective_requirements(marc_df, mara_df, resultados):
    """FERT -> '1'. HALB/ROH -> '2'."""
    col = 'Individual/ Collective Requirements'
    if marc_df is None or col not in marc_df.columns:
        return
    mt_map = _mt_map_from_mara(mara_df)
    for idx, row in marc_df.iterrows():
        plant    = _norm_str(row.get('Plant')).upper()
        mat      = _norm_str(row.get('Material'))
        mt       = mt_map.get(mat, '')
        val      = normalize_code(row.get(col), as_int=True)
        if mt not in {'FERT', 'HALB', 'ROH'}:
            continue
        expected = '1' if mt == 'FERT' else '2'
        label    = f"'{expected}'"
        if val == '':
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': f"Mandatory field blank (should be {label})",
                               'Material': mat, 'Plant': plant})
        elif val != expected:
            resultados.append({'Line': idx + 2, 'Field': col,
                               'Error': f"Invalid value: '{val}' (should be {label})",
                               'Material': mat, 'Plant': plant})


def tn_validar_production_supervisor(marc_df, mara_df, resultados):
    """TN02: must be BLANK. TN01: FERT/HALB -> 'Z01'."""
    col = 'Production Supervisor'
    if marc_df is None or col not in marc_df.columns:
        return
    mt_map = _mt_map_from_mara(mara_df)
    for idx, row in marc_df.iterrows():
        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        mt    = mt_map.get(mat, '')
        val   = _norm_str(row.get(col)).upper()
        if plant == 'TN02':
            if val != '':
                resultados.append({'Line': idx + 2, 'Field': col,
                                   'Error': f"Field must be blank or empty for TN02 (current: '{val}')",
                                   'Material': mat, 'Plant': plant})
        else:
            if mt in {'FERT', 'HALB'}:
                if val == '':
                    resultados.append({'Line': idx + 2, 'Field': col,
                                       'Error': "Mandatory field blank (should be 'Z01')",
                                       'Material': mat, 'Plant': plant})
                elif val != 'Z01':
                    resultados.append({'Line': idx + 2, 'Field': col,
                                       'Error': f"Invalid value: '{val}' (should be 'Z01')",
                                       'Material': mat, 'Plant': plant})


def tn_validar_advanced_planning_fields(marc_df, mara_df, resultados):
    """TN01 FERT -> X; TN02 FERT/HALB -> X. Heuristic/PP Proc/Req Strategy same as Palmela."""
    if marc_df is None or mara_df is None:
        return

    def _pick(df, candidates):
        return next((c for c in candidates if c in df.columns), None)

    col_adv  = _pick(marc_df, ['Advanced Planning'])
    col_heur = _pick(marc_df, ['Heuristic'])
    col_pp   = _pick(marc_df, ['PP Plng Procedure'])
    col_req  = _pick(marc_df, ['Requirement Strategy'])

    if any(c is None for c in [col_adv, col_heur, col_pp, col_req]):
        return

    mt_map = _mt_map_from_mara(mara_df)

    def _s(x):
        return '' if pd.isna(x) else str(x).strip()

    for idx, row in marc_df.iterrows():
        mat   = _s(row.get('Material'))
        plant = _s(row.get('Plant'))
        mt    = get_effective_material_type(mat, mt_map.get(mat, '').upper(),
                                            'validar_advanced_planning_fields')
        if mt not in {'FERT', 'HALB'}:
            continue

        adv  = _s(row.get(col_adv)).upper()
        heur = _s(row.get(col_heur))
        pp   = normalize_code(row.get(col_pp), as_int=True)
        req  = normalize_code(row.get(col_req), as_int=True)

        if (plant == 'TN01' and mt == 'FERT') or \
           (plant == 'TN02' and mt in {'FERT', 'HALB'}):
            if adv != 'true':
                resultados.append({'Line': idx + 2, 'Field': 'Advanced Planning',
                                   'Error': "Mandatory: must be 'X' for TN01 FERT or TN02 FERT/HALB",
                                   'Material': mat, 'Plant': plant})

        if heur == '':
            resultados.append({'Line': idx + 2, 'Field': 'Heuristic',
                               'Error': "Mandatory: cannot be empty for FERT or HALB",
                               'Material': mat, 'Plant': plant})
        if pp != '4':
            resultados.append({'Line': idx + 2, 'Field': 'PP Plng Procedure',
                               'Error': "Invalid value: should be '4'",
                               'Material': mat, 'Plant': plant})
        if req != '20':
            resultados.append({'Line': idx + 2, 'Field': 'Proposed Strategy',
                               'Error': "Invalid value: should be '20'",
                               'Material': mat, 'Plant': plant})


# =============================================================================
# MAIN ORCHESTRATOR
# =============================================================================

def run_tn_validation(
    mara_file,
    marc_file,
    mvke_file,
    marm_file=None,
    mbew_file=None,
    zplm_marc_file=None,
    zplm_mara_file=None,
    previous_report=None,
    plant_filter=None,
    progress_callback=None,
    csr_item_file=None,
    csr_head_file=None,
    matplsta_file=None,
    # Optional pre-loaded DataFrames — if provided, skips file loading for those
    _mara_df=None,
    _marc_df=None,
    _mvke_df=None,
    _marm_df=None,
    _mbew_df=None,
    _zplm_marc_df=None,
    _zplm_mara_df=None,
    _csr_item_df=None,
    _csr_head_df=None,
    _matplsta_df=None,
):
    """
    Runs the complete Tunisia validation pipeline.
    Returns: (excel_bytes, errors_df, summary_df)
    """
    def _cb(pct, msg):
        if progress_callback:
            progress_callback(min(pct, 100), msg)

    # ── 1. Load files (skip if pre-loaded DataFrames provided) ───────────────
    _cb(3,  "Loading MARA...")
    mara_df = _mara_df if _mara_df is not None else load_xlsx_fast(mara_file, sheet_name=0)
    _cb(6,  "Loading MARC...")
    marc_df = _marc_df if _marc_df is not None else load_xlsx_fast(marc_file, sheet_name=0)
    _cb(9,  "Loading MVKE...")
    mvke_df = _mvke_df if _mvke_df is not None else load_xlsx_fast(mvke_file, sheet_name=0)
    _cb(11, "Loading ZPLM MARC...")
    zplm_marc_df = _zplm_marc_df if _zplm_marc_df is not None else (load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None)
    _cb(13, "Loading ZPLM MARA...")
    zplm_mara_df = _zplm_mara_df if _zplm_mara_df is not None else (load_xlsx_fast(zplm_mara_file, sheet_name=0) if zplm_mara_file else None)
    _cb(15, "Loading optional files...")
    csr_item_df = _csr_item_df if _csr_item_df is not None else (load_xlsx_fast(csr_item_file, sheet_name=0) if csr_item_file else None)
    csr_head_df = _csr_head_df if _csr_head_df is not None else (load_xlsx_fast(csr_head_file, sheet_name=0) if csr_head_file else None)
    matplsta_df = _matplsta_df if _matplsta_df is not None else (load_xlsx_fast(matplsta_file, sheet_name=0) if matplsta_file else None)
    marm_df     = _marm_df if _marm_df is not None else (load_xlsx_fast(marm_file, sheet_name=0) if marm_file else None)
    mbew_df     = _mbew_df if _mbew_df is not None else (load_xlsx_fast(mbew_file, sheet_name=0) if mbew_file else None)

    if mara_df is None:
        raise ValueError("MARA.xlsx not loaded.")
    if marc_df is None:
        raise ValueError("MARC.xlsx not loaded.")
    if mvke_df is None:
        raise ValueError("MVKE.xlsx not loaded.")

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
    mvke_df = _normalize_df_columns(mvke_df, {'plant': 'Plant', 'werks': 'Plant', 'material': 'Material'})
    marm_df = _normalize_df_columns(marm_df, {'plant': 'Plant', 'werks': 'Plant', 'material': 'Material'})

    rolesviews_dict = load_xlsx_fast(get_rolesviews_bytesio(), sheet_name=None)
    if "Roles" not in rolesviews_dict or "Views" not in rolesviews_dict:
        raise ValueError("RolesViews.xlsx must contain sheets 'Roles' and 'Views'.")
    roles_df = rolesviews_dict["Roles"]
    views_df = rolesviews_dict["Views"]

    # ── 2. Filter to TN plants ────────────────────────────────────────────────
    tn_scope = set(TN_PLANTS)
    if plant_filter:
        pf = str(plant_filter).strip().upper()
        if pf in TN_PLANTS:
            tn_scope = {pf}

    def _ftn(df, col='Plant'):
        if df is None or df.empty or col not in df.columns:
            return df
        return df[df[col].astype(str).str.strip().str.upper().isin(tn_scope)].copy()

    marc_tn      = _ftn(marc_df)
    zplm_marc_tn = _ftn(zplm_marc_df) if zplm_marc_df is not None else None
    marc_tn      = ensure_basic_data_k(marc_tn)

    tn_materials = set(
        marc_tn['Material'].astype(str).str.strip()
        .loc[marc_tn['Material'].notna()]
    )

    def _filter_to_tn_materials(df):
        if df is None or df.empty or 'Material' not in df.columns:
            return df

        out = df[df['Material'].astype(str).str.strip().isin(tn_materials)].copy()

        plant_col = next(
            (c for c in ['Plant', 'Valuation Area', 'BWKEY', 'WERKS']
             if c in out.columns),
            None
        )
        if plant_col:
            out = out[
                out[plant_col].astype(str).str.strip().str.upper().isin(tn_scope)
            ].copy()

        return out

    # mara_df is material-level (no Plant column) — do NOT filter by tn_materials.
    # Filtering it would exclude materials that exist in MARA but have no TN marc
    # entry yet, causing MARA-based validations to silently miss them.
    # Plant-level tables are filtered normally.
    #
    # marm_df (Alternative UoM) is ALSO material-level, not plant-level, and its
    # only consumer is validar_marm_pc_ea_roh_halb, which is called with the full
    # (unfiltered) mara_df. Filtering marm_df down to TN-only materials here would
    # create a mismatch: materials that legitimately exist only in Palmela plants
    # would still appear in mara_relevant (built from the global mara_df) but be
    # missing from marm_groups (built from the TN-restricted marm_df), producing
    # false "Alternative UoM found: none" errors for every non-TN material when
    # running the combined/general report. So marm_df is intentionally left
    # unfiltered here, same as mara_df.
    mbew_df      = _filter_to_tn_materials(mbew_df)
    mvke_df      = _filter_to_tn_materials(mvke_df)
    # zplm_mara_df: kept UNFILTERED for the same reason as marm_df above.
    # validar_product_hierarchy_vs_enterproj() iterates the full (unfiltered)
    # mara_df looking up each FERT material's Enterproj Number in zplm_mara_df.
    # If zplm_mara_df were restricted to TN materials, every Palmela-only FERT
    # material would be missing from that lookup, be treated as Enterproj="000...",
    # and be falsely flagged as a "PH mismatch" when running the combined report.

    # ── 3. Build sap_like_df ──────────────────────────────────────────────────
    _cb(18, "Building SAP-like dataframe from MARC + MARA + ZPLM MARC...")
    sap_like_df = build_sap_like_df(marc_tn, mara_df, zplm_marc_tn)

    # ✅ GLOBAL FILTER — ONLY FERT / HALB / ROH
    sap_like_df = _filter_valid_material_types(sap_like_df, mara_df)

    # aplicar também às tabelas usadas em validações diretas
    marc_tn = _filter_valid_material_types(marc_tn, mara_df)
    marc_tn_full = marc_tn.copy()

    exclude_mps99_pairs = _get_mps99_material_plant_pairs(marc_tn_full)
    if exclude_mps99_pairs:
        sap_like_df = _exclude_material_plant_status_99(sap_like_df, exclude_mps99_pairs)
        marc_tn = _exclude_material_plant_status_99(marc_tn, exclude_mps99_pairs)
        if zplm_marc_tn is not None:
            zplm_marc_tn = _exclude_material_plant_status_99(
                zplm_marc_tn,
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

    resultados = []

    # ── 4A. Mandatory fields ──────────────────────────────────────────────────
    _cb(20, "Validating mandatory fields...")
    for col in ['Material', 'Material Type', 'Description', 'Base UOM',
                'Plant', 'Plant Type']:
        validar_campo_obrigatorio(sap_like_df, col, resultados)

    # ── 4B. Fields WITHOUT TN exceptions — Palmela functions directly ─────────
    _cb(22, "Validating Division, Export/Import Group, Gen. Item Category...")
    validar_division(sap_like_df, resultados)
    validar_export_import_group(sap_like_df, resultados)
    validar_general_item_category_group(sap_like_df, resultados)
    validar_batch_management(sap_like_df, resultados)  # uses Batch Management Required Ind. from MARA

    _cb(25, "Validating Gross/Net Weight...")
    validar_gross_weight(sap_like_df, resultados)
    validar_net_weight(sap_like_df, resultados)

    _cb(26, "Validating Loading Group, Lot Sizing Procedure...")
    validar_loading_group(sap_like_df, resultados)
    validar_lot_sizing_procedure(sap_like_df, resultados)

    _cb(27, "Validating Material Group...")
    validar_material_group(sap_like_df, resultados)
    validar_material_group_appoint(sap_like_df, resultados)

    _cb(28, "Validating Period Type...")
    validar_period_type(sap_like_df, resultados)

    _cb(29, "Validating Price Control, Price Unit...")
    validar_price_control(sap_like_df, resultados)
    validar_price_unit(sap_like_df, resultados)

    _cb(30, "Validating Product Hierarchy...")
    validar_product_hierarchy(sap_like_df, resultados, mvke_df, mara_df)

    _cb(31, "Validating Valuation Class, With Qty Structure...")
    validar_valuation_class(sap_like_df, resultados)
    validar_with_qty_structure(sap_like_df, resultados)

    _cb(32, "Validating Product-Related Origin, MRP Controller, PSMS...")
    validar_product_related_origin(sap_like_df, resultados)
    validar_mrp_controller(sap_like_df, resultados)
    validar_plant_specific_material_status(sap_like_df, resultados)

    _cb(33, "Validating Enterproj Number, Purchasing Group DXX...")
    validar_enterproj_number(sap_like_df, resultados)
    validar_purchasing_group_dxx(sap_like_df, resultados)

    _cb(34, "Validating Item Category Group (MVKE)...")
    validar_item_category_group(sap_like_df, resultados, mvke_df)

    # ── 4C. Fields WITH TN exceptions — TN functions ─────────────────────────
    _cb(40, "Validating Backflush (TN)...")
    tn_validar_backflush(sap_like_df, resultados)

    _cb(41, "Validating Strategy Group (TN)...")
    tn_validar_strategy_group(sap_like_df, mara_df, resultados)

    _cb(42, "Validating Availability Check (TN)...")
    tn_validar_availability_check(sap_like_df, resultados)

    _cb(43, "Validating Procurement Type (TN)...")
    tn_validar_procurement_type(sap_like_df, mara_df, resultados)

    _cb(44, "Validating Automatic Purchase Order (TN)...")
    tn_validar_automatic_purchase_order(sap_like_df, resultados)

    _cb(45, "Validating Consumption Mode (TN)...")
    tn_validar_consumption_mode(sap_like_df, mara_df, resultados)

    _cb(46, "Validating JIT Delivery Sched. (TN)...")
    tn_validar_jit_delivery_sched(sap_like_df, resultados)

    _cb(47, "Validating MRP Type (TN)...")
    tn_validar_mrp_type(sap_like_df, mara_df, resultados, zplm_marc_tn,
                         csr_item_df, csr_head_df)

    _cb(48, "Validating Planning Calendar (TN)...")
    tn_validar_planning_calendar(sap_like_df, resultados)

    _cb(49, "Validating Planning Time Fence (TN)...")
    tn_validar_planning_time_fence(sap_like_df, mara_df, resultados)

    _cb(50, "Validating Production Scheduling Profile (TN)...")
    tn_validar_production_scheduling_profile(sap_like_df, mara_df, resultados)

    _cb(51, "Validating Profit Center (TN)...")
    tn_validar_profit_center(sap_like_df, resultados)

    _cb(52, "Validating Purchasing Group (TN)...")
    tn_validar_purchasing_group(sap_like_df, resultados)

    _cb(53, "Validating Repetitive Manufacturing (TN)...")
    tn_validar_repetitive_manufacturing(sap_like_df, mara_df, resultados)

    _cb(54, "Validating Repetitive Manufacturing Profile (TN)...")
    tn_validar_repetitive_manufacturing_profile(sap_like_df, mara_df, resultados)

    _cb(55, "Validating Source List Required (TN)...")
    tn_validar_source_list_required(sap_like_df, resultados)

    _cb(56, "Validating Issue Storage Location (TN)...")
    tn_validar_storage_location(sap_like_df, mara_df, resultados)

    # ── 4D. MARC-based cross-file validations (all from Palmela) ─────────────
    _cb(57, "Validating Bwd/Fwd Consumption Per. (MARC)...")
    validar_consumption_per(marc_tn, mara_df, resultados)

    _cb(58, "Validating Per.Lot Sz.Plng Cal. (MARC)...")
    validar_per_lot_sz_plng_cal(marc_tn, resultados, mara_df, matplsta_df)

    _cb(59, "Validating Stock Determination Group (MARC)...")
    validar_stock_determination_group_marc(marc_tn, mara_df, resultados)

    _cb(60, "Validating Packaging Mat. Type (MARA)...")
    validar_packaging_material_type_mara(mara_df, resultados)

    # Batch Management per-plant check removed to avoid duplicate counting
    # (Batch Management Required Ind. is validated earlier via validar_batch_management)

    _cb(62, "Validating Version Indicator (MARC)...")
    if zplm_marc_tn is not None:
        validar_version_indicator_marc(marc_tn, zplm_marc_tn, mara_df, resultados)

    _cb(63, "Validating Country/Region of Origin (MARC)...")
    if zplm_marc_tn is not None:
        validar_country_region_origin_fp(marc_tn, zplm_marc_tn, mara_df, resultados)

    _cb(64, "Validating Product Hierarchy vs Enterproj (MARA)...")
    if zplm_mara_df is not None:
        validar_product_hierarchy_vs_enterproj(mara_df, zplm_mara_df, resultados)

    _cb(65, "Validating Alternative UoM PC (MARM)...")
    if marm_df is not None:
        validar_marm_pc_ea_roh_halb(marm_df=marm_df, mara_df=mara_df, resultados=resultados)

    _cb(66, "Validating LIFO/FIFO-relevant (MBEW)...")
    if mbew_df is not None:
        validar_lifo_fifo_relevant_mbew(mbew_df, mara_df, resultados)

    # ── 4E. TN-specific MARC cross-file rules ─────────────────────────────────
    _cb(67, "Validating Special Procurement Type (TN02 FP)...")
    tn_validar_special_procurement_type(marc_tn, zplm_marc_tn, resultados)

    _cb(68, "Validating Storage Location Ext. Procurement (TN)...")
    tn_validar_storage_location_ext_procurement(
        marc_tn, zplm_marc_tn, mara_df, resultados
    )

    _cb(69, "Validating Individual/Collective Requirements (TN)...")
    tn_validar_individual_collective_requirements(marc_tn, mara_df, resultados)

    _cb(70, "Validating Production Supervisor (TN)...")
    tn_validar_production_supervisor(marc_tn, mara_df, resultados)

    _cb(71, "Validating Account Assignment Group (TN MVKE)...")
    tn_validar_account_assignment_group(mara_df, marc_tn, mvke_df, resultados)



    _cb(73, "Validating Costing Lot Size (MARC)...")
    validar_costing_lot_size(marc_tn, mara_df, resultados)

    _cb(74, "Validating Safety Time (MARC)...")
    validar_safety_time_marc(marc_tn, mara_df, resultados)

    _cb(74, "Validating Material Price Determination Control (MBEW)...")
    if mbew_df is not None:
        validar_material_price_determination_control(mbew_df, mara_df, resultados)

    _cb(75, "Validating Advanced Planning fields (TN)...")
    tn_validar_advanced_planning_fields(marc_tn_full, mara_df, resultados)

    _cb(76, "Validating Neg. stocks in plant (MARC)...")
    neg_errors = validar_neg_stocks_in_plant(marc_tn, mara_df)
    if neg_errors:
        resultados.extend(neg_errors)

    # ── 5. Build errors DataFrame ─────────────────────────────────────────────
    _cb(78, "Building errors DataFrame...")
    errors_df = _coerce_df_for_display(pd.DataFrame(resultados))

    # ✅ Safety-net dedup: remove exact duplicates (same Material, Plant,
    # Field and Error). This is acceptable/expected to differ across
    # different plants for the same material (e.g. PT01 vs TN01), but the
    # exact same (Material, Plant, Field, Error) combination should never
    # be reported more than once — it would otherwise inflate error counts
    # if any upstream merge or source table happens to contain duplicate
    # rows for the same key.
    _dedup_key = [c for c in ['Material', 'Plant', 'Field', 'Error'] if c in errors_df.columns]
    if len(_dedup_key) == 4 and not errors_df.empty:
        for _c in _dedup_key:
            errors_df[_c] = errors_df[_c].fillna('').astype(str).str.strip()
        errors_df = errors_df.drop_duplicates(subset=_dedup_key).reset_index(drop=True)

    
    # ------------------------------------------------------------
    # ✅ DETETAR ERROS NÃO MAPEADOS NO ROLES
    # ------------------------------------------------------------
    roles_bytes = get_rolesviews_bytesio()
    roles_dict = load_xlsx_fast(roles_bytes, sheet_name=None)
    roles_df = roles_dict.get("Roles", None)

    if roles_df is not None and not roles_df.empty:
        errors_df['Field_clean'] = errors_df['Field'].astype(str).str.strip()
        roles_set = set(roles_df['Field'].astype(str).str.strip())

        missing_role_df = errors_df[~errors_df['Field_clean'].isin(roles_set)].copy()

        # DEBUG terminal
        if not missing_role_df.empty:
            print("\n--- ⚠️ ERROS NÃO MAPEADOS NO ROLES ---\n")
            for f in sorted(missing_role_df['Field_clean'].unique()):
                print(f)
            print("\n--------------------------------------\n")

        errors_df = errors_df.drop(columns=['Field_clean'], errors='ignore')
    else:
        missing_role_df = pd.DataFrame()

    if previous_report and zplm_marc_file and zplm_mara_file:
        errors_df = fill_plant_and_validate_enterproj_from_zplm(
            errors_df, latest_report=previous_report,
            zplm_marc_file=zplm_marc_file, zplm_mara_file=zplm_mara_file
        )

    if not errors_df.empty:
        errors_df = add_old_material_number_errors(mara_df, errors_df)

    # ── 6. Post-processing ────────────────────────────────────────────────────
    _cb(82, "Post-processing (Views / Roles / Classification)...")
    pf_post = plant_filter if (plant_filter and str(plant_filter).upper() in TN_PLANTS) else None

    (
        errors_df, errors_filtered_df, views_missing_output_df,
        contagem_final_df, coverage_df, view_tabs_dict
    ) = postprocess_with_views_roles_and_classification(
        errors_df, mara_df, marc_tn, roles_df, views_df, plant_filter=pf_post
    )

    # ── 7. Write Excel ────────────────────────────────────────────────────────
    _cb(90, "Generating Excel report...")
    file_bytes = io.BytesIO()

    if not contagem_final_df.empty and 'Adjusted Nº of Errors' in contagem_final_df.columns:
        contagem_final_df = contagem_final_df.sort_values(
            'Adjusted Nº of Errors', ascending=False).reset_index(drop=True)

    def _cat_missing(vtabs):
        cols = ['Material', 'Plant', 'Missing_View', 'View_Description', 'Fields_Expected']
        frames = [dfv[cols] for dfv in vtabs.values() if dfv is not None and not dfv.empty]
        if not frames:
            return pd.DataFrame(columns=cols)
        out = pd.concat(frames, ignore_index=True).drop_duplicates()
        out.sort_values(['Material', 'Plant', 'Missing_View'], inplace=True)
        return out

    view_is_missing_df = _cat_missing(view_tabs_dict)

    with pd.ExcelWriter(file_bytes, engine='openpyxl') as writer:

        pd.DataFrame().to_excel(writer, sheet_name='Executive Summary', index=False)

        base_row = 1
        if not contagem_final_df.empty:
            contagem_final_df.to_excel(
                writer,
                sheet_name='Adjusted Summary',
                index=False,
                startrow=1
            )
            base_row = len(contagem_final_df) + 5

        if not coverage_df.empty:
            coverage_df.to_excel(
                writer,
                sheet_name='Adjusted Summary',
                index=False,
                startrow=base_row
            )

        views_missing_output_df.to_excel(writer, sheet_name='Field With Error', index=False)
        errors_df.to_excel(writer, sheet_name='All Errors', index=False)
        view_is_missing_df.to_excel(writer, sheet_name='View is Missing', index=False)

        # ------------------------------------------------------------
        # ✅ NOVA SHEET → ERROS NÃO MAPEADOS
        # ------------------------------------------------------------
        if 'missing_role_df' in locals() and not missing_role_df.empty:
            missing_role_df[['Field', 'Material', 'Plant', 'Error']].to_excel(
                writer,
                sheet_name='Missing Role Mapping',
                index=False
            )

        # ------------------------------------------------------------
        # ✅ VIEWS MISSING (já existente)
        # ------------------------------------------------------------
        for view_code, df_view in view_tabs_dict.items():
            df_view.to_excel(writer, sheet_name=f'Missing_{view_code}'[:31], index=False)

        # ------------------------------------------------------------
        # ✅ Assignments sheet — enriched with owner/status from master.json
        # ------------------------------------------------------------
        try:
            _enriched = merge_assignments(errors_df)
            _asgn_df  = build_assignments_sheet(_enriched)
            _asgn_df.to_excel(writer, sheet_name='Assignments', index=False)
            save_open_assignments(_enriched)
        except Exception:
            pass

        marc_tn.to_excel(writer, sheet_name='Original Data (MARC)', index=False)

    file_bytes.seek(0)
    file_bytes = _style_sap_report(
        file_bytes,
        errors_df=errors_df,
        views_missing_output_df=views_missing_output_df,
        contagem_final_df=contagem_final_df,
        coverage_df=coverage_df,
        view_tabs_dict=view_tabs_dict,
        sap_df=marc_tn,
        plant_filter=plant_filter,
    )

    _cb(100, "Done!")

    summary_df = (
        errors_df['Field'].value_counts().rename_axis('Field').reset_index(name='Adjusted Nº of Errors')
        if 'Field' in errors_df.columns and not errors_df.empty
        else pd.DataFrame({'Field': [], 'Adjusted Nº of Errors': []})
    )

    return file_bytes, errors_df, summary_df