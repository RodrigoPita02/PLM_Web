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
#   3.15 MRP Group          — TN02 FERT: must be 'ZSTO'
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
#   3.32 Proposed Strategy  — TN01/TN02 FERT/HALB: '20' or 'Z1' (Palmela: must be '20')
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
    validar_product_hierarchy_mvke,
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
    """
    Procurement Type validation for Tunisia.

    Rules and precedence:

    1. Any material in TN02:
         Procurement Type must be 'F'.

    2. Outside TN02:
         - Plant Type PP:
             Procurement Type must be 'F' or 'X'.
         - Plant Type FP or MP:
             Procurement Type must be 'E' or 'X'.
         - When Plant Type is empty:
             - FERT or HALB: Procurement Type must be 'E' or 'X'.
             - ROH: Procurement Type must be 'F' or 'X'.

    Notes:
      - The TN02 rule has the highest precedence.
      - Therefore, TN02 materials must have 'F', regardless of
        Material Type or Plant Type.
      - The field is mandatory.
    """

    if df is None or df.empty:
        return

    col = 'Procurement Type'

    required_columns = {
        'Material',
        'Plant',
        'Plant Type',
        col,
    }

    if not required_columns.issubset(df.columns):
        return

    mt_map = _mt_map_from_mara(mara_df)

    universal_values = {
        'E',
        'F',
        'X',
    }

    for idx, row in df.iterrows():
        erros = []

        plant = _norm_str(
            row.get('Plant')
        ).upper()

        material = _norm_str(
            row.get('Material')
        )

        material_type = mt_map.get(
            material,
            _norm_str(
                row.get('Material Type', '')
            ).upper()
        )

        raw_value = row.get(col)

        procurement_type = (
            ''
            if pd.isna(raw_value)
            else str(raw_value).strip().upper()
        )

        plant_type = _norm_str(
            row.get('Plant Type')
        ).upper()

        # ========================================================
        # Mandatory and universally accepted values
        # ========================================================
        if procurement_type == '':
            if plant == 'TN02':
                erros.append(
                    "Mandatory field blank "
                    "(all TN02 materials require "
                    "Procurement Type = 'F')"
                )
            else:
                erros.append(
                    "Mandatory field blank"
                )

        elif procurement_type not in universal_values:
            erros.append(
                f"Invalid value: '{raw_value}' "
                "(allowed values: E, F, X)"
            )

        else:
            # ====================================================
            # RULE 1
            # Any material in TN02 must have F.
            #
            # This rule takes precedence over:
            #   - Material Type rules
            #   - Plant Type rules
            # ====================================================
            if plant == 'TN02':
                if procurement_type != 'F':
                    erros.append(
                        "All materials in TN02 require "
                        "Procurement Type = 'F' "
                        f"(current: '{procurement_type}')"
                    )

            # ====================================================
            # RULE 2
            # Outside TN02, Plant Type PP -> F or X
            # ====================================================
            elif plant_type == 'PP':
                if procurement_type not in {'F', 'X'}:
                    erros.append(
                        "For Plant Type 'PP', Procurement Type "
                        "must be 'F' or 'X' "
                        f"(current: '{procurement_type}')"
                    )

            # ====================================================
            # RULE 3
            # Outside TN02, Plant Type FP or MP -> E or X
            # ====================================================
            elif plant_type in {'FP', 'MP'}:
                if procurement_type not in {'E', 'X'}:
                    erros.append(
                        f"For Plant Type '{plant_type}', "
                        "Procurement Type must be 'E' or 'X' "
                        f"(current: '{procurement_type}')"
                    )

            # ====================================================
            # RULE 4
            # Fallback by Material Type when Plant Type is empty
            # ====================================================
            elif plant_type == '':
                if (
                    material_type in {'FERT', 'HALB'}
                    and procurement_type not in {'E', 'X'}
                ):
                    erros.append(
                        f"For Material Type '{material_type}', "
                        "Procurement Type must be 'E' or 'X' "
                        f"(current: '{procurement_type}')"
                    )

                elif (
                    material_type == 'ROH'
                    and procurement_type not in {'F', 'X'}
                ):
                    erros.append(
                        "For Material Type 'ROH', "
                        "Procurement Type must be 'F' or 'X' "
                        f"(current: '{procurement_type}')"
                    )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' \n'.join(erros),
                'Material': material,
                'Plant': plant,
            })


def tn_validar_automatic_purchase_order(df, resultados):
    """
    Validate the field 'Automatic Purchase Order Allowed' for Tunisia.

    The field must be marked with 'X' when:
      - Plant Type = PP, regardless of Plant
      - Plant = TN01 and Plant Type = FP
      - Plant = TN02, regardless of Plant Type

    For all other combinations, the field must be blank.

    Active checkbox values:
      - X
      - YES
      - Y
      - TRUE
      - 1
      - 1.0

    Inactive checkbox values:
      - blank
      - NO
      - N
      - FALSE
      - 0
      - 0.0

    Unrecognized non-empty values are reported as invalid.
    """

    col = 'Automatic Purchase Order Allowed'

    required_columns = {
        'Plant',
        'Plant Type',
        col,
    }

    if not required_columns.issubset(df.columns):
        return

    # ------------------------------------------------------------
    # Checkbox normalization helper
    # ------------------------------------------------------------
    def _normalize_checkbox(value):
        """
        Normalize a SAP or Excel checkbox value.

        Returns:
          - 'X' when the checkbox is active
          - '' when the checkbox is inactive
          - 'INVALID' for an unrecognized non-empty value
        """

        normalized = _norm_str(value).upper()

        if normalized in {
            'X',
            'YES',
            'Y',
            'TRUE',
            '1',
            '1.0',
        }:
            return 'X'

        if normalized in {
            '',
            'NO',
            'N',
            'FALSE',
            '0',
            '0.0',
        }:
            return ''

        return 'INVALID'

    # ------------------------------------------------------------
    # Validate each material and plant record
    # ------------------------------------------------------------
    for idx, row in df.iterrows():
        errors = []

        material = _norm_str(
            row.get('Material')
        )

        plant = _norm_str(
            row.get('Plant')
        ).upper()

        plant_type = _norm_str(
            row.get('Plant Type')
        ).upper()

        raw_value = _norm_str(
            row.get(col)
        ).upper()

        normalized_value = _normalize_checkbox(
            row.get(col)
        )

        # --------------------------------------------------------
        # Determine whether the checkbox must be marked
        #
        # Required when:
        #   1. Plant Type is PP
        #   2. Plant is TN01 and Plant Type is FP
        #   3. Plant is TN02, regardless of Plant Type
        # --------------------------------------------------------
        must_be_marked = (
            plant_type == 'PP'
            or (
                plant == 'TN01'
                and plant_type == 'FP'
            )
            or plant == 'TN02'
        )

        # --------------------------------------------------------
        # Rule for combinations where the checkbox must be active
        # --------------------------------------------------------
        if must_be_marked:
            if normalized_value != 'X':
                current_value = (
                    raw_value
                    if raw_value != ''
                    else 'blank'
                )

                if plant == 'TN02':
                    reason = (
                        "all materials in TN02 require the field "
                        "to be marked"
                    )

                elif (
                    plant == 'TN01'
                    and plant_type == 'FP'
                ):
                    reason = (
                        "TN01 materials with Plant Type 'FP' "
                        "require the field to be marked"
                    )

                else:
                    reason = (
                        "Plant Type 'PP' requires the field "
                        "to be marked"
                    )

                errors.append(
                    "Automatic Purchase Order Allowed must be 'X' "
                    f"because {reason} "
                    f"(current: '{current_value}')"
                )

        # --------------------------------------------------------
        # Rule for combinations where the field must be blank
        # --------------------------------------------------------
        else:
            if normalized_value == 'X':
                errors.append(
                    "Automatic Purchase Order Allowed must be blank "
                    f"for Plant '{plant}' and Plant Type "
                    f"'{plant_type}' (current: '{raw_value}')"
                )

            elif normalized_value == 'INVALID':
                errors.append(
                    f"Invalid checkbox value: '{raw_value}'. "
                    f"The field must be blank for Plant '{plant}' "
                    f"and Plant Type '{plant_type}'"
                )

        # --------------------------------------------------------
        # Append one aggregated error for the current record
        # --------------------------------------------------------
        if errors:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' | '.join(errors),
                'Material': material,
                'Plant': plant,
            })

def tn_validar_consumption_mode(df, mara_df, resultados):
    """
    Consumption Mode rules for Tunisia:
    - TN01 + Plant Type FP -> must be '5' (takes precedence over Material Type).
    - FERT -> '0' or empty.
    - TN02 HALB -> '0' or empty.
    - All other materials -> must be '2'.
    """
    col = 'Consumption Mode'
    if col not in df.columns:
        return
    mt_map = _mt_map_from_mara(mara_df)
    for idx, row in df.iterrows():
        erros = []
        plant = _norm_str(row.get('Plant')).upper()
        ptype = _norm_str(row.get('Plant Type')).upper()
        mat   = _norm_str(row.get('Material'))
        mt    = mt_map.get(mat, '')
        raw   = row.get(col)
        val   = normalize_code(raw, as_int=True)

        # TN01 FP is the most specific rule and therefore takes precedence,
        # including when the material type is FERT.
        if plant == 'TN01' and ptype == 'FP':
            if val == '':
                erros.append("Mandatory field blank (TN01 FP requires Consumption Mode = '5')")
            elif val != '5':
                erros.append(f"Invalid value: '{raw}' (TN01 FP requires Consumption Mode = '5')")
        elif mt == 'FERT':
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


def tn_validar_gr_processing_time(marc_df, resultados):
    """
    MARC['GR processing time'] must be exactly '3' for TN02.

    The column name is intentionally case-sensitive. In particular, this rule
    does not use or fall back to a similarly named 'GR Processing Time' field.
    """
    col = 'GR processing time'
    if marc_df is None or marc_df.empty or col not in marc_df.columns:
        return

    for idx, row in marc_df.iterrows():
        plant = _norm_str(row.get('Plant')).upper()
        if plant != 'TN02':
            continue

        mat = _norm_str(row.get('Material'))
        raw = row.get(col)
        val = normalize_code(raw, as_int=True)

        if val == '':
            error = "Mandatory field blank (TN02 requires GR processing time = '3')"
        elif val != '3':
            error = f"Invalid value: '{raw}' (TN02 requires GR processing time = '3')"
        else:
            continue

        resultados.append({
            'Line': idx + 2,
            'Field': col,
            'Error': error,
            'Material': mat,
            'Plant': plant,
        })


def tn_validar_jit_delivery_sched(df, resultados):
    """
    Validate JIT Delivery Sched. for Tunisia.

    The field must have value '1' for:
      - Any material with Plant Type = 'PP'
      - Plant TN01 with Plant Type = 'FP'
      - Any material in Plant TN02, regardless of Plant Type

    The field must be empty for all remaining cases.

    Precedence:
      1. Any material in TN02 -> '1'
      2. Plant Type PP -> '1'
      3. TN01 with Plant Type FP -> '1'
      4. All remaining cases -> empty
    """

    col = 'JIT Delivery Sched.'

    if df is None or df.empty:
        return

    required_columns = {
        'Material',
        'Plant',
        'Plant Type',
        col,
    }

    if not required_columns.issubset(df.columns):
        missing_columns = sorted(
            required_columns - set(df.columns)
        )

        print(
            "Missing columns for JIT Delivery Sched. validation: "
            f"{missing_columns}"
        )
        return

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

        raw_value = row.get(col)

        # Normalize possible Excel/SAP representations:
        # 1, 1.0, "1" and "1.0" are all treated as value "1".
        value = normalize_code(
            raw_value,
            as_int=True
        )

        # --------------------------------------------------------
        # Value 1 is mandatory when:
        #
        # 1. The material belongs to TN02, regardless of Plant Type
        # 2. Plant Type is PP
        # 3. Plant is TN01 and Plant Type is FP
        # --------------------------------------------------------
        requires_value_1 = (
            plant == 'TN02'
            or plant_type == 'PP'
            or (
                plant == 'TN01'
                and plant_type == 'FP'
            )
        )

        if requires_value_1:
            if value == '':
                if plant == 'TN02':
                    error = (
                        "Mandatory field blank "
                        "(all TN02 materials require "
                        "JIT Delivery Sched. = '1')"
                    )

                elif plant_type == 'PP':
                    error = (
                        "Mandatory field blank "
                        "(Plant Type 'PP' requires "
                        "JIT Delivery Sched. = '1')"
                    )

                else:
                    error = (
                        "Mandatory field blank "
                        "(TN01 Plant Type 'FP' requires "
                        "JIT Delivery Sched. = '1')"
                    )

                erros.append(error)

            elif value != '1':
                if plant == 'TN02':
                    error = (
                        "All TN02 materials require "
                        "JIT Delivery Sched. = '1' "
                        f"(current: '{raw_value}')"
                    )

                elif plant_type == 'PP':
                    error = (
                        "Plant Type 'PP' requires "
                        "JIT Delivery Sched. = '1' "
                        f"(current: '{raw_value}')"
                    )

                else:
                    error = (
                        "TN01 Plant Type 'FP' requires "
                        "JIT Delivery Sched. = '1' "
                        f"(current: '{raw_value}')"
                    )

                erros.append(error)

        # --------------------------------------------------------
        # Remaining cases must be empty
        # --------------------------------------------------------
        else:
            if value != '':
                erros.append(
                    "JIT Delivery Sched. must be empty "
                    f"for Plant '{plant}' and "
                    f"Plant Type '{plant_type}' "
                    f"(current: '{raw_value}')"
                )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' \n'.join(erros),
                'Material': material,
                'Plant': plant,
            })

def tn_validar_mrp_group(marc_df, mara_df, resultados):
    """
    Tunisia MRP Group validation:

    - TN02, any material and any Material Type:
        MRP Group is mandatory and must be 'ZSTO'.

    - TN01:
        No validation is applied by this function.
    """

    col = 'MRP Group'

    required_columns = {'Material', 'Plant', col}
    if (
        marc_df is None
        or marc_df.empty
        or not required_columns.issubset(marc_df.columns)
    ):
        return

    for idx, row in marc_df.iterrows():
        plant = _norm_str(row.get('Plant')).upper()
        material = _norm_str(row.get('Material'))
        value = _norm_str(row.get(col)).upper()

        # This rule applies to every material in TN02,
        # regardless of Material Type or Plant Type.
        if plant != 'TN02':
            continue

        if value == '':
            error = (
                "Mandatory field blank "
                "(TN02 requires MRP Group = 'ZSTO')"
            )
        elif value != 'ZSTO':
            error = (
                f"TN02 requires MRP Group = 'ZSTO' "
                f"(current: '{value}')"
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


def tn_validar_mrp_type(
    df,
    mara_df,
    resultados,
    zplm_marc_df=None,
    csr_item_df=None,
    csr_head_df=None
):
    """
    Validate MRP Type according to the current rule.

    Rules and precedence:

    1. PT02 with Plant Specific Material Status = 25:
         MRP Type must be 'ND'.

    2. Any material extended to TN02:
         MRP Type must be 'PD'.

    3. ROH and HALB materials in TN01 or TN02:
         MRP Type must be 'PD'.

    4. FERT materials outside TN02:
         MRP Type must be 'M1'.

    Notes:
      - The TN02 rule takes precedence over the FERT rule.
      - Therefore, a FERT material in TN02 must have MRP Type = 'PD'.
      - zplm_marc_df, csr_item_df and csr_head_df are kept in the
        function signature for compatibility with the existing call.
      - Planning Time Fence and Planning Calendar are not validated
        by this function.
    """

    if df is None or df.empty:
        return

    required_columns = {
        'Material',
        'Plant',
        'MRP Type',
    }

    if not required_columns.issubset(df.columns):
        return

    # Material -> Material Type
    mt_map = _mt_map_from_mara(mara_df)

    psms_col = next(
        (
            column
            for column in [
                'Plant Specific Material Status',
                'Plant-Specific Material Status',
            ]
            if column in df.columns
        ),
        None
    )

    def _normalize_status(value):
        """
        Normalize statuses exported as:
          25
          25.0
          '25'
          '25.0'
        """

        if value is None or pd.isna(value):
            return ''

        value_str = str(value).strip()

        if value_str == '':
            return ''

        try:
            return str(int(float(value_str)))
        except (ValueError, TypeError):
            return value_str.upper()

    for idx, row in df.iterrows():
        material = _norm_str(
            row.get('Material')
        )

        plant = _norm_str(
            row.get('Plant')
        ).upper()

        mrp_type = _norm_str(
            row.get('MRP Type')
        ).upper()

        material_type = mt_map.get(
            material,
            _norm_str(
                row.get('Material Type', '')
            ).upper()
        )

        psms = (
            _normalize_status(row.get(psms_col))
            if psms_col is not None
            else ''
        )

        expected_mrp_type = None
        rule_description = None

        # ========================================================
        # RULE 1
        # PT02 + PSMS 25 -> ND
        #
        # This is the most specific rule and therefore has the
        # highest precedence.
        # ========================================================
        if plant == 'PT02' and psms == '25':
            expected_mrp_type = 'ND'
            rule_description = (
                "PT02 with Plant Specific Material Status 25"
            )

        # ========================================================
        # RULE 2
        # Any TN02 material -> PD
        #
        # This includes:
        #   - FERT
        #   - HALB
        #   - ROH
        # ========================================================
        elif plant == 'TN02':
            expected_mrp_type = 'PD'
            rule_description = (
                "all materials in TN02"
            )

        # ========================================================
        # RULE 3
        # ROH and HALB in TN01/TN02 -> PD
        #
        # TN02 is already handled above. This condition therefore
        # mainly covers ROH and HALB materials in TN01.
        # ========================================================
        elif (
            plant in {'TN01', 'TN02'}
            and material_type in {'ROH', 'HALB'}
        ):
            expected_mrp_type = 'PD'
            rule_description = (
                f"Material Type {material_type} in {plant}"
            )

        # ========================================================
        # RULE 4
        # FERT outside TN02 -> M1
        # ========================================================
        elif (
            material_type == 'FERT'
            and plant != 'TN02'
        ):
            expected_mrp_type = 'M1'
            rule_description = (
                f"Material Type FERT outside TN02 ({plant})"
            )

        # No rule defined for this row
        if expected_mrp_type is None:
            continue

        # Correct value
        if mrp_type == expected_mrp_type:
            continue

        if mrp_type == '':
            error = (
                f"Mandatory field blank: {rule_description} "
                f"requires MRP Type = '{expected_mrp_type}'"
            )
        else:
            error = (
                f"{rule_description} requires MRP Type = "
                f"'{expected_mrp_type}' "
                f"(current: '{mrp_type}')"
            )

        resultados.append({
            'Line': idx + 2,
            'Field': 'MRP Type',
            'Error': error,
            'Material': material,
            'Plant': plant,
        })

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
    """
    Tunisia Planning Time Fence validation.

    Rules and precedence:

    1. Any material in TN02:
         Planning Time Fence must be 0.

    2. ROH and HALB materials:
         Planning Time Fence must be 0.

    3. FERT materials outside TN02:
         Planning Time Fence must be 1.

    Notes:
      - The TN02 rule takes precedence over the FERT rule.
      - Therefore, a FERT material in TN02 must have value 0.
      - Blank values are not accepted as 0.
      - The rule no longer depends on MRP Type.
    """

    if df is None or df.empty:
        return

    required_columns = {
        'Material',
        'Plant',
        'Planning Time Fence',
    }

    if not required_columns.issubset(df.columns):
        return

    mt_map = _mt_map_from_mara(mara_df)

    def _normalize_time_fence(value):
        """
        Converts Planning Time Fence values to a numeric value.

        Handles:
          - 0 and 1
          - 0.0 and 1.0
          - '0' and '1'
          - '0.0' and '1.0'
          - Decimal comma values such as '0,0' and '1,0'

        Returns:
          - None for blank or NaN
          - 'INVALID' for non-numeric values
          - float for valid numeric values
        """

        if value is None or pd.isna(value):
            return None

        value_str = str(value).strip()

        if value_str == '':
            return None

        value_str = value_str.replace(',', '.')

        try:
            return float(value_str)
        except (ValueError, TypeError):
            return 'INVALID'

    for idx, row in df.iterrows():
        material = _norm_str(
            row.get('Material')
        )

        plant = _norm_str(
            row.get('Plant')
        ).upper()

        material_type = mt_map.get(
            material,
            _norm_str(
                row.get('Material Type', '')
            ).upper()
        )

        time_fence_raw = row.get(
            'Planning Time Fence'
        )

        time_fence = _normalize_time_fence(
            time_fence_raw
        )

        expected_time_fence = None
        rule_description = None

        # ========================================================
        # RULE 1
        # Any material in TN02 -> 0
        #
        # This rule has the highest precedence and includes:
        #   - FERT
        #   - HALB
        #   - ROH
        # ========================================================
        if plant == 'TN02':
            expected_time_fence = 0.0
            rule_description = (
                "all materials in TN02"
            )

        # ========================================================
        # RULE 2
        # ROH and HALB -> 0
        # ========================================================
        elif material_type in {'ROH', 'HALB'}:
            expected_time_fence = 0.0
            rule_description = (
                f"Material Type {material_type}"
            )

        # ========================================================
        # RULE 3
        # FERT outside TN02 -> 1
        # ========================================================
        elif material_type == 'FERT':
            expected_time_fence = 1.0
            rule_description = (
                f"Material Type FERT outside TN02 ({plant})"
            )

        # No rule defined for this Material Type
        if expected_time_fence is None:
            continue

        # --------------------------------------------------------
        # Invalid non-numeric value
        # --------------------------------------------------------
        if time_fence == 'INVALID':
            error = (
                "Planning Time Fence contains a non-numeric value: "
                f"'{time_fence_raw}'. "
                f"For {rule_description}, the expected value is "
                f"'{int(expected_time_fence)}'"
            )

        # --------------------------------------------------------
        # Blank value
        #
        # Blank is not treated as 0 because the requirement says
        # that the field must have value 0 or 1.
        # --------------------------------------------------------
        elif time_fence is None:
            error = (
                "Mandatory field blank: "
                f"{rule_description} requires Planning Time Fence = "
                f"'{int(expected_time_fence)}'"
            )

        # --------------------------------------------------------
        # Incorrect numeric value
        # --------------------------------------------------------
        elif abs(time_fence - expected_time_fence) >= 1e-6:
            error = (
                f"{rule_description} requires Planning Time Fence = "
                f"'{int(expected_time_fence)}' "
                f"(current: '{time_fence_raw}')"
            )

        # Correct value
        else:
            continue

        resultados.append({
            'Line': idx + 2,
            'Field': 'Planning Time Fence',
            'Error': error,
            'Material': material,
            'Plant': plant,
        })

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
    """
    Tunisia Purchasing Group rules:

    - TN02, any material and any Plant Type:
        Purchasing Group must be 'STO'.

    - TN01:
        - FP: Purchasing Group must be 'STO'.
        - PP: Purchasing Group is mandatory.
        - MP: Purchasing Group must be blank.

    TN02 takes precedence over the Plant Type rules.
    """

    col = 'Purchasing Group'

    required_columns = {'Material', 'Plant', 'Plant Type', col}
    if not required_columns.issubset(df.columns):
        return

    for idx, row in df.iterrows():
        erros = []

        material = _norm_str(row.get('Material'))
        plant = _norm_str(row.get('Plant')).upper()
        plant_type = _norm_str(row.get('Plant Type')).upper()
        value = _norm_str(row.get(col)).upper()

        # ============================================================
        # TN02
        # All materials and all Plant Types must have STO.
        # This rule takes precedence over the Plant Type rules.
        # ============================================================
        if plant == 'TN02':
            if value == '':
                erros.append(
                    "Mandatory field blank for TN02 "
                    "(Purchasing Group must be 'STO')"
                )
            elif value != 'STO':
                erros.append(
                    f"For TN02, Purchasing Group must be 'STO' "
                    f"(current: '{value}')"
                )

        # ============================================================
        # TN01
        # ============================================================
        elif plant == 'TN01':

            # TN01 FP -> STO
            if plant_type == 'FP':
                if value == '':
                    erros.append(
                        "Mandatory field blank for TN01 Plant Type 'FP' "
                        "(Purchasing Group must be 'STO')"
                    )
                elif value != 'STO':
                    erros.append(
                        f"For TN01 Plant Type 'FP', Purchasing Group "
                        f"must be 'STO' (current: '{value}')"
                    )

            # TN01 PP -> mandatory
            elif plant_type == 'PP':
                if value == '':
                    erros.append(
                        "Mandatory field blank for TN01 Plant Type 'PP'"
                    )

            # TN01 MP -> blank
            elif plant_type == 'MP':
                if value != '':
                    erros.append(
                        f"For TN01 Plant Type 'MP', Purchasing Group "
                        f"must be blank (current: '{value}')"
                    )

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': ' | '.join(erros),
                'Material': material,
                'Plant': plant,
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
    """
    Tunisia:
    Source List must always be blank.

    Supports both fields:
    - Source list
    - Source list required
    """

    col = next(
        (
            c
            for c in ['Source list', 'Source list required']
            if c in df.columns
        ),
        None
    )

    if col is None:
        return

    for idx, row in df.iterrows():

        plant = _norm_str(row.get('Plant')).upper()
        mat   = _norm_str(row.get('Material'))
        val   = _norm_str(row.get(col)).upper()

        if val != '':
            resultados.append({
                'Line': idx + 2,
                'Field': col,
                'Error': f"Field must be blank (current: '{val}')",
                'Material': mat,
                'Plant': plant
            })

def tn_validar_storage_location(df, mara_df, resultados):
    """
    Issue Storage Location validation for Tunisia.

    TN01, Material Types FERT/HALB/ROH:
      - Plant Type FP:
          Issue Storage Location must be 'T120'.
      - Plant Type MP or PP:
          Issue Storage Location must be 'T105'.

    TN02, Material Types FERT/HALB/ROH:
      - Plant Type FP, MP or PP:
          Issue Storage Location must be 'T220'.

    Materials or Plant Types outside the defined scope are not validated
    by this function.
    """

    col = 'Issue Storage Location'

    if df is None or df.empty:
        return

    required_columns = {
        'Material',
        'Plant',
        'Plant Type',
        col,
    }

    if not required_columns.issubset(df.columns):
        missing_columns = sorted(
            required_columns - set(df.columns)
        )

        print(
            "Missing columns for TN Issue Storage Location validation: "
            f"{missing_columns}"
        )
        return

    mt_map = _mt_map_from_mara(mara_df)

    valid_material_types = {
        'FERT',
        'HALB',
        'ROH',
    }

    valid_plant_types = {
        'FP',
        'MP',
        'PP',
    }

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

        material_type = mt_map.get(
            material,
            _norm_str(
                row.get('Material Type', '')
            ).upper()
        )

        value = _norm_str(
            row.get(col)
        ).upper()

        # Only validate the explicitly defined Material Types
        # and Plant Types.
        if (
            material_type not in valid_material_types
            or plant_type not in valid_plant_types
        ):
            continue

        expected_value = None

        # ========================================================
        # TN02
        #
        # FERT/HALB/ROH + FP/MP/PP -> T220
        # ========================================================
        if plant == 'TN02':
            expected_value = 'T220'

        # ========================================================
        # TN01
        #
        # FERT/HALB/ROH:
        #   FP    -> T120
        #   MP/PP -> T105
        # ========================================================
        elif plant == 'TN01':
            if plant_type == 'FP':
                expected_value = 'T120'
            elif plant_type in {'MP', 'PP'}:
                expected_value = 'T105'

        # Plant outside TN01/TN02
        else:
            continue

        # No expected value was determined
        if expected_value is None:
            continue

        # ========================================================
        # Validate mandatory value
        # ========================================================
        if value == '':
            error = (
                "Mandatory field blank "
                f"({plant} + Material Type '{material_type}' + "
                f"Plant Type '{plant_type}' requires "
                f"Issue Storage Location = '{expected_value}')"
            )

        elif value != expected_value:
            error = (
                f"For {plant} + Material Type '{material_type}' + "
                f"Plant Type '{plant_type}', "
                f"Issue Storage Location must be '{expected_value}' "
                f"(current: '{value}')"
            )

        else:
            continue

        erros.append(error)

        resultados.append({
            'Line': idx + 2,
            'Field': col,
            'Error': ' | '.join(erros),
            'Material': material,
            'Plant': plant,
        })

def tn_validar_special_procurement_type(
    marc_df,
    zplm_marc_df,
    resultados
):
    """
    Tunisia Special Procurement Type validation.

    Rule:
      - Any material in TN02 must have
        Special Procurement Type = 'Z3'.

    Notes:
      - The rule applies regardless of:
          * Material Type
          * Plant Type
      - TN01 is not validated by this function.
      - zplm_marc_df is kept in the function signature for
        compatibility with the existing orchestrator call.
    """

    col = 'Special Procurement Type'

    if marc_df is None or marc_df.empty:
        return

    required_columns = {
        'Material',
        'Plant',
        col,
    }

    if not required_columns.issubset(marc_df.columns):
        return

    for idx, row in marc_df.iterrows():
        plant = _norm_str(
            row.get('Plant')
        ).upper()

        material = _norm_str(
            row.get('Material')
        )

        value = _norm_str(
            row.get(col)
        ).upper()

        # This rule applies to every material in TN02,
        # regardless of Material Type or Plant Type.
        if plant != 'TN02':
            continue

        if value == '':
            error = (
                "Mandatory field blank "
                "(all TN02 materials require "
                "Special Procurement Type = 'Z3')"
            )

        elif value != 'Z3':
            error = (
                "All TN02 materials require "
                "Special Procurement Type = 'Z3' "
                f"(current: '{value}')"
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


def tn_validar_storage_location_ext_procurement(
    marc_df,
    zplm_marc_df,
    mara_df,
    resultados
):
    """
    Validate MARC field 'Storage Location for Ext. Procurement'
    for Tunisia.

    Scope:
      - Material Type: FERT, HALB or ROH
      - Plant Type: FP, MP or PP

    Rules:

    TN01:
      - FP -> T120
      - MP -> T115
      - PP -> T100

    TN02:
      - FP -> T220
      - MP -> T220
      - PP -> T220

    For TN02, T220 is mandatory for all valid combinations of
    FERT/HALB/ROH and FP/MP/PP.
    """

    col = 'Storage Location for Ext. Procurement'

    required_marc_columns = {
        'Material',
        'Plant',
        col,
    }

    required_mara_columns = {
        'Material',
        'Material Type',
    }

    if (
        marc_df is None
        or mara_df is None
        or marc_df.empty
        or mara_df.empty
        or not required_marc_columns.issubset(marc_df.columns)
        or not required_mara_columns.issubset(mara_df.columns)
    ):
        return

    # Material + Plant -> Plant Type
    pt_map = _pt_map_from_zplm(zplm_marc_df)

    # Material -> Material Type
    mt_map = _mt_map_from_mara(mara_df)

    # ------------------------------------------------------------
    # Normalize MARC data
    # ------------------------------------------------------------
    marc = marc_df.copy()

    marc.columns = [
        str(column).strip()
        for column in marc.columns
    ]

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

    # One validation per Material + Plant
    marc = marc.drop_duplicates(
        subset=['Material', 'Plant'],
        keep='first'
    )

    # ------------------------------------------------------------
    # Expected values
    # ------------------------------------------------------------
    tn01_by_plant_type = {
        'FP': 'T120',
        'MP': 'T115',
        'PP': 'T100',
    }

    valid_material_types = {
        'FERT',
        'HALB',
        'ROH',
    }

    valid_plant_types = {
        'FP',
        'MP',
        'PP',
    }

    # ------------------------------------------------------------
    # Validate each MARC row
    # ------------------------------------------------------------
    for idx, row in marc.iterrows():
        material = _norm_str(
            row.get('Material')
        )

        plant = _norm_str(
            row.get('Plant')
        ).upper()

        plant_type = _norm_str(
            pt_map.get((material, plant), '')
        ).upper()

        material_type = _norm_str(
            mt_map.get(material, '')
        ).upper()

        value = _norm_str(
            row.get(col)
        ).upper()

        # Only validate FERT, HALB and ROH
        if material_type not in valid_material_types:
            continue

        # Only validate FP, MP and PP
        if plant_type not in valid_plant_types:
            continue

        # ========================================================
        # TN01
        # FP -> T120
        # MP -> T115
        # PP -> T100
        # ========================================================
        if plant == 'TN01':
            expected_value = tn01_by_plant_type[plant_type]

        # ========================================================
        # TN02
        # All Plant Types and Material Types in scope -> T220
        # ========================================================
        elif plant == 'TN02':
            expected_value = 'T220'

        # Other plants are outside the scope of this TN function
        else:
            continue

        # ========================================================
        # Validate expected value
        # ========================================================
        if value == '':
            error = (
                "Mandatory field blank "
                f"(expected '{expected_value}' for {plant} / "
                f"Material Type '{material_type}' / "
                f"Plant Type '{plant_type}')"
            )

        elif value != expected_value:
            error = (
                f"Invalid value: '{value}' "
                f"(expected '{expected_value}' for {plant} / "
                f"Material Type '{material_type}' / "
                f"Plant Type '{plant_type}')"
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


def tn_validar_advanced_planning_fields(
    marc_df,
    mara_df,
    resultados,
    matplsta_df=None
):
    """
    Tunisia Advanced Planning validation.

    TN02 override:
      - Advanced Planning must be ON for every TN02 material.
      - Product Heuristic must be 'ZVIS_PP_002' for every TN02 material.
      - Proposed Strategy must be 'Z1' for every TN02 material.

    TN01 FP/MP:
      - Proposed Strategy must be 'Z1'.
      - Advanced Planning is mandatory for TN01 FP.

    General FP/MP:
      - Proposed Strategy must be '20' unless a TN01/TN02 override applies.
      - PP Planning Procedure must be '4'.
      - Product Heuristic is mandatory.

    PP outside the TN02 override:
      - Planning Status 9 or empty:
          Do not validate Advanced Planning, Product Heuristic,
          PP Planning Procedure or Proposed Strategy.
      - Planning Status 1:
          Advanced Planning must be OFF.
          Product Heuristic, PP Planning Procedure and
          Proposed Strategy must be empty.

    PSMS rules for FP/MP:
      - PSMS 25 or 75:
          Advanced Planning may be OFF regardless of Planning Status.
          Advanced Planning may be ON only when Planning Status is 9.

    Important precedence:
      1. TN02 rules
      2. PP rules for plants other than TN02
      3. PSMS 25/75 rules for FP/MP
      4. General FP/MP rules
    """

    if marc_df is None or marc_df.empty:
        return

    # -------------------------------------------------------------------------
    # Helpers
    # -------------------------------------------------------------------------
    def _pick(candidates):
        return next(
            (
                column
                for column in candidates
                if column in marc_df.columns
            ),
            None
        )

    def _s(value):
        if value is None or pd.isna(value):
            return ''

        return str(value).strip()

    def is_true(value):
        return _s(value).upper() in {
            'TRUE',
            'X',
            '1',
            '1.0',
            'YES',
            'Y',
        }

    def is_false(value):
        return _s(value).upper() in {
            'FALSE',
            '',
            '0',
            '0.0',
            'NO',
            'N',
        }

    def add_error(
        idx,
        field,
        error,
        material,
        plant
    ):
        resultados.append({
            'Line': idx + 2,
            'Field': field,
            'Error': error,
            'Material': material,
            'Plant': plant,
        })

    # -------------------------------------------------------------------------
    # Identify required columns
    # -------------------------------------------------------------------------
    col_adv = _pick([
        'Advanced Planning',
    ])

    col_heur = _pick([
        'Heuristic',
        'Product Heuristic',
    ])

    col_pp = _pick([
        'PP Plng Procedure',
        'PP Planning Procedure',
    ])

    col_req = _pick([
        'Requirement Strategy',
        'Proposed Strategy',
    ])

    col_psms = _pick([
        'Plant-Specific Material Status',
        'Plant Specific Material Status',
    ])

    col_pt = _pick([
        'Plant Type',
        'Material Plant Type',
    ])

    required_columns = [
        col_adv,
        col_heur,
        col_pp,
        col_req,
        col_psms,
        col_pt,
    ]

    if any(column is None for column in required_columns):
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

                planning_status_map[
                    (material_key, plant_key)
                ] = normalize_code(
                    source_row.get('Planning Status'),
                    as_int=True
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
            row.get(col_pt)
        ).upper()

        # Current validation scope
        if plant_type not in {'FP', 'MP', 'PP'}:
            continue

        advanced_raw = row.get(col_adv)

        heuristic = _s(
            row.get(col_heur)
        ).upper()

        procedure = normalize_code(
            row.get(col_pp),
            as_int=True
        ).upper()

        strategy = normalize_code(
            row.get(col_req),
            as_int=True
        ).upper()

        psms = normalize_code(
            row.get(col_psms),
            as_int=True
        )

        planning_status = planning_status_map.get(
            (material, plant),
            ''
        )

        # =====================================================================
        # TN02 OVERRIDE
        #
        # Applies to every TN02 material in the FP/MP/PP validation scope.
        # These rules take precedence over the standard PP and PSMS rules.
        # =====================================================================
        if plant == 'TN02':

            # -------------------------------------------------------------
            # Advanced Planning
            # Every TN02 material must have Advanced Planning ON.
            # -------------------------------------------------------------
            if not is_true(advanced_raw):
                current_advanced = (
                    _s(advanced_raw)
                    if _s(advanced_raw) != ''
                    else 'blank'
                )

                add_error(
                    idx,
                    'Advanced Planning',
                    (
                        "Mandatory for all TN02 materials: "
                        "Advanced Planning must be 'X' or 'TRUE' "
                        f"(current: '{current_advanced}')"
                    ),
                    material,
                    plant
                )

            # -------------------------------------------------------------
            # Product Heuristic
            # Every TN02 material must have ZVIS_PP_002.
            # -------------------------------------------------------------
            if heuristic == '':
                add_error(
                    idx,
                    'Heuristic',
                    (
                        "Mandatory for all TN02 materials: "
                        "Product Heuristic must be 'ZVIS_PP_002'"
                    ),
                    material,
                    plant
                )

            elif heuristic != 'ZVIS_PP_002':
                add_error(
                    idx,
                    'Heuristic',
                    (
                        "Invalid Product Heuristic for TN02: "
                        "must be 'ZVIS_PP_002' "
                        f"(current: '{heuristic}')"
                    ),
                    material,
                    plant
                )

            # -------------------------------------------------------------
            # Proposed Strategy
            # Every TN02 material must have Z1.
            # -------------------------------------------------------------
            if strategy == '':
                add_error(
                    idx,
                    'Proposed Strategy',
                    (
                        "Mandatory for all TN02 materials: "
                        "Proposed Strategy must be 'Z1'"
                    ),
                    material,
                    plant
                )

            elif strategy != 'Z1':
                add_error(
                    idx,
                    'Proposed Strategy',
                    (
                        "Invalid Proposed Strategy for TN02: "
                        "must be 'Z1' "
                        f"(current: '{strategy}')"
                    ),
                    material,
                    plant
                )

            # -------------------------------------------------------------
            # PP Planning Procedure
            #
            # Existing FP/MP rule remains 4.
            # No new explicit PP rule was provided for this field.
            # -------------------------------------------------------------
            if plant_type in {'FP', 'MP'}:
                if procedure == '':
                    add_error(
                        idx,
                        'PP Plng Procedure',
                        (
                            "Mandatory for TN02 Plant Type "
                            f"'{plant_type}': must be '4'"
                        ),
                        material,
                        plant
                    )

                elif procedure != '4':
                    add_error(
                        idx,
                        'PP Plng Procedure',
                        (
                            "Invalid PP Planning Procedure for TN02 "
                            f"Plant Type '{plant_type}': must be '4' "
                            f"(current: '{procedure}')"
                        ),
                        material,
                        plant
                    )

            # TN02 is fully handled by the override.
            continue

        # =====================================================================
        # PLANT TYPE PP OUTSIDE TN02
        # =====================================================================
        if plant_type == 'PP':

            # PP with Planning Status 9 or empty:
            # Advanced Planning-related fields are not validated.
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
                            "Must be empty for Plant Type 'PP' "
                            "with Planning Status 1 "
                            f"(current: '{heuristic}')"
                        ),
                        material,
                        plant
                    )

                if procedure != '':
                    add_error(
                        idx,
                        'PP Plng Procedure',
                        (
                            "Must be empty for Plant Type 'PP' "
                            "with Planning Status 1 "
                            f"(current: '{procedure}')"
                        ),
                        material,
                        plant
                    )

                if strategy != '':
                    add_error(
                        idx,
                        'Proposed Strategy',
                        (
                            "Must be empty for Plant Type 'PP' "
                            "with Planning Status 1 "
                            f"(current: '{strategy}')"
                        ),
                        material,
                        plant
                    )

                if is_true(advanced_raw):
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

            # PP with other Planning Status values is outside
            # the currently defined standard validation scope.
            continue

        # =====================================================================
        # PLANT TYPES FP AND MP OUTSIDE TN02
        # =====================================================================

        # ---------------------------------------------------------------------
        # PSMS 25/75 precedence
        #
        # For FP/MP outside TN02:
        #   - Advanced Planning may be OFF.
        #   - Advanced Planning may only be ON when Planning Status is 9.
        # ---------------------------------------------------------------------
        if psms in {'25', '75'}:
            advanced_is_allowed = (
                is_false(advanced_raw)
                or (
                    is_true(advanced_raw)
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

                # Do not continue here because the independent
                # field-value validations must still be executed.

        # ---------------------------------------------------------------------
        # Planning Status 9 restriction
        # ---------------------------------------------------------------------
        if (
            planning_status == '9'
            and psms not in {'25', '75'}
        ):
            add_error(
                idx,
                'Advanced Planning',
                (
                    "Planning Status 9 not allowed for Plant Type "
                    f"'{plant_type}' unless PSMS is 25 or 75"
                ),
                material,
                plant
            )

            # This status combination is invalid, so do not continue
            # with the standard Advanced Planning requirement.
            continue

        # ---------------------------------------------------------------------
        # Advanced Planning requirement outside TN02
        #
        # TN01 FP remains mandatory.
        # TN01 MP is not made mandatory by the rules provided.
        # ---------------------------------------------------------------------
        if (
            plant == 'TN01'
            and plant_type == 'FP'
            and not is_true(advanced_raw)
        ):
            current_advanced = (
                _s(advanced_raw)
                if _s(advanced_raw) != ''
                else 'blank'
            )

            add_error(
                idx,
                'Advanced Planning',
                (
                    "Mandatory for TN01 Plant Type 'FP': "
                    "Advanced Planning must be 'X' or 'TRUE' "
                    f"(current: '{current_advanced}')"
                ),
                material,
                plant
            )

        # ---------------------------------------------------------------------
        # Product Heuristic for FP/MP outside TN02
        #
        # No fixed value was provided outside TN02.
        # Therefore, the existing mandatory-field rule is retained.
        # ---------------------------------------------------------------------
        if heuristic == '':
            add_error(
                idx,
                'Heuristic',
                (
                    "Mandatory: Product Heuristic cannot be empty "
                    f"for Plant Type '{plant_type}'"
                ),
                material,
                plant
            )

        # ---------------------------------------------------------------------
        # PP Planning Procedure for FP/MP
        # ---------------------------------------------------------------------
        if procedure == '':
            add_error(
                idx,
                'PP Plng Procedure',
                (
                    "Mandatory: PP Planning Procedure must be '4' "
                    f"for Plant Type '{plant_type}'"
                ),
                material,
                plant
            )

        elif procedure != '4':
            add_error(
                idx,
                'PP Plng Procedure',
                (
                    "Invalid PP Planning Procedure: must be '4' "
                    f"for Plant Type '{plant_type}' "
                    f"(current: '{procedure}')"
                ),
                material,
                plant
            )

        # ---------------------------------------------------------------------
        # Proposed Strategy
        #
        # Overrides:
        #   - TN01 FP/MP -> Z1
        #   - TN02 all materials -> Z1, already handled above
        #
        # General FP/MP:
        #   - 20
        # ---------------------------------------------------------------------
        if (
            plant == 'TN01'
            and plant_type in {'FP', 'MP'}
        ):
            expected_strategy = 'Z1'
        else:
            expected_strategy = '20'

        if strategy == '':
            add_error(
                idx,
                'Proposed Strategy',
                (
                    "Mandatory: Proposed Strategy must be "
                    f"'{expected_strategy}' for Plant '{plant}' "
                    f"and Plant Type '{plant_type}'"
                ),
                material,
                plant
            )

        elif strategy != expected_strategy:
            add_error(
                idx,
                'Proposed Strategy',
                (
                    f"Invalid Proposed Strategy: must be "
                    f"'{expected_strategy}' for Plant '{plant}' "
                    f"and Plant Type '{plant_type}' "
                    f"(current: '{strategy}')"
                ),
                material,
                plant
            )

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
    mard_file=None,
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
    _mard_df=None,
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
    mard_df     = _mard_df if _mard_df is not None else (load_xlsx_fast(mard_file, sheet_name=0) if mard_file else None)

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
        'mrp group': 'MRP Group',
        'disgr': 'MRP Group',
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

    # ✅ DEDUPLICATE ZPLM MARC to one row per (Material, Plant) — BEFORE any
    # consumer touches it. See sap_validation_engine.py's run_full_validation
    # for the full explanation: ZPLM MARC's change-history columns suggest
    # it can hold multiple historical rows per (Material, Plant), which
    # would otherwise multiply every validation that merges against it
    # directly (validar_batch_management_plant_marc, validar_version_indicator_marc,
    # validar_country_region_origin_fp, tn_validar_storage_location_ext_procurement).
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
    mard_tn      = _ftn(mard_df) if mard_df is not None else None
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
    advanced_planning_df = sap_like_df.copy()

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
    # Only run MBEW-derived checks when there are rows in the selected TN scope.
    # This avoids false "column not found" warnings when All Sites is selected
    # but the uploaded files contain no TN01/TN02 records.
    if sap_like_df is not None and not sap_like_df.empty:
        validar_price_control(sap_like_df, resultados)
        validar_price_unit(sap_like_df, resultados)

    _cb(30, "Validating Product Hierarchy...")
    validar_product_hierarchy(sap_like_df, resultados, mvke_df, mara_df)
    validar_product_hierarchy_mvke(sap_like_df, resultados, mvke_df)

    _cb(31, "Validating Valuation Class, With Qty Structure...")
    if sap_like_df is not None and not sap_like_df.empty:
        validar_valuation_class(sap_like_df, resultados)
        validar_with_qty_structure(sap_like_df, resultados)

    _cb(32, "Validating Product-Related Origin, MRP Controller, PSMS...")
    validar_product_related_origin(sap_like_df, resultados)
    validar_mrp_controller(sap_like_df, resultados, mard_tn)
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

    _cb(47.5, "Validating MRP Group (TN02 FERT)...")
    tn_validar_mrp_group(marc_tn, mara_df, resultados)

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
    if marc_tn is not None and not marc_tn.empty:
        validar_per_lot_sz_plng_cal(advanced_planning_df, resultados, mara_df, matplsta_df)

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

    _cb(74.2, "Validating GR processing time (TN02 MARC)...")
    tn_validar_gr_processing_time(marc_tn, resultados)

    _cb(74.4, "Validating Material Price Determination Control (MBEW)...")
    if mbew_df is not None:
        validar_material_price_determination_control(mbew_df, mara_df, resultados)

    _cb(75, "Validating Advanced Planning fields (TN)...")
    tn_validar_advanced_planning_fields(advanced_planning_df, mara_df, resultados, matplsta_df)

    _cb(76, "Validating Neg. stocks in plant (MARC)...")
    neg_errors = validar_neg_stocks_in_plant(marc_tn, mara_df)
    if neg_errors:
        resultados.extend(neg_errors)

    # ── 5. Build errors DataFrame ─────────────────────────────────────────────
    _cb(78, "Building errors DataFrame...")
    errors_df = _coerce_df_for_display(pd.DataFrame(resultados))

    # Keep a stable schema even when the selected site has zero errors/rows.
    # Without this, pd.DataFrame([]) has no columns and the Roles check below
    # raises KeyError: 'Field'.
    _error_columns = ['Line', 'Field', 'Error', 'Material', 'Plant']
    for _col in _error_columns:
        if _col not in errors_df.columns:
            errors_df[_col] = pd.Series(dtype='object')

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

    if (roles_df is not None and not roles_df.empty
            and not errors_df.empty and 'Field' in errors_df.columns
            and 'Field' in roles_df.columns):
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