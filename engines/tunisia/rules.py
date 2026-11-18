"""Regras específicas da Tunísia (TN01/TN02) — exceções ao rule book de Palmela."""

import pandas as pd

from engines.common.helpers import _norm_str, normalize_code
from engines.material_rules.fields import _prepare_mvke
from engines.material_rules.sap_like import _mt_map_from_mara, _pt_map_from_zplm

TN_PLANTS = {"TN01", "TN02"}


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
