"""Regras de Advanced Planning (MRP) partilhadas por Palmela e Tunísia."""

import pandas as pd

from engines.common.helpers import normalize_code

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
