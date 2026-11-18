"""Pós-processamento: Old Material Number, Enterproj, stocks negativos."""

import pandas as pd

from engines.common.fast_loader import load_xlsx_fast

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
