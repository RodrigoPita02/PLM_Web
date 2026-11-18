"""Regras B) — validações cruzadas entre ficheiros (MARA/MARC/MVKE/MBEW/...)."""

import math
import pandas as pd

from engines.common.helpers import _find_col_ci, _norm_str, normalize_code
from engines.material_rules.batch_utils import get_batch_management_effective_material_type
from engines.material_rules.constants import (
    MARA_MAINT_STATUS_C_EXEMPT_MATERIALS,
    MARA_MAINT_STATUS_C_FIELD,
)
from engines.material_rules.fields import _prepare_mvke
from engines.material_rules.sap_like import (
    _mt_map_from_mara,
    _pt_map_from_zplm,
    build_material_transfer_map,
    get_effective_material_type,
)

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


def validar_reference_matl_for_packing_mara(mara_df, resultados):
    """
    MARA: 'Reference matl for packing' must be 'REF_PACK_MAT_EA'
    for Material Types FERT, HALB and ROH.
    Validation is material-level (field belongs to MARA).
    """
    if mara_df is None or mara_df.empty:
        return

    ref_col = _find_col_ci(mara_df, [
        'Reference matl for packing',
        'Reference material for packing',
        'Ref. material for packing',
        'Ref. matl for packing',
    ])
    mat_col = _find_col_ci(mara_df, ['Material'])
    mt_col = _find_col_ci(mara_df, ['Material Type'])

    if ref_col is None or mat_col is None or mt_col is None:
        print("⚠️ MARA is missing 'Reference matl for packing' (or Material / Material Type).")
        return

    expected = 'REF_PACK_MAT_EA'
    valid_material_types = {'FERT', 'HALB', 'ROH'}
    mara = mara_df.drop_duplicates(subset=[mat_col], keep='first')

    for idx, row in mara.iterrows():
        material = _norm_str(row.get(mat_col))
        material_type = _norm_str(row.get(mt_col)).upper()
        if material_type not in valid_material_types:
            continue

        value = _norm_str(row.get(ref_col))

        if value == '':
            error = f"Mandatory field blank (should be '{expected}')"
        elif value.upper() != expected:
            error = f"Invalid value: '{value}' (should be '{expected}')"
        else:
            continue

        resultados.append({
            'Line': idx + 2,
            'Field': 'Reference matl for packing',
            'Error': f"{error} for Material Type {material_type}",
            'Material': material,
            'Plant': '',
        })


def validar_maintenance_status_c_roh_mara(mara_df, resultados, sap_df=None):
    """
    MARA: 'Maintenance status' must contain the letter 'C'
    for Material Type ROH (Classification view maintained).
    Validation is material-level (MARA, not MARC).

    Exception (same one as Batch Management): the materials in
    BATCH_MANAGEMENT_HALB_LIKE_MATERIALS, when ROH with Plant Type PP / COI
    in at least one plant, are treated as HALB -> they do NOT need the
    letter 'C' and are therefore never reported (nor as Missing View).
    Needs sap_df (Material / Material Type / Plant Type) to know the Plant Type.
    """
    if mara_df is None or mara_df.empty:
        return

    halb_like_materials = set()
    if (
        sap_df is not None
        and not sap_df.empty
        and {'Material', 'Material Type', 'Plant Type'}.issubset(sap_df.columns)
    ):
        for _, _r in sap_df.iterrows():
            _m = _norm_str(_r.get('Material'))
            if get_batch_management_effective_material_type(
                _m, _norm_str(_r.get('Material Type')), _norm_str(_r.get('Plant Type'))
            ) == 'HALB' and _norm_str(_r.get('Material Type')).upper() == 'ROH':
                halb_like_materials.add(_m)

    ms_col = _find_col_ci(mara_df, ['Maintenance status', 'Maintenance Status', 'PSTAT'])
    mat_col = _find_col_ci(mara_df, ['Material'])
    mt_col = _find_col_ci(mara_df, ['Material Type'])

    if ms_col is None or mat_col is None or mt_col is None:
        print("⚠️ MARA is missing 'Maintenance status' (or Material / Material Type).")
        return

    mara = mara_df.drop_duplicates(subset=[mat_col], keep='first')

    for idx, row in mara.iterrows():
        material_type = _norm_str(row.get(mt_col)).upper()
        if material_type != 'ROH':
            continue

        # Exception: listed ROH with PP/COI are treated as HALB -> no 'C' needed
        if _norm_str(row.get(mat_col)) in halb_like_materials:
            continue

        # Exception: materials that never need view 'C'
        if _norm_str(row.get(mat_col)).upper() in MARA_MAINT_STATUS_C_EXEMPT_MATERIALS:
            continue

        status = _norm_str(row.get(ms_col)).upper()
        if 'C' in status:
            continue

        resultados.append({
            'Line': idx + 2,
            'Field': MARA_MAINT_STATUS_C_FIELD,
            'Error': (
                "Maintenance status must contain the letter 'C' for Material Type ROH "
                f"(current: '{status or 'blank'}')"
            ),
            'Material': _norm_str(row.get(mat_col)),
            'Plant': '',
        })


def validar_special_procurement_type_pt02_transfer(marc_df, zplm_marc_df, resultados):
    """
    MARC: Special Procurement Type — PT02 with transfer.

    Rule:
      - PT02 materials with 'Material Transfer' marked in ZPLM MARC
        must have Special Procurement Type = 'Z2'.

    Notes:
      - PT02 without transfer is not validated by this rule.
      - PT01 is not validated by this rule.
      - Scope follows marc_df (already restricted to FERT/HALB/ROH).
    """
    col = 'Special Procurement Type'

    if marc_df is None or marc_df.empty or zplm_marc_df is None or zplm_marc_df.empty:
        return
    if not {'Material', 'Plant', col}.issubset(marc_df.columns):
        return

    transfer_map = build_material_transfer_map(zplm_marc_df)
    if not transfer_map:
        return

    for idx, row in marc_df.iterrows():
        plant = _norm_str(row.get('Plant')).upper()
        if plant != 'PT02':
            continue

        material = _norm_str(row.get('Material'))
        if not transfer_map.get((material, plant), False):
            continue

        value = _norm_str(row.get(col)).upper()

        if value == '':
            error = (
                "Mandatory field blank "
                "(PT02 with transfer requires Special Procurement Type = 'Z2')"
            )
        elif value != 'Z2':
            error = (
                "PT02 with transfer requires Special Procurement Type = 'Z2' "
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
         - ROH with Plant Type PP or COI:
             Batch Management(Plant) must be active.
             Accepted values: '1' or 'X'.

         - FERT and HALB (any Plant Type), and ROH with other Plant Types:
             Batch Management(Plant) must not be active.
             Therefore, it must not contain '1' or 'X'.

         - Exception: BATCH_MANAGEMENT_HALB_LIKE_MATERIALS, when ROH with
           PP/COI, are treated as HALB (must NOT be active).

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

        # Material Type exception: listed materials that are ROH with
        # Plant Type PP/COI are treated as HALB (Batch Management only).
        material_type = get_batch_management_effective_material_type(
            material,
            material_type,
            plant_type,
        )

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
        elif material_type == 'ROH' and plant_type in {'PP', 'COI'}:
            if batch_management_plant == '':
                erros.append(
                    "Mandatory field blank "
                    f"(ROH with Plant Type '{plant_type}' requires "
                    "Batch Management(Plant) = '1' or 'X')"
                )

            elif batch_management_plant not in active_values:
                erros.append(
                    f"For Material Type 'ROH' with Plant Type '{plant_type}', "
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
                    f"for Material Type '{material_type}' "
                    f"with Plant Type '{plant_type}' "
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
