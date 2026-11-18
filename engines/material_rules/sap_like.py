"""Construção do DataFrame 'SAP-like' (MARA+MARC+ZPLM) usado por Palmela e Tunísia."""

import re
import pandas as pd

from engines.common.helpers import _normalize_columns, normalize_code
from engines.material_rules.constants import RULE_MATERIAL_TYPE_OVERRIDES

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
