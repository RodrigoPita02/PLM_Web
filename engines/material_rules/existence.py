"""Regras D) — existência cruzada MARA↔ZPLM MARA e MARC↔ZPLM MARC."""

import pandas as pd


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
