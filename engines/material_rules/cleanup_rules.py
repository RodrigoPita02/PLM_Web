"""Regras C) — validações de Data Cleanups (FP status, LTB, stocks, ...)."""

import pandas as pd
import numpy as np

from engines.common.helpers import _normalize_code_as_str_basic
from engines.material_rules.sap_like import build_material_transfer_map

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


def validar_fp50_eop_expired(zplm_marc_df, marc_df):
    """
    FP materials (Material Plant Type = FP) with PSMS in {10,20,30,40,45,50}

    Rules:
    - All plants:
        EOP Date must not be empty
    - PT01 and all plants except PT02:
        EOP Date must be strictly in the future (EOP > today)
    - PT02 only:
        EOP Date must NOT be in the future (EOP <= today)

    Status Groups:
    - Group A: 10, 20, 30
    - Group B: 40, 45, 50
    """

    empty_cols = [
        "Material", "Plant", "Material Plant Type",
        "PSMS", "Status Group", "EOP Date", "Error"
    ]

    if zplm_marc_df is None or zplm_marc_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=empty_cols)

    # -------------------------
    # Setup & column detection
    # -------------------------
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

    m_mat    = pick(marc, ["MATERIAL", "MATNR"])
    m_plant  = pick(marc, ["PLANT", "WERKS"])
    m_psms   = pick(
        marc,
        ["PLANT-SPECIFIC MATERIAL STATUS",
         "PLANT SPECIFIC MATERIAL STATUS",
         "MTSTB"]
    )

    if any(c is None for c in [z_mat, z_plant, z_mptype, z_eop, m_mat, m_plant, m_psms]):
        return pd.DataFrame(columns=empty_cols)

    # -------------------------
    # Normalize ZPLM
    # -------------------------
    z = zplm.copy()
    z[z_mat]    = z[z_mat].astype(str).str.strip()
    z[z_plant]  = z[z_plant].astype(str).str.strip().str.upper()
    z[z_mptype] = z[z_mptype].astype(str).str.strip().str.upper()

    z_fp = z[z[z_mptype] == "FP"].copy()
    if z_fp.empty:
        return pd.DataFrame(columns=empty_cols)

    # -------------------------
    # Normalize MARC
    # -------------------------
    m = marc.copy()
    m[m_mat]   = m[m_mat].astype(str).str.strip()
    m[m_plant] = m[m_plant].astype(str).str.strip().str.upper()
    m[m_psms]  = m[m_psms].apply(
        lambda x: '' if pd.isna(x)
        else (str(int(float(str(x).strip())))
              if str(x).strip().replace('.', '').isdigit()
              else str(x).strip())
    )

    GROUP_A = {"10", "20", "30"}
    GROUP_B = {"40", "45", "50"}
    ALL_PS  = GROUP_A | GROUP_B

    m_sel = m[m[m_psms].isin(ALL_PS)].copy()
    if m_sel.empty:
        return pd.DataFrame(columns=empty_cols)

    # -------------------------
    # Intersect keys
    # -------------------------
    key_int = set(zip(z_fp[z_mat], z_fp[z_plant])) & set(zip(m_sel[m_mat], m_sel[m_plant]))
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

    # -------------------------
    # Validation
    # -------------------------
    errors = []

    for _, row in df_target.iterrows():
        psms = m_sel.loc[
            (m_sel[m_mat] == row[z_mat]) &
            (m_sel[m_plant] == row[z_plant]),
            m_psms
        ].iloc[0]

        status_group = (
            "Group A (10/20/30)" if psms in GROUP_A
            else "Group B (40/45/50)"
        )

        plant = row[z_plant]
        eop_raw = row.get(z_eop, None)
        eop = _dt(eop_raw)

        # ---- Empty mantém para todos ----
        if pd.isna(eop):
            errors.append({
                "Material": row[z_mat],
                "Plant": plant,
                "Material Plant Type": row[z_mptype],
                "PSMS": psms,
                "Status Group": status_group,
                "EOP Date": eop_raw,
                "Error": "EOP Date is empty for FP material with active/phase-out status",
            })
            continue

        # ---- PT02: futuro proibido ----
        if plant == "PT02":
            if eop > today:
                errors.append({
                    "Material": row[z_mat],
                    "Plant": plant,
                    "Material Plant Type": row[z_mptype],
                    "PSMS": psms,
                    "Status Group": status_group,
                    "EOP Date": eop_raw,
                    "Error": (
                        f"EOP Date is in the future (EOP {eop.date()} > today {today.date()}) "
                        f"— PT02 does not allow future dates"
                    ),
                })

        # ---- PT01 e restantes: regra original (expired) ----
        else:
            if eop <= today:
                errors.append({
                    "Material": row[z_mat],
                    "Plant": plant,
                    "Material Plant Type": row[z_mptype],
                    "PSMS": psms,
                    "Status Group": status_group,
                    "EOP Date": eop_raw,
                    "Error": (
                        f"EOP Date is expired (EOP {eop.date()} ≤ today {today.date()}) "
                        f"— FP material must have EOP Date in the future"
                    ),
                })

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
    LTB fields validation (Data Cleanups).

    Only materials meeting all the following conditions may have LTB values:
        - Material Plant Type = FP
        - Material Type in {FERT, HALB, ROH}
        - Plant-Specific Material Status in {55, 70, 99}

    Any other material with at least one populated LTB field generates
    a Type A error.

    LTB fields validated from ZPLM MARC:
        - LTB Qty - Quarter Calculation
        - LTB Original Qty Date
        - LTB Open Quantity
        - LTB Open Qty Update
        - LTB Original Qty
    """

    z = zplm_marc.copy()
    m = marc.copy()
    a = mara.copy()

    # ------------------------------------------------------------------
    # LTB fields from ZPLM MARC
    # ------------------------------------------------------------------
    ltb_fields = [
        'LTB Qty - Quarter Calculation',
        'LTB Original Qty Date',
        'LTB Open Quantity',
        'LTB Open Qty Update',
        'LTB Original Qty'
    ]

    ltb_quantity_fields = [
        'LTB Qty - Quarter Calculation',
        'LTB Open Quantity',
        'LTB Original Qty'
    ]

    ltb_date_fields = [
        'LTB Original Qty Date',
        'LTB Open Qty Update'
    ]

    # ------------------------------------------------------------------
    # Required columns
    # ------------------------------------------------------------------
    req_z = {
        'Material',
        'Plant',
        'Material Plant Type',
        *ltb_fields
    }

    req_m = {
        'Material',
        'Plant',
        'Plant-Specific Material Status'
    }

    req_a = {
        'Material',
        'Material Type'
    }

    missing = []

    if not req_z.issubset(z.columns):
        missing.append(
            f"ZPLM_MARC: {sorted(req_z - set(z.columns))}"
        )

    if not req_m.issubset(m.columns):
        missing.append(
            f"MARC: {sorted(req_m - set(m.columns))}"
        )

    if not req_a.issubset(a.columns):
        missing.append(
            f"MARA: {sorted(req_a - set(a.columns))}"
        )

    # ------------------------------------------------------------------
    # Output columns
    # ------------------------------------------------------------------
    cols = [
        'Material',
        'Plant',
        'Material Plant Type',
        'Material Type',
        'Plant-Specific Material Status',
        *ltb_fields,
        'Field',
        'Error'
    ]

    if missing:
        return pd.DataFrame(columns=cols)

    # ------------------------------------------------------------------
    # Prepare MARA
    # ------------------------------------------------------------------
    a_use = a[
        [
            'Material',
            'Material Type'
        ]
    ].copy()

    a_use['Material'] = (
        a_use['Material']
        .astype(str)
        .str.strip()
    )

    a_use['Material Type'] = (
        a_use['Material Type']
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ------------------------------------------------------------------
    # Prepare MARC
    # ------------------------------------------------------------------
    m_use = m[
        [
            'Material',
            'Plant',
            'Plant-Specific Material Status'
        ]
    ].copy()

    m_use['Material'] = (
        m_use['Material']
        .astype(str)
        .str.strip()
    )

    m_use['Plant'] = (
        m_use['Plant']
        .astype(str)
        .str.strip()
    )

    status_num = pd.to_numeric(
        m_use['Plant-Specific Material Status'],
        errors='coerce'
    )

    m_use['Plant-Specific Material Status'] = (
        status_num
        .astype('Int64')
        .astype(str)
        .str.replace('<NA>', '', regex=False)
        .str.strip()
    )

    # ------------------------------------------------------------------
    # Prepare ZPLM MARC
    # ------------------------------------------------------------------
    z['Material'] = (
        z['Material']
        .astype(str)
        .str.strip()
    )

    z['Plant'] = (
        z['Plant']
        .astype(str)
        .str.strip()
    )

    z['Material Plant Type'] = (
        z['Material Plant Type']
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ------------------------------------------------------------------
    # Merge reference data
    # ------------------------------------------------------------------
    z = (
        z
        .merge(
            a_use,
            on='Material',
            how='left'
        )
        .merge(
            m_use,
            on=['Material', 'Plant'],
            how='left'
        )
    )

    # ------------------------------------------------------------------
    # Authorization rules
    # ------------------------------------------------------------------
    valid_mtypes = {'FERT', 'HALB', 'ROH'}
    valid_status = {'55', '70', '99'}

    is_fp = (
        z['Material Plant Type']
        .astype(str)
        .str.strip()
        .str.upper()
        .eq('FP')
    )

    is_valid_mt = (
        z['Material Type']
        .astype(str)
        .str.strip()
        .str.upper()
        .isin(valid_mtypes)
    )

    is_valid_st = (
        z['Plant-Specific Material Status']
        .astype(str)
        .str.strip()
        .isin(valid_status)
    )

    authorized = (
        is_fp
        & is_valid_mt
        & is_valid_st
    )

    # ------------------------------------------------------------------
    # Check whether each LTB field contains a value
    # ------------------------------------------------------------------
    field_has_value = pd.DataFrame(
        index=z.index
    )

    # Quantity fields:
    # NaN, blank and zero are considered empty.
    for field in ltb_quantity_fields:
        raw_value = z[field]

        cleaned_value = (
            raw_value
            .astype(str)
            .str.strip()
        )

        numeric_value = pd.to_numeric(
            raw_value,
            errors='coerce'
        )

        is_empty = (
            raw_value.isna()
            | cleaned_value.isin(
                [
                    '',
                    '0',
                    '0.0',
                    'nan',
                    'None',
                    '<NA>'
                ]
            )
            | numeric_value.fillna(0).eq(0)
        )

        field_has_value[field] = ~is_empty

    # Date fields:
    # NaN, blank, zero and invalid dates are considered empty.
    for field in ltb_date_fields:
        raw_value = z[field]

        cleaned_value = (
            raw_value
            .astype(str)
            .str.strip()
        )

        date_value = pd.to_datetime(
            raw_value,
            errors='coerce'
        )

        is_empty = (
            raw_value.isna()
            | cleaned_value.isin(
                [
                    '',
                    '0',
                    '0.0',
                    'nan',
                    'None',
                    '<NA>'
                ]
            )
            | date_value.isna()
        )

        field_has_value[field] = ~is_empty

    # Material has an LTB value if at least one of the five fields
    # contains a valid non-empty value.
    has_any_ltb_value = field_has_value.any(axis=1)

    # ------------------------------------------------------------------
    # Type A error
    # ------------------------------------------------------------------
    mask_A = (
        ~authorized
        & has_any_ltb_value
    )

    output_fields = [
        'Material',
        'Plant',
        'Material Plant Type',
        'Material Type',
        'Plant-Specific Material Status',
        *ltb_fields
    ]

    df_A = z.loc[
        mask_A,
        output_fields
    ].copy()

    if df_A.empty:
        return pd.DataFrame(columns=cols)

    # ------------------------------------------------------------------
    # Identify exactly which LTB fields contain values
    # ------------------------------------------------------------------
    df_A['Field'] = field_has_value.loc[
        mask_A,
        ltb_fields
    ].apply(
        lambda row: ' / '.join(
            field
            for field, has_value in row.items()
            if has_value
        ),
        axis=1
    )

    df_A['Error'] = (
        "Material not authorized to have LTB values "
        "(requires FP + MT∈{FERT/HALB/ROH} "
        "+ Status∈{55,70,99})"
    )

    # ------------------------------------------------------------------
    # Final organization
    # ------------------------------------------------------------------
    df_A = df_A[cols]

    df_A.sort_values(
        by=['Plant', 'Material'],
        inplace=True,
        kind='stable'
    )

    df_A.reset_index(
        drop=True,
        inplace=True
    )

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


def validar_fp_existence_zplm_ta_ifs(zplm_marc_df, zplm_ta_ifs_df, marc_df):
    """Cleanup only: report active FP Material+Plant pairs missing in ZPLM_TA_IFS.

    ZPLM_TA_IFS is treated as a material-level interface table. The output remains
    plant-specific because FP classification and Part Status are plant-specific.
    PSMS 75/99 rows are expected to have been removed by the cleanup orchestrator.

    FP Material+Plant pairs marked for transfer in ZPLM MARC (field
    'Material Transfer' = X, see build_material_transfer_map) are skipped
    and not reported by this rule.
    """
    cols = ["Material", "Plant", "Material Plant Type", "Part Status", "Error", "Rule"]
    if any(df is None or df.empty for df in (zplm_marc_df, zplm_ta_ifs_df, marc_df)):
        return pd.DataFrame(columns=cols)

    def prep(df):
        out = df.copy()
        out.columns = [str(c).strip().upper() for c in out.columns]
        return out

    z = prep(zplm_marc_df)
    i = prep(zplm_ta_ifs_df)
    m = prep(marc_df)
    pick = lambda df, candidates: next((c for c in candidates if c in df.columns), None)
    zmat = pick(z, ["MATERIAL", "MATNR"])
    zplant = pick(z, ["PLANT", "WERKS"])
    ztype = pick(z, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])
    imat = pick(i, ["MATERIAL", "MATNR", "MATERIAL NUMBER", "PRODUCT", "PRODUCT NUMBER"])
    mmat = pick(m, ["MATERIAL", "MATNR"])
    mplant = pick(m, ["PLANT", "WERKS"])
    mstatus = pick(m, ["PLANT-SPECIFIC MATERIAL STATUS", "PLANT SPECIFIC MATERIAL STATUS", "MTSTB"])
    if not all([zmat, zplant, ztype, imat, mmat, mplant, mstatus]):
        return pd.DataFrame(columns=cols)

    z[zmat] = z[zmat].map(_normalize_code_as_str_basic).str.strip()
    z[zplant] = z[zplant].astype(str).str.strip().str.upper()
    z[ztype] = z[ztype].astype(str).str.strip().str.upper()
    i[imat] = i[imat].map(_normalize_code_as_str_basic).str.strip()
    m[mmat] = m[mmat].map(_normalize_code_as_str_basic).str.strip()
    m[mplant] = m[mplant].astype(str).str.strip().str.upper()
    m[mstatus] = m[mstatus].map(_normalize_code_as_str_basic)

    ifs_materials = set(i.loc[i[imat] != "", imat])
    status_map = (m.drop_duplicates([mmat, mplant])
                    .set_index([mmat, mplant])[mstatus].to_dict())
    transfer_map = build_material_transfer_map(zplm_marc_df)
    fp = z[z[ztype].eq("FP")].drop_duplicates([zmat, zplant])
    rows = []
    for _, row in fp.iterrows():
        material = row[zmat]
        plant = row[zplant]
        if transfer_map.get((material, plant), False):
            continue  # FP marked for transfer in ZPLM MARC — not counted
        if material and material not in ifs_materials:
            rows.append({
                "Material": material,
                "Plant": plant,
                "Material Plant Type": "FP",
                "Part Status": status_map.get((material, plant), ""),
                "Error": "FP material does not exist in ZPLM_TA_IFS",
                "Rule": "FP existence in ZPLM_TA_IFS",
            })
    return pd.DataFrame(rows, columns=cols)
