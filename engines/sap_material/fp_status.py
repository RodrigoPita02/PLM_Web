"""Validação de FP status (classificação de famílias)."""

import pandas as pd

from engines.common.fast_loader import load_xlsx_fast
from engines.common.helpers import _normalize_code_as_str_basic, _pick_col_upper

def run_fp_status_validation(mara_file, marc_file, zplm_marc_file,
                              zplm_mara_file=None, plant_filter=None):
    """
    FP Family Classification:
    Groups FP materials by family (same base, different last character).
    Classifies each as 'Correct', 'Outdated', or 'Need verify'.
    Returns: (df_fp, kpis)
    """
    mara = load_xlsx_fast(mara_file, sheet_name=0)
    marc = load_xlsx_fast(marc_file, sheet_name=0)
    zplm = load_xlsx_fast(zplm_marc_file, sheet_name=0)

    if mara is None or marc is None or zplm is None:
        raise ValueError("Missing required input files.")

    for df in [mara, marc, zplm]:
        df.columns = df.columns.astype(str).str.strip().str.upper()

    col_mat_mara  = _pick_col_upper(mara, ["MATERIAL", "MATNR"])
    col_mat_marc  = _pick_col_upper(marc, ["MATERIAL", "MATNR"])
    col_plant_marc = _pick_col_upper(marc, ["PLANT", "WERKS"])
    col_psms      = _pick_col_upper(marc, ["PLANT-SPECIFIC MATERIAL STATUS",
                                            "PLANT SPECIFIC MATERIAL STATUS", "MTSTB"])
    col_mat_zplm  = _pick_col_upper(zplm, ["MATERIAL", "MATNR"])
    col_plant_zplm = _pick_col_upper(zplm, ["PLANT", "WERKS"])
    col_mptype    = _pick_col_upper(zplm, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])

    pf = str(plant_filter).strip().upper() if plant_filter else None
    mara_set = set(mara[col_mat_mara].dropna().astype(str).str.strip())

    # Filter FP rows from ZPLM
    df_fp = zplm[zplm[col_mptype].astype(str).str.strip().str.upper() == "FP"][
        [col_mat_zplm, col_plant_zplm, col_mptype]
    ].copy()
    df_fp[col_mat_zplm]   = df_fp[col_mat_zplm].astype(str).str.strip()
    df_fp[col_plant_zplm] = df_fp[col_plant_zplm].astype(str).str.strip().str.upper()
    df_fp = df_fp[df_fp[col_mat_zplm].isin(mara_set)]
    if pf:
        df_fp = df_fp[df_fp[col_plant_zplm] == pf]

    # Filter MARC for allowed PSMS
    marc_small = marc[[col_mat_marc, col_plant_marc, col_psms]].copy()
    marc_small[col_mat_marc]   = marc_small[col_mat_marc].astype(str).str.strip()
    marc_small[col_plant_marc] = marc_small[col_plant_marc].astype(str).str.strip().str.upper()
    marc_small[col_psms]       = marc_small[col_psms].apply(_normalize_code_as_str_basic)
    marc_small = marc_small[marc_small[col_psms].isin({"40", "45", "50", "55"})]
    if pf:
        marc_small = marc_small[marc_small[col_plant_marc] == pf]

    merged = df_fp.merge(
        marc_small,
        left_on=[col_mat_zplm, col_plant_zplm],
        right_on=[col_mat_marc, col_plant_marc],
        how="inner"
    ).rename(columns={
        col_mat_zplm: "Material", col_plant_zplm: "Plant",
        col_mptype: "Material Plant Type", col_psms: "PSMS"
    })[["Material", "Plant", "Material Plant Type", "PSMS"]]

    merged["Material"] = merged["Material"].astype(str).str.strip()
    merged = merged[~merged["Material"].str.match(r"^\d+$")]

    PSMS_NAME = {"40": "PREPROD", "45": "PROD", "50": "BALOUT", "55": "SERVICE"}
    merged["PSMS Name"] = merged["PSMS"].map(PSMS_NAME)

    # Keep only families with 2+ members
    merged["_base"] = merged["Material"].str[:-1]
    base_counts = merged["_base"].value_counts()
    merged = merged[merged["_base"].isin(set(base_counts[base_counts >= 2].index))].copy()

    def classify_family(g):
        g = g.copy()
        g["suffix"] = g["Material"].str[-1]
        g = g.sort_values("suffix")
        psms_set = set(g["PSMS"])

        if psms_set == {"55"} and len(g) >= 2:
            g["Status"] = "Need verify"
            return g.drop(columns=["suffix"])

        priority = ["45", "40", "55", "50"]
        correct_psms = next((p for p in priority if p in psms_set), None)
        g["Status"] = "Correct"

        has40, has45, has55 = "40" in psms_set, "45" in psms_set, "55" in psms_set

        if correct_psms in ("45", "40") and has55:
            g.loc[g["PSMS"] == "55", "Status"] = "Need verify"
        if correct_psms == "55" and (has40 or has45):
            g.loc[g["PSMS"].isin(["40", "45"]), "Status"] = "Need verify"
        if correct_psms == "45" and has40:
            g.loc[g["PSMS"] == "40", "Status"] = "Need verify"
        if correct_psms == "40" and has45:
            g.loc[g["PSMS"] == "45", "Status"] = "Need verify"

        # Balout (PSMS=50): only the LAST one (highest suffix) is Correct.
        # All others in the same family with status 50 are Outdated.
        dup50 = g[g["PSMS"] == "50"]
        if len(dup50) >= 2:
            # Mark all as Outdated first, then restore Correct for the last one
            g.loc[dup50.index, "Status"] = "Outdated"
            g.loc[dup50.index[-1], "Status"] = "Correct"

        return g.drop(columns=["suffix"])

    df = merged.groupby(["Plant", "_base"], group_keys=False).apply(classify_family)
    df = df.drop(columns=["_base"])

    kpis = {} if df.empty else {
        "materials": df["Material"].nunique(),
        "plants":    df["Plant"].nunique(),
        "psms":      df["PSMS"].nunique(),
        "outdated":  int((df["Status"] == "Outdated").sum()),
        "correct":   int((df["Status"] == "Correct").sum()),
    }

    return df, kpis


def _filter_materials_by_base_duplicates(
    df: pd.DataFrame,
    material_col: str = "Material",
) -> pd.DataFrame:
    """
    Keeps only rows whose Material 'base' (all chars except the last) appears
    at least twice in the current scope.  Used to identify FP version families.
    The original Material value is preserved unchanged.
    """
    if df is None or df.empty or material_col not in df.columns:
        return df

    tmp = df.copy()
    tmp["_base"] = tmp[material_col].astype(str).str.strip().str[:-1]
    base_counts  = tmp["_base"].value_counts(dropna=False)
    bases_keep   = set(base_counts[base_counts >= 2].index)
    return tmp[tmp["_base"].isin(bases_keep)].drop(columns=["_base"])


def _classify_fp_versions(df: pd.DataFrame) -> pd.DataFrame:
    """
    Marks each FP material version within its family (all sharing the same base,
    i.e. all chars except the last).  Scope: (Plant, Base).

    Rules:
      - Sort by last character (A -> Z) within the group.
      - The lexicographically latest entry is 'Correct'; all others are 'Outdated'.

    Used as a lightweight alternative to the full PSMS-aware classify_family
    logic inside run_fp_status_validation.
    """
    if df.empty:
        df["Status"] = pd.Series(dtype=str)
        return df

    work = df.copy()
    work["Material"] = work["Material"].astype(str).str.strip()
    work = work[work["Material"].str.len() > 0].copy()

    work["Base"]     = work["Material"].str[:-1]
    work["LastChar"] = work["Material"].str[-1]
    work["Status"]   = "Outdated"

    for (_, _), grp in work.groupby(["Plant", "Base"]):
        if grp.empty:
            continue
        latest_idx = grp.sort_values("LastChar").index[-1]
        work.loc[latest_idx, "Status"] = "Correct"

    return work.drop(columns=["Base", "LastChar"])
