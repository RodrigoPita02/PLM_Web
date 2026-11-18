"""Helpers do SSA."""

import pandas as pd

from engines.material_rules.fields import _prepare_mvke
from engines.ssa.constants import SSA_REJECTION_EXCLUDED_ORDER_TYPES
from engines.ssa.state import _SSA_ACTIVE_DOC_TYPE_MAP

def _ssa_norm(x):
    """Normalise a cell value to stripped string; NaN/None -> ''."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    return str(x).strip()


def _ssa_pick(df, candidates):
    """Return first matching column name (case-insensitive)."""
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None


def _ssa_load(df):
    """Normalise column names: strip whitespace."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _ssa_to_float(x):
    """Convert SAP/Excel numeric-like value to float; blank/invalid -> None."""
    s = _ssa_norm(x)
    if s == '':
        return None
    try:
        return float(str(s).replace(',', '.'))
    except Exception:
        return None


def _ssa_is_zero_or_blank(x):
    """True if value is blank, None, NaN or numeric zero."""
    s = _ssa_norm(x)
    if s == '':
        return True
    try:
        return float(str(s).replace(',', '.')) == 0.0
    except Exception:
        return False


def _ssa_float_equal(a, b, tol=1e-6):
    """Numeric comparison with tolerance."""
    fa = _ssa_to_float(a)
    fb = _ssa_to_float(b)
    if fa is None or fb is None:
        return False
    return abs(fa - fb) <= tol


def _ssa_first_nonempty(*values):
    """Return first non-empty normalized value."""
    for v in values:
        s = _ssa_norm(v)
        if s != '':
            return s
    return ''


def _ssa_get_vbpa_value(vbpa_df, doc, item, partner_func, field_name="Unloading Point"):
    """
    Returns the first matching VBPA value for a given document/item/partner function.
    Prefers exact item match; if not found, returns header-level item=0/blank.
    """
    if vbpa_df is None or vbpa_df.empty:
        return ''

    vbpa = _ssa_load(vbpa_df)

    doc_col = _ssa_pick(vbpa, ["SD Document", "VBELN"])
    item_col = _ssa_pick(vbpa, ["Item (SD)", "POSNR"])
    pf_col   = _ssa_pick(vbpa, ["Partner Function", "PARVW"])
    val_col  = _ssa_pick(vbpa, [field_name])

    if not all([doc_col, item_col, pf_col, val_col]):
        return ''

    df = vbpa[vbpa[doc_col].astype(str).str.strip() == str(doc).strip()].copy()
    if df.empty:
        return ''

    df = df[df[pf_col].astype(str).str.strip().upper() == str(partner_func).strip().upper()]
    if df.empty:
        return ''

    item_str = str(item).strip()
    exact = df[df[item_col].astype(str).str.strip() == item_str]
    if not exact.empty:
        return _ssa_first_nonempty(*exact[val_col].tolist())

    hdr = df[df[item_col].astype(str).str.strip().isin(['0', '', '000000', '0000000000'])]
    if not hdr.empty:
        return _ssa_first_nonempty(*hdr[val_col].tolist())

    return ''


def _ssa_doc_type_in(doc, allowed_types):
    """
    True if the Sales Document's Sales Document Type (AUART) is a member of
    allowed_types (e.g. SSA_ORDER_TYPES_ZLZ_GROUP).

    Rule applicability depends solely on AUART — the Sales Document number
    itself is never inspected, so no rule depends on Sales Document number
    ranges.
    """
    doc_type = _SSA_ACTIVE_DOC_TYPE_MAP.get(_ssa_norm(doc), '').upper()
    return bool(doc_type) and doc_type in allowed_types


def _ssa_filter_docs(df, allowed_docs, doc_candidates):
    """Return a copy containing only allowed Sales Documents."""
    if df is None or df.empty:
        return df
    work = _ssa_load(df)
    doc_col = _ssa_pick(work, doc_candidates)
    if not doc_col:
        return work
    return work[
        work[doc_col].astype(str).str.strip().isin(allowed_docs)
    ].copy()


def _ssa_build_doc_plant_map(vbap_df):
    """
    Builds a Sales Document -> Plant lookup from VBAP, keeping the first
    non-blank Plant found for each Sales Document. Used to backfill the
    'Plant' column in the final error report for rules that do not report
    Plant themselves (e.g. header-level checks).
    """
    if vbap_df is None or vbap_df.empty:
        return {}

    work = _ssa_load(vbap_df)
    doc_col = _ssa_pick(work, ["Sales Document", "VBELN"])
    plt_col = _ssa_pick(work, ["Plant", "WERKS"])

    if doc_col is None or plt_col is None:
        return {}

    doc_plant_map = {}
    for doc, plant in zip(
        work[doc_col].map(_ssa_norm),
        work[plt_col].map(_ssa_norm),
    ):
        if doc == "" or plant == "":
            continue
        doc_plant_map.setdefault(doc, plant)

    return doc_plant_map


def _ssa_fill_missing_plant(errors_df, doc_plant_map):
    """
    Fills blank 'Plant' values in errors_df using the Sales Document ->
    Plant lookup built from VBAP via _ssa_build_doc_plant_map. Rows whose
    Sales Document has no match in VBAP, or that already carry a Plant
    value, are left unchanged.
    """
    if errors_df is None or errors_df.empty or not doc_plant_map:
        return errors_df
    if "Plant" not in errors_df.columns or "Sales Document" not in errors_df.columns:
        return errors_df

    errors_df = errors_df.copy()
    blank_mask = errors_df["Plant"].map(_ssa_norm) == ""
    errors_df.loc[blank_mask, "Plant"] = (
        errors_df.loc[blank_mask, "Sales Document"]
        .map(_ssa_norm)
        .map(doc_plant_map)
        .fillna("")
    )
    return errors_df


def _ssa_filter_non_s_items(vbap_df):
    """Exclude materials whose number starts with 'S-' (case-insensitive)."""
    if vbap_df is None or vbap_df.empty:
        return vbap_df
    work = _ssa_load(vbap_df)
    mat_col = _ssa_pick(work, ["Material", "MATNR"])
    if not mat_col:
        return work
    is_s = work[mat_col].apply(lambda value: _ssa_norm(value).upper().startswith('S-'))
    return work[~is_s].copy()


def _ssa_filter_vbpa_non_s_items(vbpa_df, vbap_df):
    """
    Exclude VBPA item rows that belong to S-* VBAP items.

    Header rows are retained only for documents that still have at least one
    non-S item. This allows header/item comparisons for mixed documents while
    excluding documents made exclusively of S-* materials.
    """
    if vbpa_df is None or vbpa_df.empty or vbap_df is None or vbap_df.empty:
        return vbpa_df

    vbpa = _ssa_load(vbpa_df)
    vbap = _ssa_load(vbap_df)
    vp_doc = _ssa_pick(vbpa, ["SD Document", "Sales Document", "VBELN"])
    vp_item = _ssa_pick(vbpa, ["Item (SD)", "Sales Document Item", "POSNR"])
    ap_doc = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    ap_item = _ssa_pick(vbap, ["Sales Document Item", "Item (SD)", "POSNR"])
    ap_mat = _ssa_pick(vbap, ["Material", "MATNR"])
    if not all([vp_doc, vp_item, ap_doc, ap_item, ap_mat]):
        return vbpa

    non_s = vbap[~vbap[ap_mat].apply(
        lambda value: _ssa_norm(value).upper().startswith('S-')
    )].copy()
    allowed_docs = set(non_s[ap_doc].astype(str).str.strip())
    allowed_items = set(zip(
        non_s[ap_doc].astype(str).str.strip(),
        non_s[ap_item].astype(str).str.strip(),
    ))
    header_values = {'0', '', '000000', '0000000000'}

    keep = []
    for _, row in vbpa.iterrows():
        doc = _ssa_norm(row[vp_doc])
        item = _ssa_norm(row[vp_item])
        keep.append(doc in allowed_docs and (item in header_values or (doc, item) in allowed_items))
    return vbpa.loc[keep].copy()


def _ssa_doc_type_map(vbak_df):
    """Build a map: Sales Document -> Sales Document Type (AUART), from VBAK."""
    if vbak_df is None or vbak_df.empty:
        return {}
    vbak = _ssa_load(vbak_df)
    doc_col  = _ssa_pick(vbak, ["Sales Document", "VBELN"])
    type_col = _ssa_pick(vbak, ["Sales Document Type", "Order Type", "AUART"])
    if not doc_col or not type_col:
        return {}
    return dict(zip(
        vbak[doc_col].astype(str).str.strip(),
        vbak[type_col].astype(str).str.strip().str.upper()
    ))


def _ssa_add_order_type_column(df):
    """
    Insert an 'Order Type' column (Sales Document Type / AUART) right after
    'Sales Document', resolved via the active doc-type map built for this
    run (_SSA_ACTIVE_DOC_TYPE_MAP). Rows whose Sales Document is unknown or
    blank get an empty value. No-op if df is empty or lacks the column.
    """
    if df is None or df.empty or "Sales Document" not in df.columns:
        return df
    df = df.copy()
    df["Order Type"] = df["Sales Document"].apply(
        lambda doc: _SSA_ACTIVE_DOC_TYPE_MAP.get(_ssa_norm(doc), "")
    )
    cols = [c for c in df.columns if c != "Order Type"]
    insert_at = cols.index("Sales Document") + 1
    cols.insert(insert_at, "Order Type")
    return df[cols]


def _ssa_item_is_done(row, status_col, rej_col):
    """
    Return True when a VBAP item is 'done' for SSA validation purposes,
    and can therefore be skipped by every rule in the SSA report:
      - Overall Status = 'C' (complete), OR
      - Reason for Rejection (VBAP-ABGRU) is filled.
    Either condition on its own is enough to skip the item.
    """
    if status_col and _ssa_norm(row[status_col]).upper() == "C":
        return True
    if rej_col and _ssa_norm(row[rej_col]) != "":
        return True
    return False


def _ssa_fully_rejected_docs(vbap_df):
    """
    Return documents for which every VBAP item is 'done': Overall Status
    = 'C', or Reason for Rejection (VBAP-ABGRU) filled (see
    _ssa_item_is_done).
    """
    if vbap_df is None or vbap_df.empty:
        return set()
    vbap = _ssa_load(vbap_df)
    doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    status_col = _ssa_pick(vbap, ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"])
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])
    if not doc_col or not status_col:
        return set()
    closed = set()
    for doc, grp in vbap.groupby(vbap[doc_col].map(_ssa_norm)):
        if doc and grp.apply(lambda r: _ssa_item_is_done(r, status_col, rej_col), axis=1).all():
            closed.add(doc)
    return closed


def _ssa_rejected_item_keys(vbap_df, order_types=SSA_REJECTION_EXCLUDED_ORDER_TYPES):
    """
    Return {(sales_document, item_key)} for every VBAP item whose Reason for
    Rejection (VBAP-ABGRU) is filled and whose Sales Document Type is in
    order_types. Needs _SSA_ACTIVE_DOC_TYPE_MAP to be set already.
    """
    if vbap_df is None or vbap_df.empty:
        return set()
    vbap = _ssa_load(vbap_df)
    doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "Item (SD)", "POSNR"])
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])
    if not doc_col or not item_col or not rej_col:
        return set()
    keys = set()
    for doc, item, rej in zip(vbap[doc_col], vbap[item_col], vbap[rej_col]):
        doc = _ssa_norm(doc)
        if not doc or _ssa_norm(rej) == "":
            continue
        if _SSA_ACTIVE_DOC_TYPE_MAP.get(doc, "").upper() not in order_types:
            continue
        keys.add((doc, _ssa_item_key(item)))
    return keys


def _ssa_drop_rejected_items(frame, rejected_keys, order_types=SSA_REJECTION_EXCLUDED_ORDER_TYPES):
    """
    Central safety net: remove from an SSA rule result every row that points
    to a rejected item (VBAP-ABGRU filled) of the given order types.
    Header-level rows (blank Item) are kept.
    """
    if frame is None or frame.empty or not rejected_keys:
        return frame
    if not {"Order Type", "Sales Document", "Item"}.issubset(frame.columns):
        return frame
    is_type = frame["Order Type"].astype(str).str.strip().str.upper().isin(order_types)
    is_rej = [
        (_ssa_norm(d), _ssa_item_key(i)) in rejected_keys
        for d, i in zip(frame["Sales Document"], frame["Item"])
    ]
    return frame[~(is_type & pd.Series(is_rej, index=frame.index))].reset_index(drop=True)


def _ssa_build_obd_set(vbfa_df):
    """
    Build the set of Sales Documents/items for which an Outbound Delivery
    has already been created, using VBFA Sales Document Flow.

    Correct document-flow logic:
      - VBELV / POSNV identify the preceding Sales Document and item.
      - VBELN identifies the Subsequent Document.
      - VBTYP_N must be 'J', meaning the Subsequent Document is a Delivery.

    Returns (obd_map, ready):
      obd_map: {sales_document: set(items)}
      ready: True only when VBFA and all mandatory OBD-gating columns exist.

    If VBTYP_N is unavailable, ready=False. The engine must not infer that a
    preceding-category value of J represents a Sales Order -> Delivery flow.
    """
    if vbfa_df is None or vbfa_df.empty:
        return {}, False

    vbfa = _ssa_load(vbfa_df)

    prec_doc_col = _ssa_pick(
        vbfa,
        ["Preceding Document", "VBELV"],
    )
    prec_item_col = _ssa_pick(
        vbfa,
        ["Item (SD) Preceding Doc.", "Preceding Item", "POSNV"],
    )
    subs_doc_col = _ssa_pick(
        vbfa,
        ["Subsequent Document", "VBELN"],
    )
    subs_cat_col = _ssa_pick(
        vbfa,
        ["Subsequent Doc. Cat.", "Preceding Document Category", "VBTYP_N"],
    )

    if not all([prec_doc_col, subs_doc_col, subs_cat_col]):
        return {}, False

    selected = [prec_doc_col, subs_doc_col, subs_cat_col]
    if prec_item_col:
        selected.append(prec_item_col)
    selected = list(dict.fromkeys(selected))

    work = vbfa.loc[:, selected].copy()
    work["_ssa_prec_doc"] = work[prec_doc_col].map(_ssa_norm)
    work["_ssa_subs_doc"] = work[subs_doc_col].map(_ssa_norm)
    work["_ssa_subs_cat"] = work[subs_cat_col].map(
        lambda value: _ssa_norm(value).upper()
    )

    if prec_item_col:
        work["_ssa_prec_item"] = work[prec_item_col].map(_ssa_norm)
    else:
        work["_ssa_prec_item"] = ""

    work = work.loc[
        (work["_ssa_prec_doc"] != "")
        & (work["_ssa_subs_doc"] != "")
        & (work["_ssa_subs_cat"] == "J")
    ]

    if work.empty:
        return {}, True

    result = (
        work.groupby("_ssa_prec_doc", sort=False)["_ssa_prec_item"]
        .agg(set)
        .to_dict()
    )
    return result, True


def _ssa_build_s_material_mvke_map(mvke_df, material_list):
    """
    Build a Material -> Acct Assmt Grp Mat. (KTGRM) map from MVKE, restricted
    to material_list.

    A material may carry more than one MVKE entry (one per Sales
    Organization / Distribution Channel combination); the first non-blank
    AAG value found for that material is used as the expected value.
    Returns {} if mvke_df is missing/empty or lacks the required columns.
    """
    if mvke_df is None or mvke_df.empty:
        return {}

    mvke_norm, cols = _prepare_mvke(mvke_df, set(material_list))
    if mvke_norm is None or mvke_norm.empty:
        return {}

    result = {}
    for material, grp in mvke_norm.groupby('_mat'):
        value = _ssa_first_nonempty(*grp['_aag'].tolist())
        if value:
            result[material] = value
    return result


def _ssa_has_obd(obd_map, doc, item):
    """True if an OBD (outbound delivery) already exists for this document/item."""
    items = obd_map.get(_ssa_norm(doc))
    if items is None:
        return False
    if '' in items:
        return True
    return _ssa_norm(item) in items


def _ssa_item_key(value):
    """Normalize SAP item numbers so 10, 10.0 and 000010 match."""
    raw = _ssa_norm(value)
    if raw == "":
        return ""
    try:
        number = float(raw.replace(",", "."))
        if number.is_integer():
            return str(int(number))
    except (TypeError, ValueError):
        pass
    stripped = raw.lstrip("0")
    return stripped if stripped else "0"


def _ssa_build_active_demand_set(vbep_df):
    """
    Build {(Sales Document, Item)} for VBEP rows with Release Type = '1'.

    VBAP and VBEP are matched by VBELN + POSNR. Item numbers are normalized
    so Excel representations such as 10, 10.0 and 000010 resolve to the same
    key. A VBEP item is considered to have active demand when at least one
    matching row has Release Type/ABART equal to '1'.

    Returns (active_demand, ready):
      active_demand: set of normalized (document, item) pairs.
      ready: True only when VBEP and all required columns are available.
    """
    if vbep_df is None or vbep_df.empty:
        return set(), False

    vbep = _ssa_load(vbep_df)
    doc_col = _ssa_pick(vbep, ["Sales Document", "SD Document", "VBELN"])
    item_col = _ssa_pick(vbep, ["Sales Document Item", "Item (SD)", "POSNR"])
    release_col = _ssa_pick(vbep, ["Release Type", "ABART"])

    if not all([doc_col, item_col, release_col]):
        return set(), False

    active_demand = set()
    for _, row in vbep.iterrows():
        if _ssa_norm(row[release_col]).upper() != "1":
            continue
        doc = _ssa_norm(row[doc_col])
        item = _ssa_item_key(row[item_col])
        if doc and item:
            active_demand.add((doc, item))

    return active_demand, True
