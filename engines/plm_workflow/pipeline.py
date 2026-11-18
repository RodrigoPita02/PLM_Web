"""Pipeline: PLM Workflow (Automotive + Timing Letter / Change Cockpit)."""

import io
import pandas as pd
from datetime import datetime

from engines.common.excel_report import _apply_conditional_formatting_in_memory
from engines.common.fast_loader import load_xlsx_fast

def run_plm_workflow_combined(auto_file, plm_file, timing_file,
                              workflowtask_file=None):
    """
    PLM / ATL Workflow Report.
    Source of truth for groups + Technical Status: C_WORKFLOWTASK (optional).
    Returns: (excel_bytes, df_pt01, df_pt02)

    If workflowtask_file is provided, groups (X marks) and Approval Status
    are derived from C_WORKFLOWTASK (Technical Status = READY / SELECTED).
    If not provided, falls back to the original approver-column logic.
    """
    # ------------------------------------------------------------------
    # 1) LOAD FILES
    # ------------------------------------------------------------------
    df_auto   = load_xlsx_fast(auto_file,   sheet_name=0)
    df_plm    = load_xlsx_fast(plm_file,    sheet_name=0)
    df_timing = load_xlsx_fast(timing_file, sheet_name=0)

    if any(df is None or df.empty for df in [df_auto, df_plm, df_timing]):
        raise ValueError("ATL files missing: Automotive / TimingLetter / Timing Letters")

    for col in ['Plant', 'Trigger ID', 'Creation Date']:
        if col not in df_auto.columns:
            raise ValueError(f"Column '{col}' missing in Automotive file.")
        if col not in df_plm.columns:
            raise ValueError(f"Column '{col}' missing in TimingLetter PLM file.")

    # Enterproj Number and Project Name are optional — present when already
    # linked to a CN in PLM. If the columns don't exist, we create them empty.
    for ep_col in ['Enterproj Number', 'Project Name']:
        if ep_col not in df_auto.columns:
            df_auto[ep_col] = ""
        if ep_col not in df_plm.columns:
            df_plm[ep_col] = ""

    # ── Build EP map HERE — before any filter or dedup ────────────────────────
    # The Automotive file is the only source of Enterproj Number / Project Name.
    # We build the lookup now, using ALL rows (including non-INAPPROVAL ones),
    # so that every CN that has these fields populated can later enrich the
    # TimingLetter rows — regardless of Approval Status.
    # Key = normalised CN string (same format used throughout the rest of the fn)
    _auto_ep_map: dict[str, dict] = {}
    for _, _ar in df_auto.iterrows():
        _key = str(_ar.get('Trigger ID', '')).strip()
        try:
            _key = str(int(float(_key)))
        except Exception:
            pass
        if not _key or _key in ('', 'nan'):
            continue
        _ep = str(_ar.get('Enterproj Number', '')).strip()
        _pn = str(_ar.get('Project Name', '')).strip()
        # Normalise numeric Enterproj values (e.g. '1234.0' → '1234')
        try:
            _ep = str(int(float(_ep))) if _ep not in ('', 'nan') else ''
        except Exception:
            pass
        if _ep not in ('', 'nan') or _pn not in ('', 'nan'):
            _auto_ep_map[_key] = {
                'Enterproj Number': _ep if _ep not in ('', 'nan') else '',
                'Project Name':     _pn if _pn not in ('', 'nan') else '',
            }
    if 'Timing Letter Number' not in df_timing.columns or 'Change Number' not in df_timing.columns:
        raise ValueError("Timing Letters file must contain 'Timing Letter Number' and 'Change Number'.")

    # ------------------------------------------------------------------
    # 2) MAP TL ↔ CN (active Timing Letters only)
    # ------------------------------------------------------------------
    EXCLUDED_TL_STATUS = {'Complete', 'Rejected'}
    status_col = next(
        (c for c in df_timing.columns
         if c.strip().lower() in ('timing letter status', 'status')), None
    )
    active_timing = (
        df_timing[~df_timing[status_col].astype(str).str.strip().isin(EXCLUDED_TL_STATUS)].copy()
        if status_col else df_timing.copy()
    )

    def _norm_id(x):
        """Normalize numeric IDs — avoids scientific notation (9.9e+11 → 990000000120)."""
        s = str(x).strip()
        try:
            return str(int(float(s)))
        except Exception:
            return s

    map_tl_to_cn = dict(zip(
        active_timing['Timing Letter Number'].apply(_norm_id),
        active_timing['Change Number'].apply(_norm_id)
    ))
    map_cn_to_tl = dict(zip(
        active_timing['Change Number'].apply(_norm_id),
        active_timing['Timing Letter Number'].apply(_norm_id)
    ))

    # Normalise PLM Trigger ID → CN number, and populate Timing Letter Number
    df_plm['_orig_trigger'] = df_plm['Trigger ID'].astype(str).str.strip()
    df_plm['Trigger ID'] = df_plm['_orig_trigger'].apply(
        lambda x: map_tl_to_cn.get(x, x)
    )
    df_plm['Timing Letter Number'] = df_plm['_orig_trigger'].apply(
        lambda x: _norm_id(x) if x in map_tl_to_cn else _norm_id(map_cn_to_tl.get(x, ""))
    )
    df_plm.drop(columns=['_orig_trigger'], inplace=True)

    # ------------------------------------------------------------------
    # 3) FILTER — keep only INAPPROVAL / REJECTED from source files
    # ------------------------------------------------------------------
    if 'Approval Status' in df_auto.columns:
        df_auto = df_auto[df_auto['Approval Status'].isin(['INAPPROVAL', 'REJECTED'])].copy()
    if 'Approval Status' in df_plm.columns:
        df_plm  = df_plm[df_plm['Approval Status'].isin(['INAPPROVAL', 'REJECTED'])].copy()
    # Note: REJECTED rows are filtered OUT of the final report after merging

    # ------------------------------------------------------------------
    # 4) DEDUPE — keep most recent per Trigger ID
    # ------------------------------------------------------------------
    for df in (df_auto, df_plm):
        df['Creation Date'] = pd.to_datetime(df['Creation Date'], errors='coerce')
        df.sort_values('Creation Date', ascending=False, inplace=True)
        df.drop_duplicates(subset=['Trigger ID'], keep='first', inplace=True)

    # ------------------------------------------------------------------
    # 5) INITIALISE GROUP COLUMNS
    # ------------------------------------------------------------------
    auto_groups = ["Pricing", "Purchasing", "MPL", "Customs",
                   "Sales", "Manufacturing", "Costing", "MDS"]
    tl_groups   = ["MPL-TL", "MDS-TL"]

    for g in auto_groups:
        df_auto[g] = ""
    for g in tl_groups:
        df_plm[g] = ""

    # ------------------------------------------------------------------
    # 6) APPLY GROUP X MARKS
    # ------------------------------------------------------------------
    if workflowtask_file is not None:
        # ── NEW LOGIC: derive groups from C_WORKFLOWTASK ──────────────
        df_wf = load_xlsx_fast(workflowtask_file, sheet_name=0)
        if df_wf is None or df_wf.empty:
            raise ValueError("C_WORKFLOWTASK file is empty or could not be loaded.")

        df_wf.columns = [str(c).strip() for c in df_wf.columns]

        # The CN/TL key is in 'Instance ID.1' (last Instance ID column)
        # 'Instance ID' (col 4) is NaN for workflow rows — use Instance ID.1
        instance_col_key = 'Instance ID.1'
        if instance_col_key not in df_wf.columns:
            # Fallback: pick last column named 'Instance ID' (case-insensitive)
            candidates = [c for c in df_wf.columns if 'Instance ID' in c.lower()]
            instance_col_key = candidates[-1] if candidates else None
        if instance_col_key is None:
            raise ValueError("C_WORKFLOWTASK: no Instance ID column found.")

        # Normalise columns
        df_wf[instance_col_key]   = df_wf[instance_col_key].astype(str).str.strip()
        df_wf['Technical Status'] = df_wf['Technical Status'].astype(str).str.upper().str.strip()

        # Work item text column — SAP exports it as 'Work item text' (lowercase 'item')
        wit_col = next((c for c in df_wf.columns
                        if c.lower().replace(' ', '') == 'workitemtext'), None)
        if wit_col is None:
            raise ValueError("C_WORKFLOWTASK: no 'Work item text' column found.")
        df_wf[wit_col] = df_wf[wit_col].astype(str).str.strip()

        # Keep only READY / SELECTED rows
        df_wf = df_wf[df_wf['Technical Status'].isin({'READY', 'SELECTED'})].copy()

        # Build workflow_map: {CN_or_TL_id: {group: "X", _status: ...}}
        # Key = Instance ID.1 value (e.g. 'CN00141782' or '990000000051')
        # For each CN/TL, collect all groups marked X across multiple rows
        workflow_map = {}
        for _, row in df_wf.iterrows():
            iid    = row[instance_col_key]
            text   = row[wit_col]

            if iid in ('nan', '', 'NaN'):
                continue

            entry = workflow_map.setdefault(iid, {})

            # Skip non-approval rows — only READY/SELECTED with group info matter
            if 'ready to approve' not in text.lower():
                continue

            # Extract group from Work item text.
            # Known patterns:
            #   TL:        "MPL-TL 990000000051/CN00142474 Ready to Approve"
            #   Standard:  "Pricing-Ready to Approve CN00141782"
            #   RT suffix: "MDS-RT4 400000000257 Ready to Approve"
            #   Two-word:  "Master Data-Ready to Approve CN00142919" → MDS
            import re as _re
            _GROUP_ALIASES = {'Master Data': 'MDS'}
            _first_word = text.split()[0].strip()
            if _first_word.endswith('-TL'):
                # Timing Letter group — always identifiable by first word
                entry[_first_word] = "X"
            else:
                _m = _re.match(r'^(.+?)(?:-Ready|-RT\d+|\s+Ready)', text, _re.IGNORECASE)
                _group_raw = _m.group(1).strip() if _m else _first_word
                entry[_GROUP_ALIASES.get(_group_raw, _group_raw)] = "X"

        # Apply workflow_map to each dataframe
        def _tln_key(raw):
            """Normalize a raw TLN cell value (float or string) to a clean string key."""
            import math
            if raw is None or (isinstance(raw, float) and math.isnan(raw)):
                return ''
            if isinstance(raw, (float, __builtins__.__class__)):
                try:
                    return str(int(raw))
                except Exception:
                    pass
            s = str(raw).strip()
            if s in ('', 'nan', 'NaN'):
                return ''
            try:
                return str(int(float(s)))
            except Exception:
                return s

        # Reset index after dedup to avoid stale index issues with df.at[]
        df_auto = df_auto.reset_index(drop=True)
        df_plm  = df_plm.reset_index(drop=True)

        def apply_wf_map_auto(df):
            """Automotive: look up by Trigger ID (CN number)."""
            for idx, row in df.iterrows():
                key = str(row['Trigger ID']).strip()
                wf  = workflow_map.get(key)
                if not wf:
                    continue
                for col, val in wf.items():
                    if col in df.columns:
                        df.at[idx, col] = val
            return df

        def apply_wf_map_tl(df):
            """TimingLetter: look up by Timing Letter Number (990* key).
            Never use Trigger ID — a TL's Trigger ID is a CN that also exists
            in workflow_map as Automotive, which would apply wrong groups."""
            if 'Timing Letter Number' not in df.columns:
                return df
            for idx, row in df.iterrows():
                tln = _tln_key(row['Timing Letter Number'])
                wf  = workflow_map.get(tln) if tln else None
                if not wf:
                    continue
                for col, val in wf.items():
                    if col in df.columns:
                        df.at[idx, col] = val
            return df

        df_auto = apply_wf_map_auto(df_auto)
        df_plm  = apply_wf_map_tl(df_plm)

    else:
        # ── FALLBACK: original approver-column logic ───────────────────
        categories_auto    = ["Pricing", "Purchasing", "MPL", "Customs",
                               "Sales", "Manufacturing", "Costing", "MDS"]
        approver_cols_auto = ['Approver1'] + [f'approver{i}' for i in range(2, 8)] + ['Approver8']

        for idx, row in df_auto.iterrows():
            for col in approver_cols_auto:
                if col in df_auto.columns:
                    clean = str(row[col]).split("(")[0].strip() if pd.notna(row[col]) else ""
                    if clean in categories_auto:
                        df_auto.at[idx, clean] = "X"

        for idx, row in df_plm.iterrows():
            for col in [c for c in df_plm.columns if "approver" in c.lower()]:
                val = str(row[col]) if pd.notna(row.get(col)) else ""
                if val.startswith('MPL'): df_plm.at[idx, 'MPL-TL'] = "X"
                if val.startswith('MDS'): df_plm.at[idx, 'MDS-TL'] = "X"

    # ------------------------------------------------------------------
    # 7) MERGE + FINAL COLUMNS
    # ------------------------------------------------------------------
    # Build Creation Date map from C_WORKFLOWTASK (earliest 'Created' per CN/TL)
    _wf_date_map: dict = {}
    if workflowtask_file is not None:
        try:
            _df_wf_dates = load_xlsx_fast(workflowtask_file, sheet_name=0)
            _df_wf_dates.columns = [str(c).strip() for c in _df_wf_dates.columns]
            _iid_col = 'Instance ID.1' if 'Instance ID.1' in _df_wf_dates.columns else None
            if _iid_col and 'Created' in _df_wf_dates.columns:
                def _norm_iid(x):
                    if pd.isna(x): return ''
                    s = str(x).strip()
                    try: return str(int(float(s)))
                    except Exception: return s
                _df_wf_dates[_iid_col] = _df_wf_dates[_iid_col].apply(_norm_iid)
                _df_wf_dates = _df_wf_dates[_df_wf_dates[_iid_col] != '']
                _df_wf_dates['Created'] = pd.to_datetime(_df_wf_dates['Created'], errors='coerce')
                _wf_date_map = _df_wf_dates.groupby(_iid_col)['Created'].min().to_dict()
        except Exception:
            _wf_date_map = {}

    def _get_creation_date(trigger_id, tln, fallback):
        """Return earliest workflow creation date for a CN or TL."""
        key_cn  = str(trigger_id).strip() if trigger_id else ''
        key_tln = str(tln).strip() if tln and str(tln).strip() not in ('', 'nan', '0') else ''
        dt = _wf_date_map.get(key_cn) or (_wf_date_map.get(key_tln) if key_tln else None)
        return dt if dt is not None else fallback

    df_auto['Source'] = "Automotive"
    df_plm['Source']  = "TimingLetter"

    group_cols = ["MPL-TL", "MDS-TL",
                  "Pricing", "Purchasing", "MPL", "Customs",
                  "Sales", "Manufacturing", "Costing", "MDS"]

    final_cols = [
        "Trigger ID", "Timing Letter Number", "Enterproj Number", "Project Name", "Plant",
        "Creation Date", "Approval Status",
    ] + group_cols + ["Aging"]

    # Ensure all columns exist in both dataframes
    for col in final_cols + ["Source"]:
        if col not in df_auto.columns: df_auto[col] = ""
        if col not in df_plm.columns:  df_plm[col]  = ""

    df_final = pd.concat([df_auto, df_plm], ignore_index=True)

    # Propagate TLN from TimingLetter rows to Automotive rows with same Trigger ID
    tln_map = (
        df_final[df_final['Source'] == 'TimingLetter']
        [['Trigger ID', 'Timing Letter Number']]
        .dropna(subset=['Timing Letter Number'])
        .drop_duplicates(subset=['Trigger ID'])
        .set_index('Trigger ID')['Timing Letter Number']
        .to_dict()
    )
    mask_auto = df_final['Source'] == 'Automotive'
    df_final.loc[mask_auto, 'Timing Letter Number'] = df_final.loc[mask_auto].apply(
        lambda row: tln_map.get(str(row['Trigger ID']).strip(), row['Timing Letter Number'])
        if (pd.isna(row['Timing Letter Number']) or str(row['Timing Letter Number']).strip() == '')
        else row['Timing Letter Number'],
        axis=1
    )

    # Remove REJECTED rows
    df_final = df_final[df_final['Approval Status'] != 'REJECTED'].copy()

    # ── Merge duplicate rows (same Trigger ID in Automotive + TimingLetter) ──
    # Combine group X marks from both sources into one row per CN/TL.
    merged_rows = []
    for trigger_id, grp in df_final.groupby('Trigger ID', sort=False):
        if len(grp) == 1:
            row_dict = grp.iloc[0].to_dict()
            # If this is a TimingLetter-only row, enrich from the Automotive map
            if str(row_dict.get('Source', '')).strip() == 'TimingLetter':
                ep_data = _auto_ep_map.get(str(trigger_id).strip(), {})
                for ep_col in ('Enterproj Number', 'Project Name'):
                    if not row_dict.get(ep_col) or \
                       str(row_dict[ep_col]).strip() in ('', 'nan'):
                        fallback = ep_data.get(ep_col, '')
                        if fallback not in ('', 'nan'):
                            row_dict[ep_col] = fallback
            merged_rows.append(row_dict)
            continue

        # Start from the Automotive row if present (for Approval Status / Plant)
        auto_rows = grp[grp['Source'] == 'Automotive']
        base_row  = auto_rows.iloc[0].to_dict() if not auto_rows.empty else grp.iloc[0].to_dict()

        # Merge ALL group X marks across all rows
        for _, other in grp.iterrows():
            for gc in group_cols:
                if str(other.get(gc, '')).strip().upper() == 'X':
                    base_row[gc] = 'X'
            # Fill TLN if missing in base
            if not base_row.get('Timing Letter Number') or \
               str(base_row['Timing Letter Number']).strip() in ('', 'nan', '0'):
                tln_other = str(other.get('Timing Letter Number', '')).strip()
                if tln_other not in ('', 'nan', '0'):
                    base_row['Timing Letter Number'] = tln_other
            # Fill Enterproj Number / Project Name if missing in base
            for ep_col in ('Enterproj Number', 'Project Name'):
                if not base_row.get(ep_col) or \
                   str(base_row[ep_col]).strip() in ('', 'nan'):
                    ep_other = str(other.get(ep_col, '')).strip()
                    if ep_other not in ('', 'nan'):
                        base_row[ep_col] = ep_other

        merged_rows.append(base_row)

    df_final = pd.DataFrame(merged_rows)

    # ── Universal EP enrichment — apply map to every row ─────────────────────
    # After all merging, do a final pass to ensure every CN that has
    # Enterproj Number / Project Name in the Automotive file gets it populated,
    # regardless of which code path the row went through above.
    def _enrich_ep(row):
        trigger = str(row.get('Trigger ID', '')).strip()
        try:
            trigger = str(int(float(trigger)))
        except Exception:
            pass
        ep_data = _auto_ep_map.get(trigger, {})
        for ep_col in ('Enterproj Number', 'Project Name'):
            current = str(row.get(ep_col, '')).strip()
            if current in ('', 'nan') and ep_data.get(ep_col, '') not in ('', 'nan'):
                row[ep_col] = ep_data[ep_col]
        return row

    if not df_final.empty:
        df_final = df_final.apply(_enrich_ep, axis=1)

    # Apply Creation Date from C_WORKFLOWTASK (earliest workflow date per CN/TL)
    df_final['Creation Date'] = df_final.apply(
        lambda row: _get_creation_date(
            row.get('Trigger ID'),
            row.get('Timing Letter Number'),
            pd.to_datetime(row.get('Creation Date'), errors='coerce')
        ),
        axis=1
    )
    df_final['Creation Date'] = pd.to_datetime(df_final['Creation Date'], errors='coerce')

    # Compute Aging and sort by Aging descending
    today = datetime.now()
    df_final['Aging'] = (today - df_final['Creation Date']).dt.days
    df_final.sort_values('Aging', ascending=False, inplace=True)
    df_final['Creation Date'] = df_final['Creation Date'].dt.strftime('%Y-%m-%d')

    # Final column selection — no Source column
    df_final = df_final[[c for c in final_cols if c in df_final.columns]]

    df_pt01  = df_final[df_final['Plant'] == 'PT01'].copy()
    df_pt02  = df_final[df_final['Plant'] == 'PT02'].copy()

    # ------------------------------------------------------------------
    # 8) EXCEL OUTPUT
    # ------------------------------------------------------------------
    xlsx_bytes = io.BytesIO()
    with pd.ExcelWriter(xlsx_bytes, engine="openpyxl") as writer:
        df_pt01.to_excel(writer, sheet_name="PT01", index=False)
        df_pt02.to_excel(writer, sheet_name="PT02", index=False)
    xlsx_bytes.seek(0)
    xlsx_bytes = _apply_conditional_formatting_in_memory(xlsx_bytes)

    return xlsx_bytes, df_pt01, df_pt02
