"""Persistência de logs de execução (.history) e cofre de relatórios (reports/)."""

import io
import os
import json
import pandas as pd
from datetime import datetime

from engines.common.helpers import _safe_name

def _ts() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def build_sap_log_payload(*, user, plant_filter, mara_file, marc_file,
                           mvke_file, errors_df, summary_df, report_path=None):
    """Log payload for a SAP Validation run."""
    def _name(f, default):
        return getattr(f, "name", default) if f is not None else default

    total_errors    = int(errors_df.shape[0])            if errors_df is not None else 0
    total_fields    = int(errors_df['Field'].nunique())   if (errors_df is not None and 'Field' in errors_df.columns) else 0
    total_materials = int(errors_df['Material'].nunique()) if (errors_df is not None and 'Material' in errors_df.columns) else 0

    payload = {
        "run_type": "SAP_VALIDATION",
        "timestamp": _ts(),
        "user": user,
        "plant_filter": plant_filter or "ALL",
        "metrics": {"total_errors": total_errors, "fields_with_error": total_fields,
                    "materials_with_error": total_materials},
        "files": {"mara": _name(mara_file, "MARA.xlsx"),
                  "marc": _name(marc_file, "MARC.xlsx"),
                  "mvke": _name(mvke_file, "MVKE.xlsx")},
        "top_fields": summary_df.head(5).to_dict(orient="records") if summary_df is not None else [],
    }
    if report_path:
        payload["report_path"] = report_path
    return payload


def build_atl_log_payload(*, user, plant_scope, df_pt01, df_pt02):
    """Log payload for an ATL/PLM Workflow run."""
    def _tln_count(df):
        if df is None or df.empty or 'Timing Letter Number' not in df.columns:
            return 0
        return int(df['Timing Letter Number'].astype(str).str.strip().replace('', pd.NA).notna().sum())

    def _cn_count(df):
        if df is None or df.empty:
            return 0
        return int(df['Trigger ID'].nunique()) if 'Trigger ID' in df.columns else 0

    return {
        "run_type": "PLM_WORKFLOW_ATL",
        "timestamp": _ts(),
        "user": user,
        "plant_scope": plant_scope,
        "metrics": {
            "PT01_rows":       int(df_pt01.shape[0]) if df_pt01 is not None else 0,
            "PT02_rows":       int(df_pt02.shape[0]) if df_pt02 is not None else 0,
            "PT01_with_TLN":   _tln_count(df_pt01),
            "PT02_with_TLN":   _tln_count(df_pt02),
            "PT01_unique_CNs": _cn_count(df_pt01),
            "PT02_unique_CNs": _cn_count(df_pt02),
        },
    }


def build_progress_log_payload(*, user, current_name, previous_name, diff_dfs):
    """Log payload for a Progress Comparison run."""
    def _n(name):
        df = diff_dfs.get(name)
        return int(df.shape[0]) if df is not None else 0
    return {
        "run_type": "SAP_VALIDATION_PROGRESS",
        "timestamp": _ts(),
        "user": user,
        "files": {"current": current_name, "previous": previous_name},
        "metrics": {
            "new_errors":      _n("New Errors"),
            "resolved_errors": _n("Resolved Errors"),
            "changed_errors":  _n("Changed Errors"),
            "new_materials":   _n("New Materials"),
            "views_lost":      _n("Views Lost"),
        },
    }


def build_mrp_log_payload(*, user, plant_filter, mara_file, marc_file,
                            zplm_marc_file, errors_df, summary_df):
    """Log payload for an MRP Validation run."""
    def _name(f, default):
        return getattr(f, "name", default) if f is not None else default
    return {
        "run_type": "MRP_VALIDATION",
        "timestamp": _ts(),
        "user": user,
        "plant_filter": plant_filter or "ALL",
        "metrics": {
            "total_errors":       int(errors_df.shape[0])            if errors_df is not None else 0,
            "fields_with_error":  int(errors_df['Field'].nunique())   if (errors_df is not None and 'Field' in errors_df.columns) else 0,
            "materials_with_error": int(errors_df['Material'].nunique()) if (errors_df is not None and 'Material' in errors_df.columns) else 0,
        },
        "files": {"mara": _name(mara_file, "MARA.xlsx"),
                  "marc": _name(marc_file, "MARC.xlsx"),
                  "zplm_marc": _name(zplm_marc_file, "ZPLM MARC.xlsx")},
    }


def save_run_log(payload: dict, *, base_dir: str = ".history",
                  max_logs: int = 200) -> str | None:
    """
    Saves a log payload as a JSON file under .history/<run_type>/.

    FIX 12: Retention policy — keeps only the most recent `max_logs` files
    per run type. Older files are deleted automatically to prevent unbounded
    growth of the .history/ directory over time.
    """
    run_type = payload.get("run_type", "UNKNOWN")
    subdir   = {"SAP_VALIDATION": "sap", "PLM_WORKFLOW_ATL": "atl",
                 "SAP_VALIDATION_PROGRESS": "progress", "MRP_VALIDATION": "mrp",
                 "DATA_CLEANUPS": "dataclean",
                 "INDIRECT_MATERIALS": "indirect"}.get(run_type, "misc")

    out_dir = os.path.join(base_dir, subdir)
    _ensure_dir(out_dir)
    fname = f"{run_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    path  = os.path.join(out_dir, fname)

    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        return None

    # Retention: delete oldest files beyond max_logs
    try:
        all_logs = sorted(
            [os.path.join(out_dir, fn) for fn in os.listdir(out_dir)
             if fn.endswith(".json")],
            key=os.path.getmtime
        )
        for old_file in all_logs[:-max_logs]:
            try:
                os.remove(old_file)
            except Exception:
                pass
    except Exception:
        pass  # retention failure is non-critical

    return path


def list_run_logs(run_type: str, *, base_dir: str = ".history", limit: int = 25) -> list:
    """Lists the last `limit` log entries for a given run type."""
    subdir_map = {"sap": "sap", "atl": "atl", "progress": "progress",
                  "mrp": "mrp", "dataclean": "dataclean", "indirect": "indirect"}
    folder = os.path.join(base_dir, subdir_map.get(run_type, ""))
    if not os.path.isdir(folder):
        return []

    files = sorted(
        (os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(".json")),
        reverse=True
    )
    out = []
    for p in files[:limit]:
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
                data["_path"] = p
                out.append(data)
        except Exception:
            continue
    return out


def save_report_to_vault(excel_bytes: io.BytesIO, *, report_type: str = "sap",
                          plant_suffix: str = "ALL") -> str:
    """Saves the generated Excel to ./reports/<report_type>/<YYYY-MM>/ and returns the path."""
    _type_labels = {
        "sap":       "SAP_Validation",
        "atl":       "PLM_Workflow",
        "mrp":       "MRP_Validation",
        "dataclean": "Data_Cleanups",
        "progress":  "SAP_Progress",
        "indirect":  "Indirect_Materials",
    }
    label = _type_labels.get(report_type, report_type.upper())
    out_dir = os.path.join(".", "reports", report_type, datetime.now().strftime("%Y-%m"))
    os.makedirs(out_dir, exist_ok=True)
    fname = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{label}_{plant_suffix}.xlsx"
    fpath = os.path.join(out_dir, fname)
    excel_bytes.seek(0)
    with open(fpath, "wb") as f:
        f.write(excel_bytes.getvalue())
    return os.path.abspath(fpath)


def build_fp_log_payload(
    user: str,
    plant_filter: str | None,
    mara_file,
    marc_file,
    zplm_marc_file,
    zplm_mara_file,
    df_result: pd.DataFrame,
) -> dict:
    """
    Builds a structured log payload for an FP Status Validation run.

    Parameters
    ----------
    user          : username / identifier
    plant_filter  : 'PT01', 'PT02', or None (all plants)
    mara_file … zplm_mara_file : file objects (must have .name attribute, or None)
    df_result     : DataFrame returned by run_fp_status_validation
    """
    def _n(df):
        return 0 if (df is None or df.empty) else len(df)

    def _nuniq(df, col):
        return 0 if (df is None or df.empty or col not in df.columns) else int(df[col].nunique())

    def _count_status(df, status):
        if df is None or df.empty or "Status" not in df.columns:
            return 0
        return int((df["Status"] == status).sum())

    return {
        "run_type":     "FP_STATUS",
        "kind":         "fp",
        "timestamp":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "user":         user,
        "plant_filter": plant_filter or "ALL",
        "files": {
            "mara":      _safe_name(mara_file,      "MARA.xlsx"),
            "marc":      _safe_name(marc_file,      "MARC.xlsx"),
            "zplm_marc": _safe_name(zplm_marc_file, "ZPLM MARC.xlsx"),
            "zplm_mara": _safe_name(zplm_mara_file, "ZPLM MARA.xlsx"),
        },
        "metrics": {
            "materials": _nuniq(df_result, "Material"),
            "plants":    _nuniq(df_result, "Plant"),
            "outdated":  _count_status(df_result, "Outdated"),
            "correct":   _count_status(df_result, "Correct"),
            "need_verify": _count_status(df_result, "Need verify"),
            "total_rows":  _n(df_result),
        },
    }


def build_dc_log_payload(
    user: str,
    plant_filter: str | None,
    mara_file,
    marc_file,
    zplm_marc_file,
    zplm_mara_file,
    mbew_file,
    ekpo_file,
    mard_file,
    df_result: pd.DataFrame,
    kpis: dict,
) -> dict:
    """
    Builds a structured log payload for a Data Cleanups validation run.

    Parameters
    ----------
    user … mard_file : standard identifiers / file objects
    df_result        : combined results DataFrame (or None)
    kpis             : dict returned alongside df_result, e.g.
                       {"materials": int, "plants": int, "outdated": int, "correct": int}
    """
    return {
        "run_type":     "DATA_CLEANUPS",
        "timestamp":    datetime.now().isoformat(timespec="seconds"),
        "user":         user,
        "plant_filter": plant_filter or "ALL",
        "metrics": {
            "materials":   int(kpis.get("materials",   0)),
            "plants":      int(kpis.get("plants",      0)),
            "outdated":    int(kpis.get("outdated",    0)),
            "correct":     int(kpis.get("correct",     0)),
            "need_verify": int(kpis.get("need_verify", 0)),
        },
        "files": {
            "mara":      _safe_name(mara_file,      None),
            "marc":      _safe_name(marc_file,      None),
            "zplm_marc": _safe_name(zplm_marc_file, None),
            "zplm_mara": _safe_name(zplm_mara_file, None),
            "mbew":      _safe_name(mbew_file,      None),
            "ekpo":      _safe_name(ekpo_file,      None),
            "mard":      _safe_name(mard_file,      None),
        },
    }
