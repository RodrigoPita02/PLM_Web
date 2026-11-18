"""Histórico de erros, tendências e métricas por owner."""

import os
import json
import pandas as pd
from datetime import datetime

from engines.assignments.store import _load_master
from engines.common.paths import ERROR_HISTORY_PATH

_ERROR_HISTORY_PATH = ERROR_HISTORY_PATH


def _load_error_history() -> dict:
    """Load error history file. Returns empty dict if missing."""
    if not os.path.exists(_ERROR_HISTORY_PATH):
        return {}
    try:
        with open(_ERROR_HISTORY_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_error_history(history: dict) -> None:
    """Persist error history file atomically."""
    os.makedirs(os.path.dirname(_ERROR_HISTORY_PATH), exist_ok=True)
    tmp = _ERROR_HISTORY_PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _ERROR_HISTORY_PATH)
    except Exception:
        pass


def track_error_count(errors_df: pd.DataFrame, run_date: str | None = None) -> None:
    """
    Records error count for each (Material, Plant, Field) combination.
    Determines status: "appeared" | "improving" | "stable" | "worsening" | "resolved" | "recurring"
    
    Should be called at end of validation run, BEFORE save_open_assignments().
    
    Args:
        errors_df: Enriched errors DataFrame with Material, Plant, Field columns
        run_date: ISO date string; defaults to today
    """
    if errors_df is None or errors_df.empty:
        return
    
    today_str = run_date or datetime.now().strftime("%Y-%m-%d")
    history = _load_error_history()
    
    # Count errors per (Material, Plant, Field)
    error_counts = {}
    for _, row in errors_df.iterrows():
        material = str(row.get("Material", "")).strip()
        plant = str(row.get("Plant", "")).strip().upper()
        field = str(row.get("Field", "")).strip()
        key = f"{material}||{plant}||{field}"
        error_counts[key] = error_counts.get(key, 0) + 1
    
    # Update history
    for key, current_count in error_counts.items():
        if key not in history:
            history[key] = {"runs": []}
        
        runs = history[key]["runs"]
        
        # Determine status
        if not runs:
            status = "appeared"
        else:
            last_run = runs[-1]
            prev_count = last_run.get("error_count", 0)
            
            if prev_count == 0:
                status = "recurring"
            elif current_count < prev_count:
                status = "improving"
            elif current_count > prev_count:
                status = "worsening"
            else:
                status = "stable"
        
        # Add run record (avoid duplicates on same date)
        if not runs or runs[-1].get("date") != today_str:
            runs.append({
                "date": today_str,
                "error_count": current_count,
                "status": status
            })
    
    # Remove keys that no longer have errors (they were resolved)
    for key in list(history.keys()):
        if key not in error_counts:
            # This field had errors before, but doesn't now
            runs = history[key]["runs"]
            if runs and runs[-1].get("error_count", 0) > 0:
                # Last run had errors; this run has none
                if not runs or runs[-1].get("date") != today_str:
                    runs.append({
                        "date": today_str,
                        "error_count": 0,
                        "status": "resolved"
                    })
    
    _save_error_history(history)


def get_error_trend(material: str, plant: str, field: str) -> dict:
    """
    Returns trend analysis for a specific (Material, Plant, Field).
    
    Returns dict with:
    - first_seen: ISO date of first error
    - last_seen: ISO date of last event
    - current_count: Current error count (0 if resolved)
    - peak_count: Highest count ever recorded
    - status: Current status (appeared/improving/stable/worsening/resolved/recurring)
    - trend: "improving" | "worsening" | "stable" | "volatile"
    - resolution_rate: Percentage solved from peak
    - history: List of all runs
    """
    history = _load_error_history()
    key = f"{material}||{plant}||{field}"
    
    if key not in history:
        return {
            "material": material,
            "plant": plant,
            "field": field,
            "first_seen": None,
            "last_seen": None,
            "current_count": 0,
            "peak_count": 0,
            "status": "never_tracked",
            "trend": "n/a",
            "resolution_rate": "0%",
            "history": []
        }
    
    runs = history[key]["runs"]
    if not runs:
        return {
            "material": material,
            "plant": plant,
            "field": field,
            "first_seen": None,
            "last_seen": None,
            "current_count": 0,
            "peak_count": 0,
            "status": "never_tracked",
            "trend": "n/a",
            "resolution_rate": "0%",
            "history": []
        }
    
    # Calculate metrics
    first_seen = runs[0]["date"]
    last_seen = runs[-1]["date"]
    current_count = runs[-1]["error_count"]
    peak_count = max(r["error_count"] for r in runs)
    status = runs[-1]["status"]
    
    # Determine trend
    if len(runs) < 2:
        trend = "stable"
    else:
        recent = runs[-3:] if len(runs) >= 3 else runs
        counts = [r["error_count"] for r in recent]
        if counts[-1] < counts[0]:
            trend = "improving"
        elif counts[-1] > counts[0]:
            trend = "worsening"
        else:
            trend = "stable"
        
        # Check for volatile (up/down/up or similar)
        if len(counts) >= 3:
            diffs = [counts[i+1] - counts[i] for i in range(len(counts)-1)]
            if any(d > 0 for d in diffs) and any(d < 0 for d in diffs):
                trend = "volatile"
    
    # Resolution rate
    if peak_count > 0:
        resolution_rate = f"{int(100 * (peak_count - current_count) / peak_count)}%"
    else:
        resolution_rate = "0%"
    
    return {
        "material": material,
        "plant": plant,
        "field": field,
        "first_seen": first_seen,
        "last_seen": last_seen,
        "current_count": current_count,
        "peak_count": peak_count,
        "status": status,
        "trend": trend,
        "resolution_rate": resolution_rate,
        "history": runs
    }


def get_owner_metrics(owner: str) -> dict:
    """
    Returns aggregated metrics for all fields assigned to an owner.
    """

    if not owner or not owner.strip():
        return {}

    master = _load_master()
    history = _load_error_history()

    # ─────────────────────────────────────────────
    # FIND ALL ASSIGNED FIELDS
    # ─────────────────────────────────────────────
    assigned_fields = set()

    for field, info in master.get("field_assignments", {}).items():
        if info.get("owner") == owner:
            assigned_fields.add((field, ""))

    for plant, fields in master.get("plant_overrides", {}).items():
        for field, info in fields.items():
            if info.get("owner") == owner:
                assigned_fields.add((field, plant))

    # ─────────────────────────────────────────────
    # METRICS
    # ─────────────────────────────────────────────
    by_field = []
    total_active = 0

    improving = worsening = stable = resolved = recurring = 0

    # ✅ PATCH — evitar duplicação
    seen_keys = set()

    for field, plant in assigned_fields:

        for key, hist_data in history.items():

            if key in seen_keys:
                continue

            parts = key.split("||")
            if len(parts) != 3:
                continue

            hist_material, hist_plant, hist_field = parts

            # Match field e plant
            if hist_field != field:
                continue

            if plant and hist_plant != plant:
                continue

            runs = hist_data.get("runs", [])
            if not runs:
                continue

            seen_keys.add(key)

            last_run = runs[-1]
            current_count = int(last_run.get("error_count", 0))
            status = last_run.get("status", "")

            # ✅ agregação
            total_active += current_count

            if status == "improving":
                improving += 1
            elif status == "worsening":
                worsening += 1
            elif status == "stable":
                stable += 1
            elif status == "resolved":
                resolved += 1
            elif status == "recurring":
                recurring += 1

            by_field.append({
                "field": hist_field,
                "material": hist_material,
                "plant": hist_plant,
                "current_count": current_count,
                "peak_count": max(r.get("error_count", 0) for r in runs),
                "status": status,
                "first_seen": runs[0].get("date") if runs else None,
                "last_seen": runs[-1].get("date") if runs else None,
            })

    # ─────────────────────────────────────────────
    # OUTPUT
    # ─────────────────────────────────────────────
    return {
        "owner": owner,
        "total_active_errors": total_active,
        "fields_count": len(assigned_fields),
        "improving": improving,
        "worsening": worsening,
        "stable": stable,
        "resolved": resolved,
        "recurring": recurring,
        "by_field": sorted(by_field, key=lambda x: -x["current_count"])
    }


def get_error_summary_df(run_date: str | None = None) -> pd.DataFrame:
    """
    Returns a summary DataFrame of all tracked errors with trend info.
    Useful for dashboards.
    
    Columns: Material | Plant | Field | Current | Peak | Status | Trend | 
             First_Seen | Last_Seen | Days_Today
    """
    history = _load_error_history()
    today_str = run_date or datetime.now().strftime("%Y-%m-%d")
    today = datetime.strptime(today_str, "%Y-%m-%d")
    
    rows = []
    for key, hist_data in history.items():
        parts = key.split("||")
        if len(parts) == 3:
            material, plant, field = parts
            runs = hist_data.get("runs", [])
            if runs:
                last_run = runs[-1]
                first_run = runs[0]
                peak_count = max(r["error_count"] for r in runs)
                
                try:
                    days_active = (today - datetime.strptime(first_run["date"], "%Y-%m-%d")).days
                except Exception:
                    days_active = 0
                
                # Determine trend
                if len(runs) < 2:
                    trend = "stable"
                else:
                    recent = runs[-3:] if len(runs) >= 3 else runs
                    counts = [r["error_count"] for r in recent]
                    if counts[-1] < counts[0]:
                        trend = "improving"
                    elif counts[-1] > counts[0]:
                        trend = "worsening"
                    else:
                        trend = "stable"
                
                rows.append({
                    "Material": material,
                    "Plant": plant,
                    "Field": field,
                    "Current_Count": last_run["error_count"],
                    "Peak_Count": peak_count,
                    "Status": last_run["status"],
                    "Trend": trend,
                    "First_Seen": first_run["date"],
                    "Last_Seen": last_run["date"],
                    "Days_Active": days_active
                })
    
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def build_owner_dashboard_df(enriched_df: pd.DataFrame) -> pd.DataFrame:
    """
    Builds a dashboard-ready DataFrame with aggregated metrics per owner.
    
    Requires enriched_df from merge_assignments().
    
    Output columns:
    Owner | Active_Errors | Fields | Improving | Worsening | Stable |
    Overdue | Priority | Score
    """

    if enriched_df is None or enriched_df.empty:
        return pd.DataFrame()

    if "Owner" not in enriched_df.columns:
        return pd.DataFrame()

    owners = sorted([o for o in enriched_df["Owner"].dropna().unique() if str(o).strip()])

    rows = []

    for owner in owners:

        # 👉 usa a tua função robusta
        metrics = get_owner_metrics(owner)

        if not metrics:
            continue

        total = metrics.get("total_active_errors", 0)
        improving = metrics.get("improving", 0)
        worsening = metrics.get("worsening", 0)
        stable = metrics.get("stable", 0)

        # ✅ OVERDUE EXTRAÍDO DO DF (mais preciso)
        owner_df = enriched_df[enriched_df["Owner"] == owner]
        overdue = int((owner_df["Status"] == "Overdue").sum()) if "Status" in owner_df.columns else 0

        fields = metrics.get("fields_count", 0)

        # ✅ SCORING (simples e eficaz)
        score = (
            total
            + (worsening * 20)
            + (overdue * 30)
            - (improving * 10)
        )

        # ✅ PRIORIDADE
        if overdue > 5 or (worsening >= 3 and total > 100):
            priority = "🔴 High"
        elif overdue > 0 or total > 50:
            priority = "🟠 Medium"
        else:
            priority = "🟢 Low"

        rows.append({
            "Owner": owner,
            "Active_Errors": total,
            "Fields": fields,
            "Improving": improving,
            "Worsening": worsening,
            "Stable": stable,
            "Overdue": overdue,
            "Score": score,
            "Priority": priority,
        })

    df = pd.DataFrame(rows)

    if df.empty:
        return df

    # ✅ ordenação (mais crítico primeiro)
    df = df.sort_values(["Score", "Active_Errors"], ascending=[False, False])

    return df.reset_index(drop=True)
