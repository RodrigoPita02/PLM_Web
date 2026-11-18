"""Sistema de atribuição: 'atribuir uma vez, persiste para sempre'."""

import os
import json
import pandas as pd
from datetime import datetime

from engines.common.paths import ASSIGNMENTS_MASTER_PATH, OPEN_ITEMS_PATH
from engines.common.rolesviews import get_rolesviews_bytesio

_ASSIGNMENTS_PATH = ASSIGNMENTS_MASTER_PATH


def _load_master() -> dict:
    """Load the assignment master file. Returns empty structure if missing."""
    if not os.path.exists(_ASSIGNMENTS_PATH):
        return {"field_assignments": {}, "plant_overrides": {}}
    try:
        with open(_ASSIGNMENTS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("field_assignments", {})
        data.setdefault("plant_overrides", {})
        return data
    except Exception:
        return {"field_assignments": {}, "plant_overrides": {}}


def _save_master(master: dict) -> None:
    """Persist the assignment master file atomically."""
    os.makedirs(os.path.dirname(_ASSIGNMENTS_PATH), exist_ok=True)
    tmp = _ASSIGNMENTS_PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(master, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _ASSIGNMENTS_PATH)
    except Exception:
        pass


def _get_assignment_info(field: str, plant: str = "", master: dict | None = None) -> dict:
    """Returns the resolved assignment info for a field/plant combination."""
    if master is None:
        master = _load_master()
    field = (field or "").strip()
    plant = (plant or "").strip().upper()

    if not field:
        return {"owner": "", "team": "", "level": "Unassigned", "plant": ""}

    override = master.get("plant_overrides", {}).get(plant, {}).get(field)
    if isinstance(override, dict) and override.get("owner"):
        return {
            "owner": override["owner"],
            "team": override.get("team", ""),
            "level": "Plant Override",
            "plant": plant,
        }

    default = master.get("field_assignments", {}).get(field)
    if isinstance(default, dict) and default.get("owner"):
        return {
            "owner": default["owner"],
            "team": default.get("team", ""),
            "level": "Default",
            "plant": "",
        }

    return {"owner": "", "team": "", "level": "Unassigned", "plant": ""}


def get_owner(field: str, plant: str = "", master: dict | None = None) -> str:
    """
    Returns the owner for a (field, plant) combination.
    Plant-level override takes priority over field-level default.
    Returns "" if not assigned.
    """
    return _get_assignment_info(field, plant, master=master)["owner"]


def set_owner(field: str, owner: str, *,
              plant: str = "",
              team: str = "",
              master: dict | None = None) -> dict:

    if master is None:
        master = _load_master()

    field = (field or "").strip()
    plant = (plant or "").strip().upper()
    owner = (owner or "").strip()
    now   = datetime.now().strftime("%Y-%m-%d")

    if not field:
        raise ValueError("Field name is required to set an assignment.")

    # ✅ FIX: reload after remove
    if not owner:
        remove_owner(field, plant=plant)
        return _load_master()

    if plant:
        prev = master.setdefault("plant_overrides", {}).setdefault(plant, {}).get(field, {})

        # ✅ PATCH
        if team is not None:
            team_val = team
        else:
            team_val = prev.get("team", "")

        master["plant_overrides"][plant][field] = {
            "owner": owner,
            "team": team_val,
            "since": now
        }

    else:
        prev = master.setdefault("field_assignments", {}).get(field, {})

        # ✅ PATCH
        if team is not None:
            team_val = team
        else:
            team_val = prev.get("team", "")

        master["field_assignments"][field] = {
            "owner": owner,
            "team": team_val,
            "since": now
        }

    _save_master(master)
    return master


def remove_owner(field: str, *, plant: str = "") -> None:
    master = _load_master()
    field = (field or "").strip()
    plant = (plant or "").strip().upper()

    # ✅ FIX: proteção
    if not field:
        return

    if plant:
        plant_dict = master.get("plant_overrides", {}).get(plant, {})
        if isinstance(plant_dict, dict):
            plant_dict.pop(field, None)
            if not plant_dict:
                master.get("plant_overrides", {}).pop(plant, None)
    else:
        master.get("field_assignments", {}).pop(field, None)

    if not master.get("plant_overrides"):
        master["plant_overrides"] = {}

    _save_master(master)


def list_all_assignments() -> pd.DataFrame:
    """
    Returns a flat DataFrame of all current assignments:
    Field | Plant | Owner | Team | Since | Level
    """
    master = _load_master()
    rows = []
    for field, info in master.get("field_assignments", {}).items():
        rows.append({
            "Field":  field,
            "Plant":  "(all plants)",
            "Owner":  info.get("owner", ""),
            "Team":   info.get("team", ""),
            "Since":  info.get("since", ""),
            "Level":  "Default",
        })
    for plant, fields in master.get("plant_overrides", {}).items():
        for field, info in fields.items():
            rows.append({
                "Field":  field,
                "Plant":  plant,
                "Owner":  info.get("owner", ""),
                "Team":   info.get("team", ""),
                "Since":  info.get("since", ""),
                "Level":  "Plant Override",
            })
    return pd.DataFrame(rows, columns=["Field", "Plant", "Owner", "Team", "Since", "Level"])


def _build_field_to_role_map(roles_df: pd.DataFrame | None) -> dict:
    """
    Builds Field -> Role mapping from Roles DataFrame.
    Returns dict: {field_name: role_name}
    """
    if roles_df is None or roles_df.empty:
        return {}
    
    field_to_role = {}
    if 'Field' in roles_df.columns and 'Role' in roles_df.columns:
        roles_copy = roles_df.copy()
        roles_copy['Field'] = roles_copy['Field'].astype(str).str.strip()
        roles_copy['Role'] = roles_copy['Role'].astype(str).str.strip()
        # If multiple roles per field, take the first one
        field_to_role = roles_copy.drop_duplicates(subset=['Field']).set_index('Field')['Role'].to_dict()
    
    return field_to_role


def merge_assignments(errors_df: pd.DataFrame,
                      run_date: str | None = None,
                      roles_df: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    Enriches errors_df with assignment data from master.json.
    """

    if errors_df is None or errors_df.empty:
        return errors_df

    # ✅ AUTO-LOAD ROLES (PATCH)
    if roles_df is None:
        try:
            from engines.common.fast_loader import load_xlsx_fast

            _rv = load_xlsx_fast(get_rolesviews_bytesio(), sheet_name=None)
            roles_df = _rv.get("Roles")
        except Exception:
            roles_df = None

    master    = _load_master()
    today_str = run_date or datetime.now().strftime("%Y-%m-%d")
    today     = datetime.strptime(today_str, "%Y-%m-%d")

    field_to_role = _build_field_to_role_map(roles_df)

    # Load history
    open_path = OPEN_ITEMS_PATH
    open_map: dict[tuple, dict] = {}

    if os.path.exists(open_path):
        try:
            with open(open_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            for k, v in raw.items():
                parts = k.split("||")
                if len(parts) == 3:
                    open_map[tuple(parts)] = v
        except Exception:
            pass

    df = errors_df.copy()

    owners, teams, statuses, days_open_list, notes_list, levels = [], [], [], [], [], []

    for _, row in df.iterrows():
        field    = str(row.get("Field", "")).strip()
        plant    = str(row.get("Plant", "")).strip().upper()
        material = str(row.get("Material", "")).strip()

        assignment = _get_assignment_info(field, plant, master=master)

        owner = assignment["owner"]
        team  = assignment["team"]
        level = assignment["level"]

        # ✅ FINAL TEAM LOGIC
        assignment_team = (team or "").strip()
        role_team       = (field_to_role.get(field, "") or "").strip()

        if assignment_team:
            final_team = assignment_team
        elif role_team:
            final_team = role_team
        else:
            final_team = ""

        # Carry forward
        item_key = (material, plant, field)
        prev     = open_map.get(item_key, {})

        notes = prev.get("notes", "")
        first_seen_str = prev.get("first_seen", today_str)

        try:
            first_seen = datetime.strptime(first_seen_str, "%Y-%m-%d")
            d_open = (today - first_seen).days
        except Exception:
            d_open = 0

        OVERDUE_DAYS = 7

        if not owner:
            status = "Unassigned"
        elif d_open == 0:
            status = "New"
        elif d_open > OVERDUE_DAYS:
            status = "Overdue"
        else:
            status = "Open"

        owners.append(owner)
        teams.append(final_team)
        statuses.append(status)
        days_open_list.append(d_open)
        notes_list.append(notes)
        levels.append(level)

    df["Owner"]     = owners
    df["Team"]      = teams
    df["Level"]     = levels
    df["Status"]    = statuses
    df["Days_Open"] = days_open_list
    df["Notes"]     = notes_list

    return df


def save_open_assignments(enriched_df: pd.DataFrame) -> None:
    """
    Persists the current open assignment items to
    .history/assignments/open_items.json for carry-forward in the next run.

    Only saves rows that have an Owner (assigned items).
    Preserves first_seen date from existing entries.
    Saves: owner, team, status, days_open, notes, first_seen, last_seen
    """
    open_path = OPEN_ITEMS_PATH
    os.makedirs(os.path.dirname(open_path), exist_ok=True)

    # Load existing
    existing: dict = {}
    if os.path.exists(open_path):
        try:
            with open(open_path, "r", encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = {}

    today_str = datetime.now().strftime("%Y-%m-%d")
    updated: dict = {}

    for _, row in enriched_df.iterrows():
        owner = str(row.get("Owner", "")).strip()
        if not owner:
            continue
        material = str(row.get("Material", "")).strip()
        plant    = str(row.get("Plant",    "")).strip().upper()
        field    = str(row.get("Field",    "")).strip()
        key      = f"{material}||{plant}||{field}"

        prev_first_seen = existing.get(key, {}).get("first_seen", today_str)
        updated[key] = {
            "owner":      owner,
            "team":       str(row.get("Team",     "")).strip(),
            "notes":      str(row.get("Notes",     "")).strip(),
            "status":     str(row.get("Status",    "")).strip(),
            "days_open":  int(row.get("Days_Open", 0)),
            "first_seen": prev_first_seen,
            "last_seen":  today_str,
        }

    try:
        tmp = open_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(updated, f, ensure_ascii=False, indent=2)
        os.replace(tmp, open_path)
    except Exception:
        pass


def build_assignments_sheet(enriched_df: pd.DataFrame) -> pd.DataFrame:
    """
    Builds the Assignments sheet DataFrame for Excel export.
    Columns: Material | Plant | Field | Error | Owner | Team |
             Level | Status | Days_Open | Notes
    Sorted: Overdue first, then Open, then New, then Unassigned.
    """
    cols = ["Material", "Plant", "Field", "Error",
            "Owner", "Team", "Level", "Status", "Days_Open", "Notes"]

    present = [c for c in cols if c in enriched_df.columns]
    df = enriched_df[present].copy()

    # Add missing cols as empty
    for c in cols:
        if c not in df.columns:
            df[c] = ""

    status_order = {"Overdue": 0, "Open": 1, "New": 2, "Unassigned": 3}
    df["_sort"] = df["Status"].map(status_order).fillna(9)
    df = df.sort_values(["_sort", "Days_Open"], ascending=[True, False]).drop(columns=["_sort"])

    return df[cols].reset_index(drop=True)


def style_assignments_sheet(ws, df: pd.DataFrame) -> None:
    """
    Applies colour-coding to the Assignments sheet in an openpyxl workbook.
    Open = orange, New = blue, Unassigned = grey.
    """
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    STATUS_FILLS = {
        "Overdue":    PatternFill(start_color="FECACA", fill_type="solid"),  # 🔴
        "Open":       PatternFill(start_color="FFE5B4", fill_type="solid"),  # 🟠
        "New":        PatternFill(start_color="DBEAFE", fill_type="solid"),  # 🔵
        "Unassigned": PatternFill(start_color="F3F4F6", fill_type="solid"),  # ⚪
    }
    header_fill  = PatternFill(start_color="1F4E79", fill_type="solid")
    header_font  = Font(bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin  = Side(style="thin", color="B8CCE4")
    bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 26

    # Header
    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = header_align
        cell.border    = bdr

    # Find Status column index
    status_col = next(
        (i for i, cell in enumerate(ws[1], start=1) if str(cell.value) == "Status"),
        None
    )

    for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        status_val = ""
        if status_col:
            status_val = str(ws.cell(row=r_idx, column=status_col).value or "")
        fill = STATUS_FILLS.get(status_val, PatternFill())
        for cell in row:
            cell.border = bdr
            if fill:
                cell.fill = fill

    # Auto-width
    for col_idx, col in enumerate(ws.columns, start=1):
        max_len = max((len(str(cell.value or "")) for cell in col), default=8)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 50)
