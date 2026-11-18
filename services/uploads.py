"""Classificação, validação e persistência dos ficheiros carregados (sem UI)."""

import io
import os
import pandas as pd
import re
import streamlit as st
from pathlib import Path
from datetime import datetime
from engines.common.run_logs import list_run_logs

# Regras: lista de (key, condição_lambda).
# Avaliadas em ordem — first match wins por ficheiro.
_CLASSIFY_RULES: list[tuple[str, object]] = [
    # ZPLM primeiro (mais específico)
    ("zplm_ta_ifs", lambda n: "zplm_ta_ifs" in n or ("zplm" in n and "ta" in n and "ifs" in n)),
    ("zplm_marc",   lambda n: "zplm" in n and "marc" in n),
    ("zplm_mara",   lambda n: "zplm" in n and "mara" in n),
    ("csr_item",    lambda n: "csr" in n and "item" in n),
    ("csr_head",    lambda n: "csr" in n and "head" in n),
    ("matplsta",    lambda n: "matplsta" in n),
    # Tabelas base
    ("mara",        lambda n: "mara" in n),
    ("marc",        lambda n: "marc" in n),
    ("mvke",        lambda n: "mvke" in n),
    ("mbew",        lambda n: "mbew" in n),
    ("ekpo",        lambda n: "ekpo" in n),
    ("mard",        lambda n: "mard" in n),
    ("marm",        lambda n: "marm" in n),
    # ATL
    ("auto",        lambda n: "automotive" in n),
    ("plm",         lambda n: "timingletter" in n or "timing_letter" in n),
    ("timing",      lambda n: "timing" in n and "workflow" not in n),
    ("workflow_task", lambda n: "workflowtask" in n or "workflow_task" in n or "c_workflowtask" in n),
    # Progress
    ("current",     lambda n: any(k in n for k in ["current", "novo", "atual"])),
    ("previous",    lambda n: any(k in n for k in ["previous", "anterior", "old"])),
]


ALL_KEYS = [k for k, _ in _CLASSIFY_RULES]


def classify_uploaded_files(files) -> dict:
    """
    Classifica uma lista de UploadedFile pelas regras centrais.
    Retorna dict com todas as keys; first-match por ficheiro.
    """
    mapping = {k: None for k in ALL_KEYS}
    if not files:
        return mapping
    for f in files:
        name = f.name.lower()
        for key, rule in _CLASSIFY_RULES:
            if rule(name):
                if mapping[key] is None:     # não sobrescreve se já foi preenchido
                    mapping[key] = f
                break
    return mapping


LABELS_PT = {
    "mara":      "MARA.xlsx",
    "marc":      "MARC.xlsx",
    "mvke":      "MVKE.xlsx",
    "zplm_marc": "ZPLM MARC.xlsx",
    "zplm_mara": "ZPLM MARA.xlsx",
    "zplm_ta_ifs": "ZPLM_TA_IFS.xlsx",
    "mbew":      "MBEW.xlsx",
    "ekpo":      "EKPO.xlsx",
    "mard":      "MARD.xlsx",
    "marm":      "MARM.xlsx",
    "auto":      "PLM Automotive Validation.xlsx",
    "plm":       "PLM TimingLetter Validation.xlsx",
    "timing":    "Timing Letters.xlsx",
    "workflow_task": "C_WORKFLOWTASK.xlsx",
    "current":   "Current Report (Progress)",
    "previous":  "Previous Report (Progress)",
    "matplsta":  "MATPLSTA.xlsx",
}


DISPLAY_TO_INTERNAL = {v.lower(): k for k, v in LABELS_PT.items()}


def display_to_internal_key(display_name: str):
    return DISPLAY_TO_INTERNAL.get((display_name or "").strip().lower())


# Colunas-chave esperadas por tipo de ficheiro (basta ter UMA das alternativas)
FILE_EXPECTED_COLS: dict[str, list[list[str]]] = {
    "mara":      [["Material", "Material Type"]],
    "marc":      [["Material", "Plant", "Maintenance status", "Cntry/Reg of Origin"]],
    "mvke":      [["Material", "Sales Organization"]],
    "zplm_marc": [["Material", "Plant", "Material Plant Type"]],
    "zplm_mara": [["Material", "Enterproj Number"]],
    "zplm_ta_ifs": [["Material"], ["MATNR"]],
    "mbew":      [["Material", "LIFO/FIFO-relevant"]],
    "ekpo":      [["Material"]],
    "mard":      [["Material", "Plant"]],
    "marm":      [["Material", "Alternative Unit of Measure", "Numerator", "Denominator"]],
    "auto":      [["Plant", "Approval Status", "Trigger ID"]],
    "plm":       [["Plant", "Approval Status", "Trigger ID"]],
    "timing":    [["Timing Letter Number", "Change Number"]],
    "workflow_task": [["Instance ID", "Work item text"]],
    "csr_item":  [["Material"]],
    "csr_head":  [["Plant"]],
    "matplsta":  [["Product Number", "Location", "Planning Status"]],
}


def _check_file_cols(file_obj, expected_col_sets: list[list[str]]) -> tuple[bool, str]:
    """
    Returns (ok, message).
    ok=True if at least one column set is fully present in the file.
    Uses only the first 5 rows for speed.
    """
    try:
        import io as _io
        if hasattr(file_obj, "getvalue"):
            bio = _io.BytesIO(file_obj.getvalue())
        elif hasattr(file_obj, "read"):
            file_obj.seek(0)
            bio = _io.BytesIO(file_obj.read())
        else:
            bio = file_obj
        bio.seek(0)
        df_head = pd.read_excel(bio, engine="openpyxl", nrows=5)
        actual = set(str(c).strip() for c in df_head.columns)
        for col_set in expected_col_sets:
            if all(col in actual for col in col_set):
                return True, ""
        # All sets failed — report what the best match was
        best = max(expected_col_sets, key=lambda s: sum(c in actual for c in s))
        missing = [c for c in best if c not in actual]
        return False, f"Expected columns not found: {', '.join(missing)}"
    except Exception as e:
        return False, f"Could not read file: {e}"


def validate_uploaded_files(persisted: dict, keys_to_check: list[str]) -> dict[str, str]:
    """
    Returns a dict key -> error_message for every file that fails column validation.
    Only checks files that are present in persisted.
    """
    errors = {}
    for key in keys_to_check:
        meta = persisted.get(key)
        if not meta:
            continue
        expected = FILE_EXPECTED_COLS.get(key)
        if not expected:
            continue
        bio = __import__("io").BytesIO(meta.get("bytes", b""))
        ok, msg = _check_file_cols(bio, expected)
        if not ok:
            errors[key] = msg
    return errors


# Keys obrigatórias por página
REQUIRED_KEYS: dict[str, list[str]] = {
    "📦 SAP Material Validation":         ["mara", "marc", "mvke", "zplm_marc", "zplm_mara", "marm", "mbew"],
    "📦 Indirect Materials Report":         ["mara", "marc", "mvke", "mbew"],
    "Automotive + Timing Letter (A & TL)":["auto", "plm", "timing"],
    "📊 SAP Validation Progress": [],   # gerido pelo modo Auto/Manual
    "🧹 Data Cleanups":                   ["mara", "marc", "zplm_marc", "zplm_mara", "zplm_ta_ifs", "mbew", "ekpo", "mard"],
}


def persist_uploaded_mapping(files_map: dict, state_key: str) -> dict:
    if not files_map:
        st.session_state[state_key] = {}
        return {}
    persisted = {}
    for k, f in files_map.items():
        if f is None:
            continue
        try:
            content = f.getvalue()
        except Exception:
            try:
                f.seek(0); content = f.read()
            except Exception:
                continue
        persisted[k] = {"name": getattr(f, "name", f"{k}.xlsx"), "bytes": content}
    st.session_state[state_key] = persisted
    return persisted


def rehydrate_persisted_mapping(state_key: str) -> dict:
    stored = st.session_state.get(state_key) or {}
    out = {}
    for k, meta in stored.items():
        if not meta:
            continue
        bio = io.BytesIO(meta.get("bytes", b""))
        try:
            bio.name = meta.get("name", f"{k}.xlsx")
        except Exception:
            pass
        out[k] = bio
    return out


_TS_PATTERN = re.compile(r'(\d{8}_\d{6})')


def extract_timestamp_from_name(filename: str) -> datetime | None:
    if not filename:
        return None
    m = _TS_PATTERN.search(filename)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y%m%d_%H%M%S")
    except Exception:
        return None


def pick_current_previous_from_uploads(files: list) -> tuple:
    diagnostics = []
    candidates = []
    for f in files or []:
        ts = extract_timestamp_from_name(getattr(f, "name", ""))
        if ts is None:
            diagnostics.append(f"⚠ No timestamp: {getattr(f, 'name', '')}")
        candidates.append((ts, f))

    with_ts    = sorted([(ts, f) for ts, f in candidates if ts], key=lambda x: x[0], reverse=True)
    without_ts = [(ts, f) for ts, f in candidates if not ts]

    ordered = [f for _, f in with_ts] + [f for _, f in without_ts]
    curr = ordered[0] if len(ordered) >= 1 else None
    prev = ordered[1] if len(ordered) >= 2 else None

    if len(with_ts) < 2:
        diagnostics.append("ℹ️ Fewer than 2 files with valid timestamp; using upload order.")

    return curr, prev, diagnostics


def pick_last_two_reports(plant_scope: str):
    import glob
    scope = (plant_scope or "ALL").upper()

    logs = list_run_logs("sap", limit=200) or []
    cand = []
    for log in logs:
        if str(log.get("plant_filter", "ALL")).upper() != scope:
            continue
        path = log.get("report_path")
        if path and os.path.exists(path):
            cand.append((log.get("timestamp", ""), os.path.abspath(path)))

    if len(cand) >= 2:
        cand.sort(key=lambda x: x[0], reverse=True)
        return cand[0][1], cand[1][1], len(cand)

    # Fallback: varrer vault
    base = Path(__file__).resolve().parent / "reports" / "sap"
    paths = []
    for m in (os.listdir(base) if base.exists() else []):
        paths.extend(base.glob(f"{m}/*SAP_Validation_Adjusted_Report_{scope}.xlsx"))
    paths = sorted([str(p) for p in paths], reverse=True)

    if len(paths) >= 2:
        return paths[0], paths[1], len(paths)
    if len(paths) == 1:
        return paths[0], None, 1
    return None, None, 0


def has_sheet(uploaded, sheet_name: str):
    try:
        xl = pd.ExcelFile(uploaded)
        return sheet_name in xl.sheet_names, xl.sheet_names
    except Exception as e:
        return False, [f"Error opening file: {e}"]
