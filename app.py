# app.py
import streamlit as st
import pandas as pd
import re
import io
import os
import time
from pathlib import Path
from datetime import datetime

from utils.dashboard import (
    render_sap_run_kpis, render_sap_run_charts,
    render_atl_run_kpis, render_atl_run_charts,
    render_progress_kpis, render_progress_charts,
    render_mrp_run_kpis, render_mrp_run_charts,
    render_evolution_dashboard
)
from utils.fast_loader import load_xlsx_fast
from utils.identify import get_executor_identity
from utils.layout import render_footer, render_header, inject_theme_overrides
from sap_validation_engine import (
    run_full_validation,
    run_ssa_validation,
    run_plm_workflow_combined,
    build_sap_log_payload,
    build_atl_log_payload,
    build_progress_log_payload,
    save_run_log,
    list_run_logs,
    run_mrp_validation,
    build_mrp_log_payload,
    run_data_cleanups_validation,
    build_dc_log_payload,
    save_report_to_vault,
    _coerce_df_for_display,
    # Assignment system
    get_owner,
    set_owner,
    remove_owner,
    list_all_assignments,
    merge_assignments,
    save_open_assignments,
    build_assignments_sheet,
)

# ==========================================================
# CONFIGURAÇÃO INICIAL
# ==========================================================
st.set_page_config(page_title="Validation Hub", layout="wide", page_icon="🧪")
start_time = time.time()

# ==========================================================
# CSS GLOBAL — injetado UMA única vez
# ==========================================================
def inject_css(css_path: str):
    try:
        css_file = Path(__file__).resolve().parent / css_path
        if css_file.exists():
            st.markdown(f"<style>{css_file.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
        else:
            st.warning(f"⚠️ CSS não encontrado: {css_file}")
    except Exception as e:
        st.error(f"Erro ao carregar CSS: {e}")

inject_css("assets/styles.css")

# Header global
theme_name = render_header("🧪 Validation Hub", show_theme_picker=True)

# ==========================================================
# WIZARD UNIVERSAL — uma única função para todas as páginas
# ==========================================================
def render_wizard(current_step: int, steps: list[str]):
    """Wizard visual reutilizável. Recebe o passo atual e a lista de labels."""
    cols = st.columns(len(steps))
    for idx, (col, label) in enumerate(zip(cols, steps), start=1):
        with col:
            if current_step == idx:
                klass = "active"
            elif current_step > idx:
                klass = "done"
            else:
                klass = "pending"
            st.markdown(
                f"""
                <div class="wizard-step">
                    <div class="step-circle {klass}">{idx}</div>
                    <div class="step-label">{label}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    completion = (current_step - 1) / max(len(steps) - 1, 1)
    st.markdown(
        f'<div class="progress-line" style="width:{completion * 100:.0f}%"></div>',
        unsafe_allow_html=True,
    )

WIZARD_STEPS_DEFAULT  = ["Files", "Scope", "Validate", "Results"]
WIZARD_STEPS_PROGRESS = ["Mode", "Files", "Validate", "Results"]

# ==========================================================
# CLASSIFICADOR CENTRAL DE FICHEIROS
# ==========================================================
# Regras: lista de (key, condição_lambda).
# Avaliadas em ordem — first match wins por ficheiro.
_CLASSIFY_RULES: list[tuple[str, object]] = [
    # ZPLM primeiro (mais específico)
    ("zplm_marc",   lambda n: "zplm" in n and "marc" in n),
    ("zplm_mara",   lambda n: "zplm" in n and "mara" in n),
    ("csr_item",    lambda n: "csr" in n and "item" in n),
    ("csr_head",    lambda n: "csr" in n and "head" in n),
    # SAP export
    ("sap",         lambda n: "sapui5" in n or ("item" in n and "browse" in n)),
    # Tabelas base
    ("mara",        lambda n: "mara" in n),
    ("marc",        lambda n: "marc" in n),
    ("mvke",        lambda n: "mvke" in n),
    ("qmat",        lambda n: "qmat" in n),
    ("mbew",        lambda n: "mbew" in n),
    ("ekpo",        lambda n: "ekpo" in n),
    ("mard",        lambda n: "mard" in n),
    # ATL
    ("auto",        lambda n: "automotive" in n),
    ("plm",         lambda n: "timingletter" in n or "timing_letter" in n),
    ("timing",      lambda n: "timing" in n),
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

# ==========================================================
# LABELS PARA DISPLAY
# ==========================================================
LABELS_PT = {
    "sap":       "SAPUI5 Export (Item Browse Report.xlsx)",
    "mara":      "MARA.xlsx",
    "marc":      "MARC.xlsx",
    "mvke":      "MVKE.xlsx",
    "zplm_marc": "ZPLM MARC.xlsx",
    "zplm_mara": "ZPLM MARA.xlsx",
    "qmat":      "QMAT.xlsx",
    "mbew":      "MBEW.xlsx",
    "ekpo":      "EKPO.xlsx",
    "mard":      "MARD.xlsx",
    "auto":      "PLM Automotive Validation.xlsx",
    "plm":       "PLM TimingLetter Validation.xlsx",
    "timing":    "Timing Letters.xlsx",
    "current":   "Current Report (Progress)",
    "previous":  "Previous Report (Progress)",
}

DISPLAY_TO_INTERNAL = {v.lower(): k for k, v in LABELS_PT.items()}

def display_to_internal_key(display_name: str):
    return DISPLAY_TO_INTERNAL.get((display_name or "").strip().lower())

# Colunas-chave esperadas por tipo de ficheiro (basta ter UMA das alternativas)
FILE_EXPECTED_COLS: dict[str, list[list[str]]] = {
    "sap":       [["Material", "Plant", "Material Type", "MRP Type"]],
    "mara":      [["Material", "Material Type"]],
    "marc":      [["Material", "Plant", "Maintenance status"]],
    "mvke":      [["Material", "Sales Organization"]],
    "zplm_marc": [["Material", "Plant", "Material Plant Type"]],
    "zplm_mara": [["Material", "Enterproj Number"]],
    "qmat":      [["Material", "Plant"]],
    "mbew":      [["Material"]],
    "ekpo":      [["Material"]],
    "mard":      [["Material", "Plant"]],
    "auto":      [["Plant", "Approval Status", "Trigger ID"]],
    "plm":       [["Plant", "Approval Status", "Trigger ID"]],
    "timing":    [["Timing Letter Number", "Change Number"]],
    "csr_item":  [["Material"]],
    "csr_head":  [["Plant"]],
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
    "📦 SAP Material Validation":         ["sap", "mara", "marc", "mvke", "zplm_marc", "zplm_mara", "qmat"],
    "🧮 MRP Validation":                  ["mara", "marc"],
    "Automotive + Timing Letter (A & TL)":["auto", "plm", "timing"],
    "📊 SAP Validation Progress": [],   # gerido pelo modo Auto/Manual
    "🧹 Data Cleanups":                   ["mara", "marc", "zplm_marc", "zplm_mara", "mbew", "ekpo", "mard"],
}

# ==========================================================
# PERSISTÊNCIA DE FICHEIROS
# ==========================================================
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

# ==========================================================
# HELPERS DE UPLOAD UI — expanders reutilizáveis
# ==========================================================
def show_file_status(required_keys: list[str], files: dict, persisted: dict,
                     optional_keys: list[str] | None = None):
    """
    Mostra um único expander com o status de cada ficheiro obrigatório:
    ✅ identificado no upload atual | 💾 guardado na sessão | ❌ em falta

    Se optional_keys for fornecido, mostra também uma secção de opcionais
    com ⬜ (não carregado) ou ✅ (presente).

    Também mostra ficheiros não reconhecidos com sugestão de renomeação.
    """
    with st.expander("📋 File Status", expanded=True):
        for key in required_keys:
            label = LABELS_PT.get(key, key)
            in_upload    = files.get(key) is not None
            in_persisted = key in persisted

            if in_upload:
                icon = "✅"
                detail = getattr(files[key], "name", "")
            elif in_persisted:
                icon = "💾"
                detail = persisted[key].get("name", "")
            else:
                icon = "❌"
                detail = "Missing"

            st.markdown(f"{icon} **{label}** — `{detail}`")

        if optional_keys:
            st.markdown("---")
            st.caption("🔓 Optional — upload to unlock extra validation rules")
            for key in optional_keys:
                label = LABELS_PT.get(key, key)
                in_upload    = files.get(key) is not None
                in_persisted = key in persisted

                if in_upload:
                    icon   = "✅"
                    detail = getattr(files[key], "name", "")
                elif in_persisted:
                    icon   = "✅"
                    detail = persisted[key].get("name", "")
                else:
                    icon   = "⬜"
                    detail = "Not uploaded (optional)"

                st.markdown(f"{icon} **{label}** — `{detail}`")

        # ── Unrecognised files ────────────────────────────────────────────────
        known_keys  = set(required_keys) | set(optional_keys or [])
        all_matched = {k for k, v in files.items() if v is not None}
        # Files that were uploaded but didn't match any known slot
        # We need to check the raw upload list — files dict has None for unmatched
        # So: any uploaded file whose name doesn't map to a recognised key
        unrecognised = []
        for key, fobj in files.items():
            if fobj is not None and key not in known_keys:
                fname = getattr(fobj, "name", "")
                if fname:
                    unrecognised.append(fname)

        if unrecognised:
            st.markdown("---")
            st.caption("⚠️ Unrecognised files — could not be automatically matched:")
            # Build suggestion map: keyword in filename → what it might be
            _SUGGESTIONS = {
                "mara":       "MARA.xlsx",
                "marc":       "MARC.xlsx",
                "mvke":       "MVKE.xlsx",
                "qmat":       "QMAT.xlsx",
                "mbew":       "MBEW.xlsx",
                "ekpo":       "EKPO.xlsx",
                "mard":       "MARD.xlsx",
                "zplm":       "ZPLM MARC.xlsx or ZPLM MARA.xlsx",
                "sapui5":     "SAPUI5 Export.xlsx",
                "automotive": "PLM Automotive Validation.xlsx",
                "timing":     "Timing Letters.xlsx",
            }
            for fname in unrecognised:
                name_lower = fname.lower()
                suggestion = next(
                    (label for kw, label in _SUGGESTIONS.items() if kw in name_lower),
                    None
                )
                hint = f" → might be **{suggestion}**?" if suggestion else                        " → rename to match expected filename (e.g. MARA.xlsx, MARC.xlsx)"
                st.markdown(f"❓ `{fname}`{hint}")

# ==========================================================
# HELPERS — Timestamp / Progress mode
# ==========================================================
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

# ==========================================================
# AUXILIARES
# ==========================================================
def has_sheet(uploaded, sheet_name: str):
    try:
        xl = pd.ExcelFile(uploaded)
        return sheet_name in xl.sheet_names, xl.sheet_names
    except Exception as e:
        return False, [f"Error opening file: {e}"]

# colorize_errors removed — superseded by Error_Type column and _coerce_df_for_display

# ==========================================================
# IDENTIDADE
# ==========================================================
if "executor_name" not in st.session_state:
    st.session_state["executor_name"] = ""

def identity_input(label: str, key: str) -> str:
    """
    FIX 9: All pages share a single CDS ID stored in session_state["executor_name"].
    The per-page key is kept only for widget uniqueness but the value is always
    read from and written to the shared session key — so navigating between pages
    never clears the CDS ID the user already typed.
    """
    executor_name = st.text_input(
        label,
        value=st.session_state.get("executor_name", ""),
        placeholder="CDS ID...",
        key=key,
    )
    if executor_name.strip():
        st.session_state["executor_name"] = executor_name.strip()
    if not st.session_state.get("executor_name", ""):
        st.warning("⚠ Please enter your CDS ID to continue.")
        st.stop()
    return st.session_state["executor_name"]


def render_history_with_download(run_type: str, limit: int = 10,
                                  col_map: dict | None = None,
                                  label: str = "History") -> None:
    """
    Renders a history expander that shows the last N runs.
    For runs that have a saved report on disk, shows an inline download button.

    col_map: dict of {display_label: log_key_path}
      e.g. {"When": "timestamp", "Who": "user", "Errors": "metrics.total_errors"}
    """
    def _get(log, dotted_key):
        parts = dotted_key.split(".")
        v = log
        for p in parts:
            v = v.get(p, {}) if isinstance(v, dict) else {}
        return v if v != {} else "—"

    with st.expander(f"🕘 {label} (last {limit} runs)", expanded=False):
        logs = list_run_logs(run_type, limit=limit)
        if not logs:
            st.caption("No history yet.")
            return

        if col_map is None:
            col_map = {"When": "timestamp", "Who": "user"}

        for log in logs:
            row_cols = list(col_map.keys())
            # Add a Download column if a report_path exists on disk
            has_file = bool(log.get("report_path") and os.path.exists(log["report_path"]))
            n_cols = len(row_cols) + (1 if has_file else 0)
            cols = st.columns([3] * len(row_cols) + ([2] if has_file else []))

            for i, (display, key) in enumerate(col_map.items()):
                with cols[i]:
                    val = _get(log, key)
                    st.caption(display)
                    st.markdown(f"**{val}**")

            if has_file:
                with cols[-1]:
                    st.caption("Report")
                    try:
                        with open(log["report_path"], "rb") as _f:
                            _data = _f.read()
                        st.download_button(
                            "⬇️",
                            data=_data,
                            file_name=os.path.basename(log["report_path"]),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{run_type}_{log.get('timestamp','').replace(':','-')}",
                            help=f"Download: {os.path.basename(log['report_path'])}",
                        )
                    except Exception:
                        st.caption("—")
            st.divider()

# ==========================================================
# ASSIGNMENT PANEL — render in SAP Step 4
# ==========================================================
def render_assignment_panel(errors_df: pd.DataFrame) -> None:
    """
    Renders the 📋 Assign & Track panel in SAP Step 4.
    Shows all distinct fields with errors, their current owner,
    and allows inline editing — with changes saved immediately to master.json.
    """
    if errors_df is None or errors_df.empty:
        return

    st.markdown("## 📋 Assign & Track")
    st.markdown(
        """<div style="background:var(--app-bg-secondary);border-left:4px solid var(--primary);
        border-radius:6px;padding:10px 16px;margin-bottom:16px;font-size:0.85rem;color:var(--app-fg)">
        Assign each error field to a responsible person. Assignments are <strong>saved permanently</strong>
        — on the next validation run, owners are pre-filled automatically. Edit any time with the ✏️ button.
        </div>""",
        unsafe_allow_html=True,
    )

    # Build summary: distinct (Field, Plant) combinations with error counts
    _fe = errors_df[errors_df.get("Error_Type", pd.Series("Field Error", index=errors_df.index)) == "Field Error"]         if "Error_Type" in errors_df.columns else errors_df

    if _fe.empty:
        st.info("No field errors to assign.")
        return

    # Group by Field (global) — plant overrides handled via expander
    field_counts = (
        _fe.groupby("Field")
           .agg(Errors=("Field","count"),
                Plants=("Plant", lambda x: ", ".join(sorted(x.astype(str).str.strip().unique()))))
           .reset_index()
           .sort_values("Errors", ascending=False)
    ) if "Field" in _fe.columns else pd.DataFrame()

    if field_counts.empty:
        return

    # Header row
    hcols = st.columns([4, 2, 3, 2, 1])
    for col, label in zip(hcols, ["Field", "Errors", "Owner", "Team", "Action"]):
        col.markdown(f"<div style='font-size:0.75rem;font-weight:700;text-transform:uppercase;"
                     f"letter-spacing:.06em;color:var(--app-fg-muted)'>{label}</div>",
                     unsafe_allow_html=True)

    st.markdown("<hr style='margin:4px 0 8px 0;border-color:var(--app-border)'>",
                unsafe_allow_html=True)

    # Load master once — avoid repeated disk reads inside the loop
    from sap_validation_engine import _load_master
    _master_cache = _load_master()

    for _, row in field_counts.iterrows():
        field  = row["Field"]
        n_errs = row["Errors"]
        plants = row.get("Plants", "")

        current_owner = get_owner(field, master=_master_cache)
        _fd = _master_cache.get("field_assignments", {}).get(field, {})
        current_team = _fd.get("team", "") if isinstance(_fd, dict) else ""

        edit_key = f"asgn_edit_{field}"
        if edit_key not in st.session_state:
            st.session_state[edit_key] = False

        cols = st.columns([4, 2, 3, 2, 1])

        with cols[0]:
            st.markdown(f"**{field}**<br>"
                        f"<span style='font-size:0.75rem;color:var(--app-fg-muted)'>{plants}</span>",
                        unsafe_allow_html=True)
        with cols[1]:
            _badge_color = "var(--color-error)" if n_errs > 50 else                            "var(--color-warn)"  if n_errs > 10 else "var(--app-fg-muted)"
            st.markdown(f"<span style='font-weight:700;color:{_badge_color}'>{n_errs:,}</span>",
                        unsafe_allow_html=True)
        with cols[2]:
            if st.session_state[edit_key]:
                new_owner = st.text_input("Owner", value=current_owner,
                                          key=f"inp_owner_{field}", label_visibility="collapsed",
                                          placeholder="Name or CDS ID…")
            else:
                if current_owner:
                    st.markdown(f"👤 **{current_owner}**", unsafe_allow_html=True)
                else:
                    st.markdown("<span style='color:var(--app-fg-muted);font-style:italic'>"
                                "Unassigned</span>", unsafe_allow_html=True)
        with cols[3]:
            if st.session_state[edit_key]:
                new_team = st.text_input("Team", value=current_team,
                                         key=f"inp_team_{field}", label_visibility="collapsed",
                                         placeholder="Team (optional)…")
            else:
                if current_team:
                    st.markdown(f"<span style='font-size:0.8rem;color:var(--app-fg-muted)'>"
                                f"{current_team}</span>", unsafe_allow_html=True)

        with cols[4]:
            if st.session_state[edit_key]:
                if st.button("💾", key=f"save_{field}",
                             help=f"Save assignment for {field}"):
                    _owner_val = st.session_state.get(f"inp_owner_{field}", "").strip()
                    _team_val  = st.session_state.get(f"inp_team_{field}",  "").strip()
                    set_owner(field, _owner_val, team=_team_val)
                    st.session_state[edit_key] = False
                    st.rerun()
            else:
                if st.button("✏️", key=f"edit_{field}",
                             help=f"Edit assignment for {field}"):
                    st.session_state[edit_key] = True
                    st.rerun()

        st.markdown("<hr style='margin:4px 0;border-color:var(--app-border)'>",
                    unsafe_allow_html=True)

    # ── Bulk assign by plant override ─────────────────────────────────────
    with st.expander("🏭 Plant-level overrides (assign a field differently per plant)", expanded=False):
        st.caption("Use this when PT01 and PT02 have different responsible people for the same field.")
        _ov_field  = st.selectbox("Field", options=sorted(field_counts["Field"].tolist()),
                                  key="ov_field")
        _ov_plant  = st.selectbox("Plant", options=["PT01", "PT02"], key="ov_plant")
        _ov_owner  = st.text_input("Owner for this plant", key="ov_owner",
                                   placeholder="Name or CDS ID…")
        _ov_team   = st.text_input("Team (optional)", key="ov_team")
        if st.button("💾 Save plant override", key="ov_save"):
            if _ov_owner.strip():
                set_owner(_ov_field, _ov_owner.strip(),
                          plant=_ov_plant, team=_ov_team.strip())
                st.success(f"✅ Saved: {_ov_field} / {_ov_plant} → {_ov_owner.strip()}")
                st.rerun()
            else:
                st.warning("Please enter an owner name.")

    # ── View all current assignments + export ─────────────────────────────
    with st.expander("📜 All current assignments", expanded=False):
        _all_asgn = list_all_assignments()
        if _all_asgn.empty:
            st.caption("No assignments saved yet.")
        else:
            st.dataframe(_all_asgn, use_container_width=True)

            # Export: per-owner summary
            _exp_cols = st.columns([2, 1])
            with _exp_cols[0]:
                # Allow removing an assignment
                _rm_cols = st.columns(3)
                with _rm_cols[0]:
                    _rm_field = st.selectbox("Remove assignment for field",
                                             options=["—"] + sorted(_all_asgn["Field"].unique().tolist()),
                                             key="rm_field")
                with _rm_cols[1]:
                    _rm_plant = st.selectbox("Plant (leave blank for default)",
                                             options=["(default)", "PT01", "PT02"],
                                             key="rm_plant")
                with _rm_cols[2]:
                    st.markdown("<div style='margin-top:22px'>", unsafe_allow_html=True)
                    if st.button("🗑️ Remove", key="rm_confirm"):
                        if _rm_field != "—":
                            _plant_arg = "" if _rm_plant == "(default)" else _rm_plant
                            remove_owner(_rm_field, plant=_plant_arg)
                            st.success(f"Removed assignment for {_rm_field}" +
                                       (f" / {_rm_plant}" if _plant_arg else ""))
                            st.rerun()
                    st.markdown("</div>", unsafe_allow_html=True)

            with _exp_cols[1]:
                st.markdown("<div style='margin-top:4px'>", unsafe_allow_html=True)
                # Build per-owner summary Excel
                try:
                    _enriched_exp = merge_assignments(errors_df) if errors_df is not None else pd.DataFrame()
                    if not _enriched_exp.empty and "Owner" in _enriched_exp.columns:
                        # Summary by owner
                        _owner_summary = (
                            _enriched_exp[_enriched_exp["Owner"] != ""]
                            .groupby(["Owner", "Team"])
                            .agg(
                                Fields=("Field", lambda x: ", ".join(sorted(x.unique()))),
                                Total_Errors=("Field", "count"),
                                Overdue=("Status", lambda x: (x == "Overdue").sum()),
                                Avg_Days_Open=("Days_Open", "mean"),
                            )
                            .round(1)
                            .reset_index()
                            .sort_values("Total_Errors", ascending=False)
                        )
                        # Full assignments sheet
                        _asgn_sheet = build_assignments_sheet(_enriched_exp)
                        _exp_buf = io.BytesIO()
                        with pd.ExcelWriter(_exp_buf, engine="openpyxl") as _ew:
                            _owner_summary.to_excel(_ew, sheet_name="By Owner", index=False)
                            _asgn_sheet.to_excel(_ew, sheet_name="All Assignments", index=False)
                            _all_asgn.to_excel(_ew, sheet_name="Field Owners", index=False)
                        _exp_buf.seek(0)
                        st.download_button(
                            "⬇️ Export Assignments",
                            data=_exp_buf.getvalue(),
                            file_name=f"Assignments_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_assignments_export",
                            help="Download per-owner summary + full assignment list",
                        )
                except Exception:
                    st.caption("Export unavailable")
                st.markdown("</div>", unsafe_allow_html=True)


# ==========================================================
# SIDEBAR MENU
# ==========================================================
st.sidebar.header("Report Selection")

main_choice = st.sidebar.radio(
    "Navigation",
    [
        "📈 Evolution Dashboard",
        "📦 SAP Material Validation",
        "🚗 PLM Workflow",
        "🧮 MRP Validation",
        "📊 SAP Validation Progress",
        "🧹 Data Cleanups",
        "📋 SSA Validation",
    ],
    label_visibility="collapsed",
)

sub_choice = None
if main_choice == "🚗 PLM Workflow":
    st.sidebar.markdown(
        "<div class='sidebar-submenu-title'>▼ Options</div>",
        unsafe_allow_html=True,
    )
    with st.sidebar.container():
        st.sidebar.markdown("<div class='sidebar-submenu-wrap'>", unsafe_allow_html=True)
        sub_choice = st.sidebar.radio(
            "PLM Submenu",
            ["Automotive + Timing Letter (A & TL)", "Change Cockpit"],
            key="submenu_plm",
            label_visibility="collapsed",
        )
        st.sidebar.markdown("</div>", unsafe_allow_html=True)

# ── About / Changelog ────────────────────────────────────────────────────────
st.sidebar.markdown("---")
with st.sidebar.expander("ℹ️ About Validation Hub", expanded=False):
    st.markdown("""
    **Version:** v1.4  
    **Build:** """ + datetime.now().strftime("%Y-%m-%d") + """

    **Recent changes:**
    - 📋 Assignment system — assign fields to owners, track overdue
    - 📊 Assignment Health in Evolution Dashboard
    - ✅ File column validation before running
    - 🔍 Unrecognised file detection with suggestions
    - ⬇️ Re-download previous reports from history
    - 🌙 Full dark mode support (CSS variables)

    **Modules:** SAP Validation · PLM Workflow · MRP Validation · Data Cleanups · Progress · Evolution Dashboard
    """)

# ==========================================================
# 📈 EVOLUTION DASHBOARD
# ==========================================================
if main_choice == "📈 Evolution Dashboard":
    render_evolution_dashboard()
    render_footer(start_time, version="v1.4")

# ==========================================================
# PAGE 1 — SAP MATERIAL VALIDATION
# ==========================================================
elif main_choice == "📦 SAP Material Validation":

    st.subheader("📦 SAP Material Validation Tool")

    effective_executor = identity_input("Who performs?", "sap_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    if "sap_step" not in st.session_state:
        st.session_state.sap_step = 1

    render_wizard(st.session_state.sap_step, WIZARD_STEPS_DEFAULT)

    required_keys = REQUIRED_KEYS["📦 SAP Material Validation"]
    persisted_now = st.session_state.get("sap_files_persisted") or {}

    # Guard: se ficheiros em falta volta ao passo 1
    if st.session_state.sap_step > 1 and any(k not in persisted_now for k in required_keys):
        st.session_state.sap_step = 1
        st.rerun()

    # --- STEP 1: FILES ---
    if st.session_state.sap_step == 1:
        st.markdown("### 📂 Step 1 — Load Files")
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--accent);border-radius:6px;padding:12px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>📋 What files do you need?</strong><br><br>
        <table style="border-collapse:collapse;width:100%">
        <tr><td style="padding:3px 8px">📄 <strong>SAPUI5 Export</strong></td><td style="color:var(--app-fg-muted)">SAP → Logistics → Item Browse Report (export as .xlsx)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MARA</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → MARA (general material data)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MARC</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → MARC (plant-level material data)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MVKE</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → MVKE (sales organisation data)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>ZPLM MARC / MARA</strong></td><td style="color:var(--app-fg-muted)">PLM custom tables (ZPLM_MARC / ZPLM_MARA export)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>QMAT</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → QMAT (inspection types)</td></tr>
        <tr><td style="padding:3px 8px; color:var(--color-ok)">🔓 <strong>CSR Item / Head</strong></td><td style="color:var(--app-fg-muted)">Optional — enables CSR exemption rules for MRP Type</td></tr>
        </table>
        </div>
        """, unsafe_allow_html=True)

        uploads = st.file_uploader(
            "Drop all necessary files here",
            type="xlsx",
            accept_multiple_files=True,
            key="sap_multi_uploader",
        )

        files = classify_uploaded_files(uploads)
        if uploads:
            persist_uploaded_mapping(files, "sap_files_persisted")

        persisted = st.session_state.get("sap_files_persisted") or {}
        show_file_status(required_keys, files, persisted,
                         optional_keys=["csr_item", "csr_head"])

        # Quick column validation
        _col_errs_sap = validate_uploaded_files(persisted, required_keys)
        if _col_errs_sap:
            for _k, _msg in _col_errs_sap.items():
                st.warning(f"⚠️ **{LABELS_PT.get(_k, _k)}** — {_msg}")

        disabled = any(k not in persisted for k in required_keys) or bool(_col_errs_sap)
        c1, c2 = st.columns(2)
        with c1:
            with st.expander("🧹 Clear uploaded files"):
                if st.button("Clear persistent uploads", key="sap_clear"):
                    st.session_state["sap_files_persisted"] = {}
                    st.success("Cleared.")
                    st.rerun()
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, key="sap_next1", use_container_width=True):
                st.session_state.sap_step = 2
                st.rerun()

    # --- STEP 2: SCOPE ---
    elif st.session_state.sap_step == 2:
        st.markdown("### 🌍 Step 2 — Select Scope")
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--color-ok);border-radius:6px;padding:10px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>ℹ️ What does scope mean?</strong><br>
        <strong>Scope of download</strong> controls which plant's data appears in the exported Excel report.<br>
        <strong>Apply filter to displayed tables</strong> also filters the results shown here in the app — leave unchecked to see all plants in the UI while exporting only the chosen plant.
        </div>
        """, unsafe_allow_html=True)

        export_scope = st.radio("Scope of download:", ["All", "Only PT01", "Only PT02"], horizontal=True)
        plant_filter_export = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(export_scope)

        apply_filter_display = st.checkbox("Apply filter to displayed tables", value=False)
        plant_filter_display = plant_filter_export if apply_filter_display else None

        st.session_state.sap_export_scope    = export_scope
        st.session_state.sap_plant_export    = plant_filter_export
        st.session_state.sap_plant_display   = plant_filter_display

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="sap_back2", use_container_width=True):
                st.session_state.sap_step = 1
                st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="sap_next2", use_container_width=True):
                st.session_state.sap_step = 3
                st.rerun()

    # --- STEP 3: VALIDATE ---
    elif st.session_state.sap_step == 3:
        st.markdown("### ⚙️ Step 3 — Run Validation")

        # Confirmation summary
        _persisted3 = st.session_state.get("sap_files_persisted") or {}
        _scope3     = st.session_state.get("sap_export_scope", "All")
        _exec3      = st.session_state.get("executor_name", "—")
        _n_files    = len(_persisted3)
        _optional3  = [k for k in ["csr_item","csr_head"] if k in _persisted3]

        st.markdown(f"""
        <div style="background:var(--app-bg-secondary);border:1px solid var(--app-border);border-radius:10px;padding:16px 20px;margin-bottom:16px">
        <div style="font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--app-fg-muted);margin-bottom:10px">
            ✅ Ready to validate — confirm before running
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;font-size:0.87rem;color:var(--app-fg)">
            <div><span style="color:var(--app-fg-muted)">Executor</span><br><strong>{_exec3}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Scope</span><br><strong>{_scope3}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Files loaded</span><br><strong>{_n_files} required{"  +  " + str(len(_optional3)) + " optional" if _optional3 else ""}</strong></div>
        </div>
        {"<div style='margin-top:10px;font-size:0.8rem;color:var(--color-ok)'>🔓 CSR exemption rules <strong>enabled</strong> — active CSRs will exempt FP materials from MRP = ND errors.</div>" if _optional3 else "<div style='margin-top:10px;font-size:0.8rem;color:var(--app-fg-muted)'>🔓 CSR files not uploaded — CSR exemption rules disabled.</div>"}
        </div>
        """, unsafe_allow_html=True)

        if st.button("▶️ Validate now", type="primary", key="sap_run"):
            files = rehydrate_persisted_mapping("sap_files_persisted")
            sap_file  = files.get("sap")
            mara_file = files.get("mara")
            marc_file = files.get("marc")
            mvke_file = files.get("mvke")
            zplm_marc = files.get("zplm_marc")
            zplm_mara = files.get("zplm_mara")
            qmat_file = files.get("qmat")
            csr_item  = files.get("csr_item")
            csr_head  = files.get("csr_head")

            plant_filter_export  = st.session_state.sap_plant_export
            plant_filter_display = st.session_state.sap_plant_display

            ok_sheet, sheets = has_sheet(sap_file, "SAPUI5 Export")
            if not ok_sheet:
                st.error(f"Sheet 'SAPUI5 Export' not found. Available: {sheets}")
                st.stop()

            _pb_slot = st.empty()
            progress_bar = _pb_slot.progress(0, text="Initializing validation...")

            def update_progress(percent: int, message: str = ""):
                progress_bar.progress(percent, text=message)

            with st.spinner("Validating materials..."):
                excel_display, errors_df, summary_df = run_full_validation(
                    sap_file=sap_file, mara_file=mara_file, marc_file=marc_file,
                    mvke_file=mvke_file, qmat_file=qmat_file,
                    zplm_marc_file=zplm_marc, zplm_mara_file=zplm_mara,
                    previous_report=None,
                    plant_filter=plant_filter_display,
                    progress_callback=update_progress,
                    csr_item_file=csr_item, csr_head_file=csr_head,
                )

                try:
                    st.session_state["sap_view_missing_df"] = pd.read_excel(excel_display, sheet_name="View is Missing")
                except Exception:
                    st.session_state["sap_view_missing_df"] = pd.DataFrame()

                # FIX 1: Avoid running the full engine twice.
                # When the export scope is different from the display scope,
                # we only need to re-run the styling step on the already-computed
                # errors_df filtered to the export scope — not re-run all validations.
                # The engine already supports plant_filter on _style_sap_report.
                if plant_filter_export != plant_filter_display:
                    from sap_validation_engine import _style_sap_report
                    from sap_validation_engine import postprocess_with_views_roles_and_classification
                    from rolesviews_data import get_rolesviews_bytesio
                    from utils.fast_loader import load_xlsx_fast as _lxf
                    import io as _io
                    # Re-filter errors to export scope and rebuild Excel
                    _rv = _lxf(get_rolesviews_bytesio(), sheet_name=None)
                    _roles = _rv.get("Roles"); _views = _rv.get("Views")
                    _mara_df = _lxf(mara_file, sheet_name=0)
                    _marc_df = _lxf(marc_file, sheet_name=0)
                    (_edf, _, _vmout, _cdf, _covdf, _vtabs) =                         postprocess_with_views_roles_and_classification(
                            errors_df, _mara_df, _marc_df, _roles, _views,
                            plant_filter=plant_filter_export)
                    _vis_df = pd.concat(
                        [df for df in _vtabs.values() if df is not None and not df.empty],
                        ignore_index=True
                    ) if _vtabs else pd.DataFrame()
                    _buf = _io.BytesIO()
                    with pd.ExcelWriter(_buf, engine="openpyxl") as _w:
                        pd.DataFrame().to_excel(_w, sheet_name="Executive Summary", index=False)
                        _cdf.to_excel(_w, sheet_name="Adjusted Summary", index=False, startrow=1)
                        _vmout.to_excel(_w, sheet_name="Field With Error", index=False)
                        _edf.to_excel(_w, sheet_name="All Errors", index=False)
                        _vis_df.to_excel(_w, sheet_name="View is Missing", index=False)
                        for _vc, _vdf in _vtabs.items():
                            _vdf.to_excel(_w, sheet_name=f"Missing_{_vc}"[:31], index=False)
                        _lxf(sap_file, sheet_name="SAPUI5 Export").to_excel(
                            _w, sheet_name="Original Data", index=False)
                    _buf.seek(0)
                    excel_export = _style_sap_report(
                        _buf, errors_df=_edf, views_missing_output_df=_vmout,
                        contagem_final_df=_cdf, coverage_df=_covdf,
                        view_tabs_dict=_vtabs, sap_df=None,
                        plant_filter=plant_filter_export)
                else:
                    excel_export = excel_display

            suffix = plant_filter_export or "ALL"
            report_path = save_report_to_vault(excel_export, report_type="sap", plant_suffix=suffix)
            st.session_state.sap_report_path  = report_path
            st.session_state.sap_errors_df    = errors_df
            st.session_state.sap_summary_df   = summary_df
            st.session_state.sap_excel_export = excel_export

            if save_history:
                payload = build_sap_log_payload(
                    user=effective_executor, plant_filter=plant_filter_export,
                    sap_file=sap_file, mara_file=mara_file, marc_file=marc_file,
                    mvke_file=mvke_file, errors_df=errors_df, summary_df=summary_df,
                    report_path=report_path,
                )
                save_run_log(payload)

            _pb_slot.empty()  # clear progress bar before navigating
            st.session_state.sap_step = 4
            st.rerun()

    # --- STEP 4: RESULTS ---
    elif st.session_state.sap_step == 4:
        st.markdown("### 📊 Step 4 — Results")

        errors_df    = st.session_state.sap_errors_df
        summary_df   = st.session_state.sap_summary_df
        excel_export = st.session_state.sap_excel_export
        report_path  = st.session_state.sap_report_path

        # FIX 4: count fields/materials consistently with Adjusted Summary
        # (Field Errors only — excludes Missing View errors)
        _fe_only = (errors_df[errors_df["Error_Type"] == "Field Error"]
                    if "Error_Type" in errors_df.columns else errors_df)
        _fe_cnt   = int(_fe_only.shape[0])
        _f_cnt    = int(_fe_only["Field"].nunique())    if "Field"    in _fe_only.columns else 0
        _m_cnt    = int(_fe_only["Material"].nunique()) if "Material" in _fe_only.columns else 0

        c = st.columns(3)
        with c[0]:
            st.markdown(f'<div class="card"><h4>Field Errors (Actionable)</h4><div class="big">{_fe_cnt:,}</div></div>', unsafe_allow_html=True)
        with c[1]:
            st.markdown(f'<div class="card"><h4>Distinct Fields</h4><div class="big">{_f_cnt}</div></div>', unsafe_allow_html=True)
        with c[2]:
            st.markdown(f'<div class="card"><h4>Materials with Errors</h4><div class="big">{_m_cnt:,}</div></div>', unsafe_allow_html=True)

        # Auto-insight — a plain-language summary of the results
        _top_field = _fe_only['Field'].value_counts().index[0] if (not _fe_only.empty and 'Field' in _fe_only.columns) else "—"
        _top_plant = (
            _fe_only['Plant'].astype(str).str.strip().replace('', 'General')
            .value_counts().index[0]
            if (not _fe_only.empty and 'Plant' in _fe_only.columns) else "—"
        )
        _mv_count = int((errors_df['Error_Type'] == 'Missing View').sum()) if 'Error_Type' in errors_df.columns else 0
        _insight_color = "var(--app-bg-secondary)" if _fe_cnt > 1000 else ("var(--app-bg-secondary)" if _fe_cnt > 200 else "var(--app-bg-secondary)")
        _insight_border = "var(--color-error)" if _fe_cnt > 1000 else ("var(--color-warn)" if _fe_cnt > 200 else "var(--color-ok)")
        st.markdown(f"""
        <div style="background:{_insight_color};border-left:4px solid {_insight_border};border-radius:8px;
                    padding:12px 16px;margin-bottom:16px;font-size:0.9rem;line-height:1.6;color:var(--app-fg)">
            <strong>📋 Validation Summary</strong><br>
            Found <strong>{_fe_cnt:,} actionable errors</strong> across <strong>{_m_cnt:,} materials</strong>
            in <strong>{_f_cnt} distinct fields</strong>.
            The most affected field is <strong>{_top_field}</strong> and the most affected plant is <strong>{_top_plant}</strong>.
            {"Additionally, " + str(_mv_count) + " errors were suppressed because the required SAP View is not yet active." if _mv_count else ""}
        </div>
        """, unsafe_allow_html=True)

        # FIX 7: Charts first, then tables in expanders
        st.markdown("## 📈 Dashboard — SAP Material Validation")
        render_sap_run_kpis(errors_df)
        render_sap_run_charts(errors_df)

        download_name = os.path.basename(report_path) if report_path else "SAP_Validation_Adjusted_Report_ALL.xlsx"
        st.download_button("⬇️ Download Excel", data=excel_export.getvalue(),
                           file_name=download_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # ── Assignment panel ──────────────────────────────────────────────
        render_assignment_panel(errors_df)

        with st.expander("📋 All Errors", expanded=False):
            # Quick search filter
            search = st.text_input("🔍 Filter by Material, Field or Error", key="sap_err_search")
            df_show = _coerce_df_for_display(errors_df)
            if search.strip():
                mask = df_show.apply(
                    lambda col: col.astype(str).str.contains(search.strip(), case=False, na=False)
                ).any(axis=1)
                df_show = df_show[mask]
            st.caption(f"Showing {len(df_show):,} of {len(errors_df):,} rows")
            st.dataframe(df_show, use_container_width=True)

        with st.expander("📋 Adjusted Summary", expanded=False):
            st.dataframe(_coerce_df_for_display(summary_df), use_container_width=True)

        render_history_with_download(
            "sap", limit=10, label="SAP History",
            col_map={"When": "timestamp", "Who": "user",
                     "Plant": "plant_filter", "Errors": "metrics.total_errors"}
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔁 New Validation", key="sap_new", use_container_width=True):
                st.session_state.sap_step = 1
                st.rerun()
        with c2:
            if st.button("🧹 Clear uploads & start over", key="sap_clear4", use_container_width=True):
                st.session_state["sap_files_persisted"] = {}
                st.session_state.sap_step = 1
                st.rerun()

    render_footer(start_time, version="v1.4")

# ==========================================================
# PAGE 2A — PLM WORKFLOW: ATL
# ==========================================================
elif main_choice == "🚗 PLM Workflow" and sub_choice == "Automotive + Timing Letter (A & TL)":

    st.subheader("🚗 PLM Workflow (ATL)")

    effective_executor = identity_input("Who performs?", "atl_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    if "atl_step" not in st.session_state:
        st.session_state.atl_step = 1

    render_wizard(st.session_state.atl_step, WIZARD_STEPS_DEFAULT)

    required_keys = REQUIRED_KEYS["Automotive + Timing Letter (A & TL)"]

    # --- STEP 1 ---
    if st.session_state.atl_step == 1:
        st.markdown("### 📂 Step 1 — Load Files")
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--accent);border-radius:6px;padding:12px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>📋 What files do you need?</strong><br><br>
        <table style="border-collapse:collapse;width:100%">
        <tr><td style="padding:3px 8px">📄 <strong>Automotive Validation</strong></td><td style="color:var(--app-fg-muted)">PLM → Automotive validation export (filename must contain <em>automotive</em>)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>TimingLetter Validation</strong></td><td style="color:var(--app-fg-muted)">PLM → Timing Letter validation export (filename must contain <em>timingletter</em>)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>Timing Letters</strong></td><td style="color:var(--app-fg-muted)">PLM → Active Timing Letters list (filename must contain <em>timing</em>)</td></tr>
        </table>
        </div>
        """, unsafe_allow_html=True)

        uploads = st.file_uploader(
            "Drop Automotive + TimingLetter + Timing Letter files",
            type="xlsx", accept_multiple_files=True, key="atl_multi_uploader",
        )
        files = classify_uploaded_files(uploads)
        if uploads:
            persist_uploaded_mapping(files, "atl_files_persisted")

        persisted = st.session_state.get("atl_files_persisted") or {}
        show_file_status(required_keys, files, persisted)

        _col_errs_atl = validate_uploaded_files(persisted, required_keys)
        if _col_errs_atl:
            for _k, _msg in _col_errs_atl.items():
                st.warning(f"⚠️ **{LABELS_PT.get(_k, _k)}** — {_msg}")

        disabled = any(k not in persisted for k in required_keys) or bool(_col_errs_atl)
        c1, c2 = st.columns(2)
        with c1:
            with st.expander("🧹 Clear persistent uploads"):
                if st.button("Clear ATL uploads", key="atl_clear"):
                    st.session_state["atl_files_persisted"] = {}
                    st.success("Cleared.")
                    st.rerun()
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, key="atl_next1", use_container_width=True):
                st.session_state.atl_step = 2
                st.rerun()

    # --- STEP 2 ---
    elif st.session_state.atl_step == 2:
        st.markdown("### 🌍 Step 2 — Choose Scope")

        atl_scope = st.radio("Scope:", ["All", "Only PT01", "Only PT02"], horizontal=True)
        atl_filter_export = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(atl_scope)
        apply_filter_display = st.checkbox("Apply filter to displayed tables", value=False)
        atl_filter_display = atl_filter_export if apply_filter_display else None

        st.session_state.atl_scope          = atl_scope
        st.session_state.atl_filter_export  = atl_filter_export
        st.session_state.atl_filter_display = atl_filter_display

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="atl_back2", use_container_width=True):
                st.session_state.atl_step = 1
                st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="atl_next2", use_container_width=True):
                st.session_state.atl_step = 3
                st.rerun()

    # --- STEP 3 ---
    elif st.session_state.atl_step == 3:
        st.markdown("### ⚙️ Step 3 — Run ATL Workflow")

        _exec_atl   = st.session_state.get("executor_name", "—")
        _scope_atl  = st.session_state.get("atl_scope", "All")
        _pers_atl   = st.session_state.get("atl_files_persisted") or {}

        st.markdown(f"""
        <div style="background:var(--app-bg-secondary);border:1px solid var(--app-border);border-radius:10px;padding:16px 20px;margin-bottom:16px">
        <div style="font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--app-fg-muted);margin-bottom:10px">
            ✅ Ready to process — confirm before running
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;font-size:0.87rem;color:var(--app-fg)">
            <div><span style="color:var(--app-fg-muted)">Executor</span><br><strong>{_exec_atl}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Scope</span><br><strong>{_scope_atl}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Files loaded</span><br><strong>{len(_pers_atl)} of 3</strong></div>
        </div>
        <div style="margin-top:10px;font-size:0.8rem;color:var(--app-fg-muted)">
            The workflow will match Change Numbers from Automotive and TimingLetter sources,
            filter active Timing Letters (excludes Complete/Rejected), and compute Aging for each CN.
        </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("▶️ Run Workflow ATL", type="primary", key="atl_run"):
            files = rehydrate_persisted_mapping("atl_files_persisted")
            auto_file   = files.get("auto")
            plm_file    = files.get("plm")
            timing_file = files.get("timing")

            if any(f is None for f in [auto_file, plm_file, timing_file]):
                st.error("⚠️ Required files are missing.")
                st.stop()

            with st.spinner("Processing ATL Workflow..."):
                excel_all, df_pt01, df_pt02 = run_plm_workflow_combined(
                    auto_file=auto_file, plm_file=plm_file, timing_file=timing_file
                )

            st.session_state.atl_pt01      = df_pt01
            st.session_state.atl_pt02      = df_pt02
            st.session_state.atl_excel_all = excel_all

            if save_history:
                payload = build_atl_log_payload(
                    user=effective_executor, plant_scope=st.session_state.atl_scope,
                    df_pt01=df_pt01, df_pt02=df_pt02,
                )
                save_run_log(payload)

            st.session_state.atl_step = 4
            st.rerun()

    # --- STEP 4 ---
    elif st.session_state.atl_step == 4:
        st.markdown("### 📊 Step 4 — Results")

        df_pt01   = st.session_state.atl_pt01
        df_pt02   = st.session_state.atl_pt02
        excel_all = st.session_state.atl_excel_all
        scope     = st.session_state.atl_filter_display

        _total_cns = (len(df_pt01) if df_pt01 is not None else 0) + (len(df_pt02) if df_pt02 is not None else 0)
        _tln_cns   = 0
        for _df_tln in [df_pt01, df_pt02]:
            if _df_tln is not None and not _df_tln.empty and 'Timing Letter Number' in _df_tln.columns:
                _tln_cns += int(_df_tln['Timing Letter Number'].astype(str).str.strip().replace('', pd.NA).notna().sum())
        _no_tln = _total_cns - _tln_cns
        _ic_plm = "var(--app-bg-secondary)" if _total_cns > 50 else ("var(--app-bg-secondary)" if _total_cns > 20 else "var(--app-bg-secondary)")
        _ib_plm = "var(--color-error)" if _total_cns > 50 else ("var(--color-warn)" if _total_cns > 20 else "var(--color-ok)")
        st.markdown(f"""
        <div style="background:{_ic_plm};border-left:4px solid {_ib_plm};border-radius:8px;
                    padding:12px 16px;margin-bottom:16px;font-size:0.9rem;line-height:1.6;color:var(--app-fg)">
            <strong>📋 PLM Workflow Summary</strong><br>
            Found <strong>{_total_cns:,} open Change Numbers</strong> pending approval.
            <strong>{_tln_cns:,}</strong> have a Timing Letter Number assigned.
            {"<strong>" + str(_no_tln) + " CNs have no Timing Letter</strong> — these may need follow-up." if _no_tln else "All CNs have a Timing Letter assigned. ✅"}
        </div>
        """, unsafe_allow_html=True)

        st.markdown("## 📈 Dashboard — PLM Workflow (ATL)")
        render_atl_run_kpis(df_pt01, df_pt02)
        render_atl_run_charts(df_pt01, df_pt02)

        with st.expander("📋 PT01 Change Numbers", expanded=False):
            _srch_pt01 = st.text_input("🔍 Filter PT01", key="plm_pt01_search")
            _df_pt01_s = _coerce_df_for_display(df_pt01) if df_pt01 is not None else pd.DataFrame()
            if _srch_pt01.strip() and not _df_pt01_s.empty:
                _mask_p01 = _df_pt01_s.apply(lambda col: col.astype(str).str.contains(_srch_pt01.strip(), case=False, na=False)).any(axis=1)
                _df_pt01_s = _df_pt01_s[_mask_p01]
            st.caption(f"{len(_df_pt01_s):,} rows")
            st.dataframe(_df_pt01_s, use_container_width=True)

        with st.expander("📋 PT02 Change Numbers", expanded=False):
            _srch_pt02 = st.text_input("🔍 Filter PT02", key="plm_pt02_search")
            _df_pt02_s = _coerce_df_for_display(df_pt02) if df_pt02 is not None else pd.DataFrame()
            if _srch_pt02.strip() and not _df_pt02_s.empty:
                _mask_p02 = _df_pt02_s.apply(lambda col: col.astype(str).str.contains(_srch_pt02.strip(), case=False, na=False)).any(axis=1)
                _df_pt02_s = _df_pt02_s[_mask_p02]
            st.caption(f"{len(_df_pt02_s):,} rows")
            st.dataframe(_df_pt02_s, use_container_width=True)

        suffix = st.session_state.atl_filter_export or "ALL"
        st.download_button(
            f"⬇️ Download PLM_Workflow_{suffix}.xlsx",
            data=excel_all.getvalue(), file_name=f"PLM_Workflow_{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        render_history_with_download(
            "atl", limit=10, label="ATL History",
            col_map={"When": "timestamp", "Who": "user",
                     "Scope": "plant_scope", "PT01": "metrics.PT01_rows", "PT02": "metrics.PT02_rows"}
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Run ATL", key="atl_new", use_container_width=True):
                st.session_state.atl_step = 1
                st.rerun()
        with c2:
            if st.button("🧹 Clear uploads & start over", key="atl_clear4", use_container_width=True):
                st.session_state["atl_files_persisted"] = {}
                st.session_state.atl_step = 1
                st.rerun()

    render_footer(start_time, version="v1.4")

# ==========================================================
# PAGE 2B — PLM WORKFLOW: CHANGE COCKPIT
# ==========================================================
elif main_choice == "🚗 PLM Workflow" and sub_choice == "Change Cockpit":

    st.subheader("🛠️ Change Cockpit (In Progress...)")
    effective_executor = identity_input("Who performs?", "cc_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    if "cc_step" not in st.session_state:
        st.session_state.cc_step = 1

    render_wizard(st.session_state.cc_step, WIZARD_STEPS_DEFAULT)

    if st.session_state.cc_step == 1:
        st.markdown("### 📂 Step 1 — Load Files")
        uploads = st.file_uploader("Drop Change Cockpit input(s) here",
                                   type=["xlsx", "csv"], accept_multiple_files=True, key="cc_multi_uploader")
        # TODO: define required CC files and classify
        c1, c2 = st.columns(2)
        with c2:
            if st.button("➡️ Continue to Scope", key="cc_next1", use_container_width=True):
                st.session_state.cc_step = 2
                st.rerun()

    elif st.session_state.cc_step == 2:
        st.markdown("### 🌍 Step 2 — Select Scope")
        cc_scope = st.radio("Scope:", ["All", "Only PT01", "Only PT02"], horizontal=True)
        st.session_state.cc_scope = cc_scope
        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="cc_back2", use_container_width=True):
                st.session_state.cc_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="cc_next2", use_container_width=True):
                st.session_state.cc_step = 3; st.rerun()

    elif st.session_state.cc_step == 3:
        st.markdown("### ⚙️ Step 3 — Run Change Cockpit")
        st.info("Define here the processing logic for Change Cockpit.")
        if st.button("▶️ Run Change Cockpit", type="primary", key="cc_run"):
            with st.spinner("Running Change Cockpit..."):
                st.session_state.cc_result_df = None
                st.session_state.cc_excel     = None
            st.session_state.cc_step = 4; st.rerun()

    elif st.session_state.cc_step == 4:
        st.markdown("### 📈 Step 4 — Results")
        st.success("Results placeholder — plug your tables and dashboards here.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Run", key="cc_new", use_container_width=True):
                st.session_state.cc_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear & Start Over", key="cc_clear4", use_container_width=True):
                for k in ["cc_result_df", "cc_excel"]:
                    st.session_state.pop(k, None)
                st.session_state.cc_step = 1; st.rerun()

    render_footer(start_time, version="v1.4")

# ==========================================================
# PAGE 3 — MRP VALIDATION
# ==========================================================
elif main_choice == "🧮 MRP Validation":

    st.subheader("🧮 MRP Validation")
    effective_executor = identity_input("Who performs?", "mrp_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    if "mrp_step" not in st.session_state:
        st.session_state.mrp_step = 1

    render_wizard(st.session_state.mrp_step, WIZARD_STEPS_DEFAULT)

    required_keys = REQUIRED_KEYS["🧮 MRP Validation"]

    if st.session_state.mrp_step == 1:
        st.markdown("### 📂 Step 1 — Load Files")
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--accent);border-radius:6px;padding:12px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>📋 What files do you need?</strong><br><br>
        <table style="border-collapse:collapse;width:100%">
        <tr><td style="padding:3px 8px">📄 <strong>MARA</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → MARA (general material data, required for Material Type)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MARC</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → MARC (contains Advanced Planning, Heuristic, PP Plng Procedure fields)</td></tr>
        <tr><td style="padding:3px 8px; color:var(--color-ok)">🔓 <strong>ZPLM MARC</strong></td><td style="color:var(--app-fg-muted)">Optional — enables Material Transfer detection for MRP rules</td></tr>
        </table>
        </div>
        """, unsafe_allow_html=True)
        uploads = st.file_uploader("Drop MARA + MARC + (optional) ZPLM MARC",
                                   type="xlsx", accept_multiple_files=True, key="mrp_multi_uploader")
        files = classify_uploaded_files(uploads)
        if uploads:
            persist_uploaded_mapping(files, "mrp_files_persisted")

        persisted = st.session_state.get("mrp_files_persisted") or {}
        show_file_status(required_keys, files, persisted)

        _col_errs_mrp = validate_uploaded_files(persisted, required_keys)
        if _col_errs_mrp:
            for _k, _msg in _col_errs_mrp.items():
                st.warning(f"⚠️ **{LABELS_PT.get(_k, _k)}** — {_msg}")

        disabled = any(k not in persisted for k in required_keys) or bool(_col_errs_mrp)
        c1, c2 = st.columns(2)
        with c1:
            with st.expander("🧹 Clear persistent uploads"):
                if st.button("Clear MRP uploads", key="mrp_clear"):
                    st.session_state["mrp_files_persisted"] = {}
                    st.success("Cleared."); st.rerun()
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, key="mrp_next1", use_container_width=True):
                st.session_state.mrp_step = 2; st.rerun()

    elif st.session_state.mrp_step == 2:
        st.markdown("### 🌍 Step 2 — Choose Scope")
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--color-ok);border-radius:6px;padding:10px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>ℹ️ What does scope mean?</strong><br>
        Choose which plant to validate. <strong>All</strong> validates both PT01 and PT02 together.
        The scope also determines which plant appears in the exported Excel report.
        </div>
        """, unsafe_allow_html=True)
        export_scope = st.radio("Scope:", ["All", "Only PT01", "Only PT02"], horizontal=True)
        plant_filter_export = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(export_scope)
        apply_filter_display = st.checkbox("Apply filter to displayed tables", value=False)
        plant_filter_display = plant_filter_export if apply_filter_display else None

        st.session_state.mrp_export_scope  = export_scope
        st.session_state.mrp_plant_export  = plant_filter_export
        st.session_state.mrp_plant_display = plant_filter_display

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="mrp_back2", use_container_width=True):
                st.session_state.mrp_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="mrp_next2", use_container_width=True):
                st.session_state.mrp_step = 3; st.rerun()

    elif st.session_state.mrp_step == 3:
        st.markdown("### ⚙️ Step 3 — Run MRP Validation")

        _persisted_mrp = st.session_state.get("mrp_files_persisted") or {}
        _scope_mrp     = st.session_state.get("mrp_export_scope", "All")
        _exec_mrp      = st.session_state.get("executor_name", "—")
        _has_zplm      = "zplm_marc" in _persisted_mrp

        st.markdown(f"""
        <div style="background:var(--app-bg-secondary);border:1px solid var(--app-border);border-radius:10px;padding:16px 20px;margin-bottom:16px">
        <div style="font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--app-fg-muted);margin-bottom:10px">
            ✅ Ready to validate — confirm before running
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;font-size:0.87rem;color:var(--app-fg)">
            <div><span style="color:var(--app-fg-muted)">Executor</span><br><strong>{_exec_mrp}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Scope</span><br><strong>{_scope_mrp}</strong></div>
            <div><span style="color:var(--app-fg-muted)">ZPLM MARC</span><br><strong>{"✅ Loaded — Material Transfer detection enabled" if _has_zplm else "⬜ Not loaded — Transfer detection disabled"}</strong></div>
        </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("▶️ Execute MRP Validation", type="primary", key="mrp_run"):
            files     = rehydrate_persisted_mapping("mrp_files_persisted")
            mara_file = files.get("mara")
            marc_file = files.get("marc")
            zplm_marc = files.get("zplm_marc")

            plant_filter_export  = st.session_state.mrp_plant_export
            plant_filter_display = st.session_state.mrp_plant_display

            with st.spinner("Validating MRP files..."):
                excel_display, errors_df, summary_df = run_mrp_validation(
                    mara_file=mara_file, marc_file=marc_file,
                    zplm_marc_file=zplm_marc, plant_filter=plant_filter_display,
                )
                if plant_filter_export != plant_filter_display:
                    excel_export, _, _ = run_mrp_validation(
                        mara_file=mara_file, marc_file=marc_file,
                        zplm_marc_file=zplm_marc, plant_filter=plant_filter_export,
                    )
                else:
                    excel_export = excel_display

            st.session_state.mrp_errors_df    = errors_df
            st.session_state.mrp_summary_df   = summary_df
            st.session_state.mrp_excel_export = excel_export

            if save_history:
                payload = build_mrp_log_payload(
                    user=effective_executor, plant_filter=plant_filter_export,
                    mara_file=mara_file, marc_file=marc_file, zplm_marc_file=zplm_marc,
                    errors_df=errors_df, summary_df=summary_df,
                )
                save_run_log(payload)

            st.session_state.mrp_step = 4; st.rerun()

    elif st.session_state.mrp_step == 4:
        st.markdown("### 📊 Step 4 — Results")

        errors_df    = st.session_state.mrp_errors_df
        summary_df   = st.session_state.mrp_summary_df
        excel_export = st.session_state.mrp_excel_export

        # FIX 4: count fields/materials consistently with Adjusted Summary
        # (Field Errors only — excludes Missing View errors)
        _fe_only = (errors_df[errors_df["Error_Type"] == "Field Error"]
                    if "Error_Type" in errors_df.columns else errors_df)
        _fe_cnt   = int(_fe_only.shape[0])
        _f_cnt    = int(_fe_only["Field"].nunique())    if "Field"    in _fe_only.columns else 0
        _m_cnt    = int(_fe_only["Material"].nunique()) if "Material" in _fe_only.columns else 0

        c = st.columns(3)
        with c[0]:
            st.markdown(f'<div class="card"><h4>Field Errors (Actionable)</h4><div class="big">{_fe_cnt:,}</div></div>', unsafe_allow_html=True)
        with c[1]:
            st.markdown(f'<div class="card"><h4>Distinct Fields</h4><div class="big">{_f_cnt}</div></div>', unsafe_allow_html=True)
        with c[2]:
            st.markdown(f'<div class="card"><h4>Materials with Errors</h4><div class="big">{_m_cnt:,}</div></div>', unsafe_allow_html=True)

        _top_mrp = _fe_only['Field'].value_counts().index[0] if (not _fe_only.empty and 'Field' in _fe_only.columns) else "—"
        _ic = "var(--app-bg-secondary)" if _fe_cnt > 500 else ("var(--app-bg-secondary)" if _fe_cnt > 100 else "var(--app-bg-secondary)")
        _ib = "var(--color-error)" if _fe_cnt > 500 else ("var(--color-warn)" if _fe_cnt > 100 else "var(--color-ok)")
        st.markdown(f"""
        <div style="background:{_ic};border-left:4px solid {_ib};border-radius:8px;
                    padding:12px 16px;margin-bottom:16px;font-size:0.9rem;line-height:1.6;color:var(--app-fg)">
            <strong>📋 MRP Validation Summary</strong><br>
            Found <strong>{_fe_cnt:,} MRP errors</strong> across <strong>{_m_cnt:,} materials</strong>
            in <strong>{_f_cnt} distinct fields</strong>.
            The most affected field is <strong>{_top_mrp}</strong>.
        </div>
        """, unsafe_allow_html=True)

        st.markdown("## 📈 Dashboard — MRP Validation")
        render_mrp_run_kpis(errors_df)
        render_mrp_run_charts(errors_df)

        with st.expander("📋 All Errors", expanded=False):
            _srch_mrp = st.text_input("🔍 Filter by Material, Field or Error", key="mrp_err_search")
            _df_mrp   = _coerce_df_for_display(errors_df)
            if _srch_mrp.strip():
                _mask_mrp = _df_mrp.apply(
                    lambda col: col.astype(str).str.contains(_srch_mrp.strip(), case=False, na=False)
                ).any(axis=1)
                _df_mrp = _df_mrp[_mask_mrp]
            st.caption(f"Showing {len(_df_mrp):,} of {len(errors_df):,} rows")
            st.dataframe(_df_mrp, use_container_width=True)

        with st.expander("📋 Summary", expanded=False):
            st.dataframe(_coerce_df_for_display(summary_df), use_container_width=True)

        render_history_with_download(
            "mrp", limit=10, label="MRP History",
            col_map={"When": "timestamp", "Who": "user",
                     "Plant": "plant_filter", "Errors": "metrics.total_errors"}
        )

        suffix = st.session_state.mrp_plant_export or "ALL"
        st.download_button(
            f"⬇️ Download MRP_Validation_Adjusted_Report_{suffix}.xlsx",
            data=excel_export.getvalue(),
            file_name=f"MRP_Validation_Adjusted_Report_{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Validation", key="mrp_new", use_container_width=True):
                st.session_state.mrp_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear uploads & start over", key="mrp_clear4", use_container_width=True):
                st.session_state["mrp_files_persisted"] = {}
                st.session_state.mrp_step = 1; st.rerun()

    render_footer(start_time, version="v1.4")

# ==========================================================
# PAGE 4 — SAP MATERIAL VALIDATION PROGRESS
# ==========================================================
elif main_choice == "📊 SAP Validation Progress":

    st.subheader("📊 SAP Validation Progress")
    effective_executor = identity_input("Who performs?", "progress_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    if "progress_step" not in st.session_state:
        st.session_state.progress_step = 1

    render_wizard(st.session_state.progress_step, WIZARD_STEPS_PROGRESS)

    if st.session_state.progress_step == 1:
        st.markdown("### 🧭 Step 1 — Choose Mode")
        mode = st.radio("Comparison Mode:", ["Auto", "Manual"], horizontal=True)
        st.session_state.progress_mode = mode

        if mode == "Auto":
            st.markdown("### 🌍 Scope (Previous Reports)")
            prog_scope = st.radio("Scope:", ["ALL", "PT01", "PT02"], horizontal=True)
            st.session_state.prog_scope = prog_scope

        if st.button("➡️ Continue to Files", key="prog_next1", use_container_width=True):
            st.session_state.progress_step = 2; st.rerun()

    elif st.session_state.progress_step == 2:
        mode = st.session_state.progress_mode
        st.markdown("### 📂 Step 2 — Files")

        if mode == "Auto":
            prog_scope = st.session_state.prog_scope
            curr_path, prev_path, avail = pick_last_two_reports(prog_scope)

            with st.expander("🔍 Reports found automatically", expanded=True):
                st.write(f"**Available for {prog_scope}:** {avail}")
                st.write(f"**Most Recent:** `{curr_path or '—'}`")
                st.write(f"**Previous:** `{prev_path or '—'}`")
                if avail == 1:
                    st.info("Only **1** report available. Need one more to compare.")
                elif avail == 0:
                    st.warning("No archived reports found.")

            st.session_state.progress_curr_path = curr_path
            st.session_state.progress_prev_path = prev_path
            disabled = not (curr_path and prev_path)

        else:
            st.markdown("Load **2 files** (any order). The newest will be detected automatically.")
            uploads = st.file_uploader(
                "Current + Previous (no renaming needed)",
                type="xlsx", accept_multiple_files=True, key="progress_manual_uploader",
            )
            curr_file, prev_file, notes = pick_current_previous_from_uploads(uploads or [])

            for note in notes:
                st.caption(note)

            with st.expander("📋 Files identified", expanded=True):
                st.markdown(f"{'✅' if curr_file else '❌'} **Current** — `{getattr(curr_file, 'name', 'Missing')}`")
                st.markdown(f"{'✅' if prev_file else '❌'} **Previous** — `{getattr(prev_file, 'name', 'Missing')}`")

            st.session_state.manual_curr = curr_file
            st.session_state.manual_prev = prev_file
            disabled = not (curr_file and prev_file)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="prog_back2", use_container_width=True):
                st.session_state.progress_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", disabled=disabled, key="prog_next2", use_container_width=True):
                st.session_state.progress_step = 3; st.rerun()

    elif st.session_state.progress_step == 3:
        st.markdown("### ⚙️ Step 3 — Compare Reports")

        if st.button("▶️ Execute Differences", type="primary", key="prog_run"):
            from sap_validation_engine import compare_adjusted_reports

            mode = st.session_state.progress_mode
            if mode == "Auto":
                curr_path = st.session_state.progress_curr_path
                prev_path = st.session_state.progress_prev_path
                with open(curr_path, "rb") as f:
                    curr_bytes = io.BytesIO(f.read())
                with open(prev_path, "rb") as f:
                    prev_bytes = io.BytesIO(f.read())
            else:
                curr_bytes = st.session_state.manual_curr
                prev_bytes = st.session_state.manual_prev

            with st.spinner("Comparing reports..."):
                excel_diff, diff_dfs = compare_adjusted_reports(curr_bytes, prev_bytes)

            st.session_state.progress_diff_excel    = excel_diff
            st.session_state.progress_diff_dfs      = diff_dfs
            st.session_state.progress_current_name  = (
                curr_path if mode == "Auto" else getattr(curr_bytes, "name", "current.xlsx")
            )
            st.session_state.progress_previous_name = (
                prev_path if mode == "Auto" else getattr(prev_bytes, "name", "previous.xlsx")
            )

            if save_history:
                payload = build_progress_log_payload(
                    user=effective_executor,
                    current_name=os.path.basename(st.session_state.progress_current_name),
                    previous_name=os.path.basename(st.session_state.progress_previous_name),
                    diff_dfs=diff_dfs,
                )
                save_run_log(payload)

            st.session_state.progress_step = 4; st.rerun()

    elif st.session_state.progress_step == 4:
        st.markdown("### 📊 Step 4 — Results")

        diff_dfs = st.session_state.progress_diff_dfs

        # Auto-insight
        _p_new  = len(diff_dfs.get("New Errors",      pd.DataFrame()))
        _p_res  = len(diff_dfs.get("Resolved Errors",  pd.DataFrame()))
        _p_chg  = len(diff_dfs.get("Changed Errors",   pd.DataFrame()))
        _p_mats = len(diff_dfs.get("New Materials",    pd.DataFrame()))
        _p_vl   = len(diff_dfs.get("Views Lost",       pd.DataFrame()))
        _p_net  = _p_new - _p_res
        _p_trend_icon  = "📈" if _p_net > 0 else ("📉" if _p_net < 0 else "➡️")
        _p_trend_label = "getting worse" if _p_net > 0 else ("improving" if _p_net < 0 else "stable")
        _p_border = "var(--color-error)" if _p_net > 0 else ("var(--color-ok)" if _p_net < 0 else "var(--color-warn)")
        st.markdown(f"""
        <div style="background:var(--app-bg-secondary);border-left:4px solid {_p_border};border-radius:8px;
                    padding:12px 16px;margin-bottom:16px;font-size:0.9rem;line-height:1.6;color:var(--app-fg)">
            <strong>{_p_trend_icon} Progress Summary</strong><br>
            <strong>{_p_new:,} new errors</strong> appeared, <strong>{_p_res:,} were resolved</strong>,
            and <strong>{_p_chg:,} changed</strong> their error message.
            Net change: <strong style="color:{_p_border}">{_p_net:+,} errors</strong> — trend is <strong>{_p_trend_label}</strong>.
            {"<br>" + str(_p_mats) + " new materials entered the report." if _p_mats else ""}
            {"<br><strong>" + str(_p_vl) + " views were lost</strong> since the previous run — these may need activation in SAP." if _p_vl else ""}
        </div>
        """, unsafe_allow_html=True)

        labels = ["New Errors", "Resolved Errors", "Changed Errors", "New Materials", "Views Lost"]
        tabs = st.tabs(labels)
        for tab, key in zip(tabs, labels):
            with tab:
                _df_prog = _coerce_df_for_display(diff_dfs.get(key, pd.DataFrame()))
                st.caption(f"{len(_df_prog):,} rows")
                _srch_prog = st.text_input("🔍 Filter", key=f"prog_search_{key.replace(' ','_')}")
                if _srch_prog.strip() and not _df_prog.empty:
                    _mask_prog = _df_prog.apply(
                        lambda col: col.astype(str).str.contains(_srch_prog.strip(), case=False, na=False)
                    ).any(axis=1)
                    _df_prog = _df_prog[_mask_prog]
                st.dataframe(_df_prog, use_container_width=True)

        st.markdown("## 📈 Dashboard — Progress")
        render_progress_kpis(diff_dfs)
        render_progress_charts(diff_dfs)

        st.markdown("---")
        st.download_button(
            "⬇️ Download SAP_Validation_Progress.xlsx",
            data=st.session_state.progress_diff_excel.getvalue(),
            file_name="SAP_Validation_Progress.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        render_history_with_download(
            "progress", limit=10, label="Progress History",
            col_map={"When": "timestamp", "Who": "user",
                     "New": "metrics.new_errors", "Resolved": "metrics.resolved_errors",
                     "Changed": "metrics.changed_errors"}
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Comparison", key="prog_new", use_container_width=True):
                st.session_state.progress_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear all & restart", key="prog_clear4", use_container_width=True):
                for k in ["progress_curr_path", "progress_prev_path",
                          "progress_manual_curr", "progress_manual_prev",
                          "progress_diff_dfs", "progress_diff_excel"]:
                    st.session_state.pop(k, None)
                st.session_state.progress_step = 1; st.rerun()

    render_footer(start_time, version="v1.4")

# ==========================================================
# PAGE 5 — DATA CLEANUPS
# ==========================================================
elif main_choice == "🧹 Data Cleanups":

    st.subheader("🧹 Data Cleanups Dashboard")
    effective_executor = identity_input("Who performs?", "dataclean_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    if "dataclean_step" not in st.session_state:
        st.session_state.dataclean_step = 1

    render_wizard(st.session_state.dataclean_step, WIZARD_STEPS_DEFAULT)

    required_keys = REQUIRED_KEYS["🧹 Data Cleanups"]

    if st.session_state.dataclean_step == 1:
        st.markdown("### 📂 Step 1 — Load Files")
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--accent);border-radius:6px;padding:12px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>📋 What files do you need? (all 7 required)</strong><br><br>
        <table style="border-collapse:collapse;width:100%">
        <tr><td style="padding:3px 8px">📄 <strong>MARA</strong></td><td style="color:var(--app-fg-muted)">General material data — Material Type, Old Material Number</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MARC</strong></td><td style="color:var(--app-fg-muted)">Plant-level data — PSMS status, Production Storage Location</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>ZPLM MARC</strong></td><td style="color:var(--app-fg-muted)">PLM plant data — Material Plant Type, EOP Date, Service Liability Date, LTB fields</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>ZPLM MARA</strong></td><td style="color:var(--app-fg-muted)">PLM general data — Enterproj Number, Product Hierarchy</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MBEW</strong></td><td style="color:var(--app-fg-muted)">Valuation data — Standard Price per plant</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>EKPO</strong></td><td style="color:var(--app-fg-muted)">Purchase order items — open POs to close validation</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MARD</strong></td><td style="color:var(--app-fg-muted)">Storage location stock — obsolete materials with no stock</td></tr>
        </table>
        </div>
        """, unsafe_allow_html=True)

        uploads = st.file_uploader("Drop all 7 Excel files here",
                                   type=["xlsx"], accept_multiple_files=True, key="dc_multi_uploader")
        files = classify_uploaded_files(uploads)
        if uploads:
            persist_uploaded_mapping(files, "dc_files_persisted")

        persisted = st.session_state.get("dc_files_persisted") or {}
        show_file_status(required_keys, files, persisted)

        _col_errs_dc = validate_uploaded_files(persisted, required_keys)
        if _col_errs_dc:
            for _k, _msg in _col_errs_dc.items():
                st.warning(f"⚠️ **{LABELS_PT.get(_k, _k)}** — {_msg}")

        disabled = any(k not in persisted for k in required_keys) or bool(_col_errs_dc)
        c1, c2 = st.columns(2)
        with c1:
            with st.expander("🧹 Clear"):
                if st.button("Clear stored files", key="dc_clear"):
                    st.session_state["dc_files_persisted"] = {}
                    st.success("Cleared."); st.rerun()
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, key="dc_next1", use_container_width=True):
                st.session_state.dataclean_step = 2; st.rerun()

    elif st.session_state.dataclean_step == 2:
        st.markdown("### 🗺️ Step 2 — Choose Plant Scope")
        st.caption("Choose the plant(s) to include in data cleanup validations.")

        export_scope = st.radio("Scope:", ["All Plants", "Only PT01", "Only PT02"], horizontal=True)
        plant_filter_export = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(export_scope)
        apply_filter_display = st.checkbox("Apply filter to displayed tables", value=False)
        plant_filter_display = plant_filter_export if apply_filter_display else None

        st.session_state.dc_export_scope  = export_scope
        st.session_state.dc_plant_export  = plant_filter_export
        st.session_state.dc_plant_display = plant_filter_display

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="dc_back2", use_container_width=True):
                st.session_state.dataclean_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="dc_next2", use_container_width=True):
                st.session_state.dataclean_step = 3; st.rerun()

    elif st.session_state.dataclean_step == 3:
        st.markdown("### ⚙️ Step 3 — Run Data Cleanups")

        _exec_dc    = st.session_state.get("executor_name", "—")
        _scope_dc   = st.session_state.get("dc_export_scope", "All Plants")
        _pers_dc    = st.session_state.get("dc_files_persisted") or {}
        st.markdown(f"""
        <div style="background:var(--app-bg-secondary);border:1px solid var(--app-border);border-radius:10px;padding:16px 20px;margin-bottom:16px">
        <div style="font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--app-fg-muted);margin-bottom:10px">
            ✅ Ready to validate — confirm before running
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;font-size:0.87rem;color:var(--app-fg)">
            <div><span style="color:var(--app-fg-muted)">Executor</span><br><strong>{_exec_dc}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Scope</span><br><strong>{_scope_dc}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Files loaded</span><br><strong>{len(_pers_dc)} of 7</strong></div>
        </div>
        <div style="margin-top:10px;font-size:0.8rem;color:var(--app-fg-muted)">
            Will run: FP Level Status · EOP vs Service Liability · Standard Cost · LTB fields ·
            EKPO open POs · MARD stock · Cross-existence MARA/MARC ↔ ZPLM
        </div>
        </div>
        """, unsafe_allow_html=True)

        files          = rehydrate_persisted_mapping("dc_files_persisted")
        mara_file      = files.get("mara")
        marc_file      = files.get("marc")
        zplm_marc_file = files.get("zplm_marc")
        zplm_mara_file = files.get("zplm_mara")
        mbew_file      = files.get("mbew")
        ekpo_file      = files.get("ekpo")
        mard_file      = files.get("mard")

        plant_filter_export  = st.session_state.dc_plant_export
        plant_filter_display = st.session_state.dc_plant_display

        if st.button("▶️ Run Data Cleanups", type="primary", key="dc_run"):
            with st.spinner("Running all Data Cleanup validations…"):
                excel_bytes, results_dict, kpis = run_data_cleanups_validation(
                    mara_file=mara_file, marc_file=marc_file,
                    zplm_marc_file=zplm_marc_file, zplm_mara_file=zplm_mara_file,
                    mbew_file=mbew_file, ekpo_file=ekpo_file, mard_file=mard_file,
                    plant_filter=plant_filter_display,
                )

            st.session_state.dc_excel   = excel_bytes
            st.session_state.dc_kpis    = kpis
            st.session_state.dc_results = results_dict

            if save_history:
                payload = build_dc_log_payload(
                    user=effective_executor, plant_filter=plant_filter_export,
                    mara_file=mara_file, marc_file=marc_file,
                    zplm_marc_file=zplm_marc_file, zplm_mara_file=zplm_mara_file,
                    mbew_file=mbew_file, ekpo_file=ekpo_file, mard_file=mard_file,
                    df_result=results_dict.get("FR_Status", pd.DataFrame()),
                    kpis=kpis,
                )
                save_run_log(payload)

            st.session_state.dataclean_step = 4; st.rerun()

    elif st.session_state.dataclean_step == 4:
        st.markdown("### 📊 Step 4 — Results")

        results_dict = st.session_state.get("dc_results", {})
        kpis         = st.session_state.get("dc_kpis", {})

        # KPI cards — consistent with other pages
        c = st.columns(4)
        with c[0]:
            st.markdown(f'<div class="card"><h4>Materials</h4><div class="big">{kpis.get("materials", 0)}</div></div>', unsafe_allow_html=True)
        with c[1]:
            st.markdown(f'<div class="card"><h4>Plants</h4><div class="big">{kpis.get("plants", 0)}</div></div>', unsafe_allow_html=True)
        with c[2]:
            st.markdown(f'<div class="card"><h4>Outdated</h4><div class="big">{kpis.get("outdated", 0)}</div></div>', unsafe_allow_html=True)
        with c[3]:
            st.markdown(f'<div class="card"><h4>Correct</h4><div class="big">{kpis.get("correct", 0)}</div></div>', unsafe_allow_html=True)

        if results_dict:
            _dc_tabs = st.tabs([f"📄 {name}" for name in results_dict.keys()])
            for _dc_tab, (_dc_name, _dc_df) in zip(_dc_tabs, results_dict.items()):
                with _dc_tab:
                    st.caption(f"{len(_dc_df):,} rows")
                    _srch_dc = st.text_input("🔍 Filter", key=f"dc_search_{_dc_name}")
                    _df_dc = _coerce_df_for_display(_dc_df)
                    if _srch_dc.strip() and not _df_dc.empty:
                        _mask_dc = _df_dc.apply(
                            lambda col: col.astype(str).str.contains(_srch_dc.strip(), case=False, na=False)
                        ).any(axis=1)
                        _df_dc = _df_dc[_mask_dc]
                    st.dataframe(_df_dc, use_container_width=True)

        excel_bytes = st.session_state.get("dc_excel")
        st.download_button(
            "⬇️ Download Data_Cleanups_Report.xlsx",
            data=excel_bytes.getvalue() if excel_bytes else b"",
            file_name="Data_Cleanups_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Run", key="dc_new", use_container_width=True):
                st.session_state.dataclean_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear & Restart", key="dc_clear4", use_container_width=True):
                st.session_state["dc_files_persisted"] = {}
                st.session_state.dataclean_step = 1; st.rerun()



# ==========================================================
# PAGE 6 — SSA VALIDATION
# ==========================================================
elif main_choice == "📋 SSA Validation":

    st.subheader("📋 Schedule Agreement (SSA) Validation")
    effective_executor = identity_input("Who performs?", "ssa_exec")

    if "ssa_step" not in st.session_state:
        st.session_state.ssa_step = 1

    render_wizard(st.session_state.ssa_step, WIZARD_STEPS_DEFAULT)

    ssa_required = ["vbak", "vbap", "vbkd", "kna1"]
    ssa_optional = ["marc", "vbpa"]

    if st.session_state.ssa_step == 1:
        st.markdown("### 📂 Step 1 — Load Files")
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--accent);border-radius:6px;padding:12px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>📋 What files do you need?</strong><br><br>
        <table style="border-collapse:collapse;width:100%">
        <tr><td style="padding:3px 8px">📄 <strong>VBAK</strong></td><td style="color:var(--app-fg-muted)">SA Header — Customer Groups, Overall Status</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>VBAP</strong></td><td style="color:var(--app-fg-muted)">SA Items — Plant, Route, IFS ID, Supplier Code, Material</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>VBKD</strong></td><td style="color:var(--app-fg-muted)">SA Business data — Incoterms</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>KNA1</strong></td><td style="color:var(--app-fg-muted)">Customer master — Country (GB check)</td></tr>
        <tr><td style="padding:3px 8px;color:var(--color-ok)">🔓 <strong>MARC</strong></td><td style="color:var(--app-fg-muted)">Optional — Material Status 70/75 + Country of Origin checks</td></tr>
        <tr><td style="padding:3px 8px;color:var(--color-ok)">🔓 <strong>VBPA</strong></td><td style="color:var(--app-fg-muted)">Optional — Partner country override for DDP+GB rule</td></tr>
        </table>
        </div>
        """, unsafe_allow_html=True)

        uploads = st.file_uploader(
            "Drop VBAK, VBAP, VBKD, KNA1 + optional MARC and VBPA",
            type="xlsx", accept_multiple_files=True, key="ssa_uploader",
        )

        def classify_ssa_files(files):
            mapping = {k: None for k in ["vbak","vbap","vbkd","vbpa","kna1","marc"]}
            for f in (files or []):
                name = f.name.lower()
                if "vbak" in name:   mapping["vbak"] = f
                elif "vbap" in name: mapping["vbap"] = f
                elif "vbkd" in name: mapping["vbkd"] = f
                elif "vbpa" in name: mapping["vbpa"] = f
                elif "kna1" in name: mapping["kna1"] = f
                elif "marc" in name: mapping["marc"] = f
            return mapping

        files_map = classify_ssa_files(uploads)
        if uploads:
            persist_uploaded_mapping(files_map, "ssa_files_persisted")

        persisted = st.session_state.get("ssa_files_persisted") or {}

        with st.expander("📋 File Status", expanded=True):
            label_map = {
                "vbak": "VBAK.xlsx", "vbap": "VBAP.xlsx",
                "vbkd": "VBKD.xlsx", "kna1": "KNA1.xlsx",
                "marc": "MARC.xlsx (optional)", "vbpa": "VBPA.xlsx (optional)"
            }
            for key in ["vbak","vbap","vbkd","kna1","marc","vbpa"]:
                in_up   = files_map.get(key) is not None
                in_pers = key in persisted
                if in_up:
                    icon   = "✅"
                    detail = getattr(files_map[key], "name", "")
                elif in_pers:
                    icon   = "💾"
                    detail = persisted[key].get("name", "")
                elif key in ssa_optional:
                    icon   = "⬜"
                    detail = "Not uploaded (optional)"
                else:
                    icon   = "❌"
                    detail = "Missing"
                st.markdown(f"{icon} **{label_map[key]}** — `{detail}`")

        disabled = any(k not in persisted for k in ssa_required)
        c1, c2 = st.columns(2)
        with c1:
            with st.expander("🧹 Clear"):
                if st.button("Clear SSA uploads", key="ssa_clear1"):
                    st.session_state["ssa_files_persisted"] = {}
                    st.success("Cleared."); st.rerun()
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, key="ssa_next1", use_container_width=True):
                st.session_state.ssa_step = 2; st.rerun()

    elif st.session_state.ssa_step == 2:
        st.markdown("### 🌍 Step 2 — Choose Scope")
        ssa_scope = st.radio("Filter by Plant:", ["All Plants", "Only PT01", "Only PT02"], horizontal=True)
        st.session_state.ssa_plant_filter = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(ssa_scope)
        st.session_state.ssa_scope = ssa_scope

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="ssa_back2", use_container_width=True):
                st.session_state.ssa_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue", key="ssa_next2", use_container_width=True):
                st.session_state.ssa_step = 3; st.rerun()

    elif st.session_state.ssa_step == 3:
        st.markdown("### ⚙️ Step 3 — Run SSA Validation")

        _exec_ssa  = st.session_state.get("executor_name", "—")
        _scope_ssa = st.session_state.get("ssa_scope", "All Plants")
        _pers_ssa  = st.session_state.get("ssa_files_persisted") or {}
        _has_marc  = "marc" in _pers_ssa
        _has_vbpa  = "vbpa" in _pers_ssa

        st.markdown(f"""
        <div style="background:var(--app-bg-secondary);border:1px solid var(--app-border);border-radius:10px;padding:16px 20px;margin-bottom:16px">
        <div style="font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--app-fg-muted);margin-bottom:10px">
            ✅ Ready to validate
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;font-size:0.87rem;color:var(--app-fg)">
            <div><span style="color:var(--app-fg-muted)">Executor</span><br><strong>{_exec_ssa}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Scope</span><br><strong>{_scope_ssa}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Optional files</span><br><strong>{"MARC ✅" if _has_marc else "MARC ⬜"} &nbsp; {"VBPA ✅" if _has_vbpa else "VBPA ⬜"}</strong></div>
        </div>
        <div style="margin-top:10px;font-size:0.8rem;color:var(--app-fg-muted)">
            Rules: Mandatory Fields · Incoterms DDP+GB · Incompletion Log{"· Rejected Status 70/75 · Country of Origin" if _has_marc else ""}
        </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("▶️ Run SSA Validation", type="primary", key="ssa_run"):
            files = rehydrate_persisted_mapping("ssa_files_persisted")
            with st.spinner("Running SSA validations..."):
                excel_bytes, results_dict, summary_df = run_ssa_validation(
                    vbak_file=files.get("vbak"),
                    vbap_file=files.get("vbap"),
                    vbkd_file=files.get("vbkd"),
                    kna1_file=files.get("kna1"),
                    marc_file=files.get("marc"),
                    vbpa_file=files.get("vbpa"),
                    plant_filter=st.session_state.get("ssa_plant_filter"),
                )
            st.session_state.ssa_excel   = excel_bytes
            st.session_state.ssa_results = results_dict
            st.session_state.ssa_summary = summary_df
            st.session_state.ssa_step = 4; st.rerun()

    elif st.session_state.ssa_step == 4:
        st.markdown("### 📊 Step 4 — Results")

        results_dict = st.session_state.get("ssa_results", {})
        summary_df   = st.session_state.get("ssa_summary", pd.DataFrame())
        excel_bytes  = st.session_state.get("ssa_excel")

        total_errors = sum(len(v) for v in results_dict.values())
        total_rules  = len(results_dict)
        total_docs   = 0
        if results_dict:
            _all = pd.concat(list(results_dict.values()), ignore_index=True)
            if "Sales Document" in _all.columns:
                total_docs = int(_all["Sales Document"].nunique())

        c = st.columns(3)
        with c[0]:
            st.markdown(f'<div class="card"><h4>Total Errors</h4><div class="big">{total_errors:,}</div></div>', unsafe_allow_html=True)
        with c[1]:
            st.markdown(f'<div class="card"><h4>Rules with Errors</h4><div class="big">{total_rules}</div></div>', unsafe_allow_html=True)
        with c[2]:
            st.markdown(f'<div class="card"><h4>Documents Affected</h4><div class="big">{total_docs:,}</div></div>', unsafe_allow_html=True)

        if results_dict:
            tabs = st.tabs([f"📄 {name}" for name in results_dict.keys()])
            for tab, (name, df) in zip(tabs, results_dict.items()):
                with tab:
                    st.caption(f"{len(df):,} errors")
                    search = st.text_input("🔍 Filter", key=f"ssa_search_{name.replace(' ','_').replace('/','_')}")
                    df_show = _coerce_df_for_display(df)
                    if search.strip() and not df_show.empty:
                        mask = df_show.apply(lambda col: col.astype(str).str.contains(search.strip(), case=False, na=False)).any(axis=1)
                        df_show = df_show[mask]
                    st.dataframe(df_show, use_container_width=True)

        with st.expander("📋 Summary", expanded=False):
            st.dataframe(_coerce_df_for_display(summary_df), use_container_width=True)

        if excel_bytes:
            _scope_label = st.session_state.get("ssa_scope","ALL").replace(" ","_")
            st.download_button(
                "⬇️ Download SSA_Validation_Report.xlsx",
                data=excel_bytes.getvalue(),
                file_name=f"SSA_Validation_Report_{_scope_label}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Validation", key="ssa_new", use_container_width=True):
                st.session_state.ssa_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear & Restart", key="ssa_clear4", use_container_width=True):
                st.session_state["ssa_files_persisted"] = {}
                st.session_state.ssa_step = 1; st.rerun()

    render_footer(start_time, version="v1.3")
    render_footer(start_time, version="v1.4")