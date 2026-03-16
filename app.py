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
)

# ==========================================================
# CONFIGURAÇÃO INICIAL
# ==========================================================
st.set_page_config(page_title="SAP & ATL Validation Hub", layout="wide")
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
theme_name = render_header("🧪 Validation Hub — SAP & ATL", show_theme_picker=True)

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
    ("sap",         lambda n: any(k in n for k in ["sapui5", "item", "browse", "report"])),
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
    ("timing",      lambda n: "timing" in n and "letter" in n and "timingletter" not in n and "timing_letter" not in n),
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

# Keys obrigatórias por página
REQUIRED_KEYS: dict[str, list[str]] = {
    "📦 SAP Material Validation":         ["sap", "mara", "marc", "mvke", "zplm_marc", "zplm_mara", "qmat"],
    "🧮 MRP Validation":                  ["mara", "marc"],
    "Automotive + Timing Letter (A & TL)":["auto", "plm", "timing"],
    "📈 SAP Material Validation Progress": [],   # gerido pelo modo Auto/Manual
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

def colorize_errors(df):
    def highlight_row(row):
        error_text = str(row.get("Error", "")).lower()
        if "mandatory" in error_text or "obrigatório" in error_text:
            return ["background-color: #ffcccc"] * len(row)
        if "invalid" in error_text or "inválido" in error_text:
            return ["background-color: #ffe5b4"] * len(row)
        if "view" in error_text and "missing" in error_text:
            return ["background-color: #fff7a8"] * len(row)
        return [""] * len(row)
    return df.style.apply(highlight_row, axis=1)

# ==========================================================
# IDENTIDADE
# ==========================================================
if "executor_name" not in st.session_state:
    st.session_state["executor_name"] = ""

def identity_input(label: str, key: str) -> str:
    executor_name = st.text_input(
        label,
        value=st.session_state["executor_name"],
        placeholder="CDS ID...",
        key=key,
    )
    st.session_state["executor_name"] = executor_name.strip()
    if not st.session_state["executor_name"]:
        st.warning("⚠ Please enter your CDS ID to continue.")
        st.stop()
    return st.session_state["executor_name"]

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
        "📈 SAP Material Validation Progress",
        "🧹 Data Cleanups",
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

# ==========================================================
# 📈 EVOLUTION DASHBOARD
# ==========================================================
if main_choice == "📈 Evolution Dashboard":
    render_evolution_dashboard()
    render_footer(start_time, version="v1.2")

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

        disabled = any(k not in persisted for k in required_keys)
        c1, c2 = st.columns(2)
        with c1:
            with st.expander("🧹 Clear uploaded files"):
                if st.button("Clear persistent uploads", key="sap_clear"):
                    st.session_state["sap_files_persisted"] = {}
                    st.success("Cleared.")
                    st.rerun()
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, key="sap_next1", width='stretch'):
                st.session_state.sap_step = 2
                st.rerun()

    # --- STEP 2: SCOPE ---
    elif st.session_state.sap_step == 2:
        st.markdown("### 🌍 Step 2 — Select Scope")

        export_scope = st.radio("Scope of download:", ["All", "Only PT01", "Only PT02"], horizontal=True)
        plant_filter_export = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(export_scope)

        apply_filter_display = st.checkbox("Apply filter to displayed tables", value=False)
        plant_filter_display = plant_filter_export if apply_filter_display else None

        st.session_state.sap_export_scope    = export_scope
        st.session_state.sap_plant_export    = plant_filter_export
        st.session_state.sap_plant_display   = plant_filter_display

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="sap_back2", width='stretch'):
                st.session_state.sap_step = 1
                st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="sap_next2", width='stretch'):
                st.session_state.sap_step = 3
                st.rerun()

    # --- STEP 3: VALIDATE ---
    elif st.session_state.sap_step == 3:
        st.markdown("### ⚙️ Step 3 — Run Validation")
        st.info("Click the button below to validate the uploaded files.")

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

            progress_bar = st.progress(0, text="Initializing validation...")

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

                # Exportação com filtro diferente só se necessário
                if plant_filter_export != plant_filter_display:
                    excel_export, _, _ = run_full_validation(
                        sap_file=sap_file, mara_file=mara_file, marc_file=marc_file,
                        mvke_file=mvke_file, qmat_file=qmat_file,
                        zplm_marc_file=zplm_marc, zplm_mara_file=zplm_mara,
                        previous_report=None, plant_filter=plant_filter_export,
                        csr_item_file=csr_item, csr_head_file=csr_head,
                    )
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

            st.session_state.sap_step = 4
            st.rerun()

    # --- STEP 4: RESULTS ---
    elif st.session_state.sap_step == 4:
        st.markdown("### 📊 Step 4 — Results")

        errors_df    = st.session_state.sap_errors_df
        summary_df   = st.session_state.sap_summary_df
        excel_export = st.session_state.sap_excel_export
        report_path  = st.session_state.sap_report_path

        c = st.columns(3)
        with c[0]:
            st.markdown(f'<div class="card"><h4>Total Errors</h4><div class="big">{errors_df.shape[0]}</div></div>', unsafe_allow_html=True)
        with c[1]:
            total_fields = int(errors_df["Field"].nunique()) if "Field" in errors_df.columns else 0
            st.markdown(f'<div class="card"><h4>Distinct Fields</h4><div class="big">{total_fields}</div></div>', unsafe_allow_html=True)
        with c[2]:
            total_mats = int(errors_df["Material"].nunique()) if "Material" in errors_df.columns else 0
            st.markdown(f'<div class="card"><h4>Materials with Errors</h4><div class="big">{total_mats}</div></div>', unsafe_allow_html=True)

        st.header("All Errors")
        st.dataframe(_coerce_df_for_display(errors_df), width='stretch')
        st.header("Summary")
        st.dataframe(_coerce_df_for_display(summary_df), width='stretch')

        st.markdown("## 📈 Dashboard — SAP Material Validation")
        render_sap_run_kpis(errors_df)
        render_sap_run_charts(errors_df)

        download_name = os.path.basename(report_path) if report_path else "SAP_Validation_Adjusted_Report_ALL.xlsx"
        st.download_button("⬇️ Download Excel", data=excel_export.getvalue(),
                           file_name=download_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with st.expander("🕘 SAP History (last 10 runs)"):
            logs = list_run_logs("sap", limit=10)
            if logs:
                st.dataframe(pd.DataFrame([
                    {"When": l.get("timestamp"), "Who": l.get("user"),
                     "Filter": l.get("plant_filter"), "Errors": l.get("metrics", {}).get("total_errors", 0)}
                    for l in logs
                ]), width='stretch')
            else:
                st.write("No history.")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔁 New Validation", key="sap_new", width='stretch'):
                st.session_state.sap_step = 1
                st.rerun()
        with c2:
            if st.button("🧹 Clear uploads & start over", key="sap_clear4", width='stretch'):
                st.session_state["sap_files_persisted"] = {}
                st.session_state.sap_step = 1
                st.rerun()

    render_footer(start_time, version="v1.2")

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

        uploads = st.file_uploader(
            "Drop Automotive + TimingLetter + Timing Letter files",
            type="xlsx", accept_multiple_files=True, key="atl_multi_uploader",
        )
        files = classify_uploaded_files(uploads)
        if uploads:
            persist_uploaded_mapping(files, "atl_files_persisted")

        persisted = st.session_state.get("atl_files_persisted") or {}
        show_file_status(required_keys, files, persisted)

        disabled = any(k not in persisted for k in required_keys)
        c1, c2 = st.columns(2)
        with c1:
            with st.expander("🧹 Clear persistent uploads"):
                if st.button("Clear ATL uploads", key="atl_clear"):
                    st.session_state["atl_files_persisted"] = {}
                    st.success("Cleared.")
                    st.rerun()
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, key="atl_next1", width='stretch'):
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
            if st.button("⬅️ Back", key="atl_back2", width='stretch'):
                st.session_state.atl_step = 1
                st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="atl_next2", width='stretch'):
                st.session_state.atl_step = 3
                st.rerun()

    # --- STEP 3 ---
    elif st.session_state.atl_step == 3:
        st.markdown("### ⚙️ Step 3 — Run ATL Workflow")

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

        if scope == "PT01":
            st.header("PT01"); st.dataframe(df_pt01, width='stretch')
        elif scope == "PT02":
            st.header("PT02"); st.dataframe(df_pt02, width='stretch')
        else:
            st.header("PT01"); st.dataframe(df_pt01, width='stretch')
            st.header("PT02"); st.dataframe(df_pt02, width='stretch')

        st.markdown("## 📈 Dashboard — PLM Workflow (ATL)")
        render_atl_run_kpis(df_pt01, df_pt02)
        render_atl_run_charts(df_pt01, df_pt02)

        suffix = st.session_state.atl_filter_export or "ALL"
        st.download_button(
            f"⬇️ Download PLM_Workflow_{suffix}.xlsx",
            data=excel_all.getvalue(), file_name=f"PLM_Workflow_{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("🕑 ATL History (last 10 runs)"):
            logs = list_run_logs("atl", limit=10)
            if logs:
                st.dataframe(pd.DataFrame([
                    {"When": l["timestamp"], "Who": l["user"], "Scope": l["plant_scope"],
                     "PT01 rows": l["metrics"]["PT01_rows"], "PT02 rows": l["metrics"]["PT02_rows"]}
                    for l in logs
                ]), width='stretch')
            else:
                st.write("No history.")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Run ATL", key="atl_new", width='stretch'):
                st.session_state.atl_step = 1
                st.rerun()
        with c2:
            if st.button("🧹 Clear uploads & start over", key="atl_clear4", width='stretch'):
                st.session_state["atl_files_persisted"] = {}
                st.session_state.atl_step = 1
                st.rerun()

    render_footer(start_time, version="v1.2")

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
            if st.button("➡️ Continue to Scope", key="cc_next1", width='stretch'):
                st.session_state.cc_step = 2
                st.rerun()

    elif st.session_state.cc_step == 2:
        st.markdown("### 🌍 Step 2 — Select Scope")
        cc_scope = st.radio("Scope:", ["All", "Only PT01", "Only PT02"], horizontal=True)
        st.session_state.cc_scope = cc_scope
        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="cc_back2", width='stretch'):
                st.session_state.cc_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="cc_next2", width='stretch'):
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
            if st.button("🔄 New Run", key="cc_new", width='stretch'):
                st.session_state.cc_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear & Start Over", key="cc_clear4", width='stretch'):
                for k in ["cc_result_df", "cc_excel"]:
                    st.session_state.pop(k, None)
                st.session_state.cc_step = 1; st.rerun()

    render_footer(start_time, version="v1.2")

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
        uploads = st.file_uploader("Drop MARA + MARC + (optional) ZPLM MARC",
                                   type="xlsx", accept_multiple_files=True, key="mrp_multi_uploader")
        files = classify_uploaded_files(uploads)
        if uploads:
            persist_uploaded_mapping(files, "mrp_files_persisted")

        persisted = st.session_state.get("mrp_files_persisted") or {}
        show_file_status(required_keys, files, persisted)

        disabled = any(k not in persisted for k in required_keys)
        c1, c2 = st.columns(2)
        with c1:
            with st.expander("🧹 Clear persistent uploads"):
                if st.button("Clear MRP uploads", key="mrp_clear"):
                    st.session_state["mrp_files_persisted"] = {}
                    st.success("Cleared."); st.rerun()
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, key="mrp_next1", width='stretch'):
                st.session_state.mrp_step = 2; st.rerun()

    elif st.session_state.mrp_step == 2:
        st.markdown("### 🌍 Step 2 — Choose Scope")
        export_scope = st.radio("Scope:", ["All", "Only PT01", "Only PT02"], horizontal=True)
        plant_filter_export = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(export_scope)
        apply_filter_display = st.checkbox("Apply filter to displayed tables", value=False)
        plant_filter_display = plant_filter_export if apply_filter_display else None

        st.session_state.mrp_export_scope  = export_scope
        st.session_state.mrp_plant_export  = plant_filter_export
        st.session_state.mrp_plant_display = plant_filter_display

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="mrp_back2", width='stretch'):
                st.session_state.mrp_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="mrp_next2", width='stretch'):
                st.session_state.mrp_step = 3; st.rerun()

    elif st.session_state.mrp_step == 3:
        st.markdown("### ⚙️ Step 3 — Run MRP Validation")
        st.info("Run all MRP Validation files.")

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

        c = st.columns(3)
        with c[0]:
            st.markdown(f'<div class="card"><h4>Total Errors</h4><div class="big">{errors_df.shape[0]}</div></div>', unsafe_allow_html=True)
        with c[1]:
            total_fields = int(errors_df["Field"].nunique()) if "Field" in errors_df.columns else 0
            st.markdown(f'<div class="card"><h4>Distinct Fields</h4><div class="big">{total_fields}</div></div>', unsafe_allow_html=True)
        with c[2]:
            total_mats = int(errors_df["Material"].nunique()) if "Material" in errors_df.columns else 0
            st.markdown(f'<div class="card"><h4>Materials with Errors</h4><div class="big">{total_mats}</div></div>', unsafe_allow_html=True)

        st.header("All Errors")
        st.dataframe(_coerce_df_for_display(errors_df), width='stretch')
        st.header("Summary")
        st.dataframe(_coerce_df_for_display(summary_df), width='stretch')

        st.markdown("## 📈 Dashboard — MRP Validation")
        render_mrp_run_kpis(errors_df)
        render_mrp_run_charts(errors_df)

        with st.expander("🕑 MRP History (last 10 runs)"):
            logs = list_run_logs("mrp", limit=10)
            if logs:
                st.dataframe(pd.DataFrame([
                    {"When": l["timestamp"], "Who": l["user"],
                     "Filter": l["plant_filter"], "Errors": l["metrics"]["total_errors"]}
                    for l in logs
                ]), width='stretch')
            else:
                st.write("No history.")

        suffix = st.session_state.mrp_plant_export or "ALL"
        st.download_button(
            f"⬇️ Download MRP_Validation_Adjusted_Report_{suffix}.xlsx",
            data=excel_export.getvalue(),
            file_name=f"MRP_Validation_Adjusted_Report_{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Validation", key="mrp_new", width='stretch'):
                st.session_state.mrp_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear uploads & start over", key="mrp_clear4", width='stretch'):
                st.session_state["mrp_files_persisted"] = {}
                st.session_state.mrp_step = 1; st.rerun()

    render_footer(start_time, version="v1.2")

# ==========================================================
# PAGE 4 — SAP MATERIAL VALIDATION PROGRESS
# ==========================================================
elif main_choice == "📈 SAP Material Validation Progress":

    st.subheader("📈 SAP Material Validation Progress")
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

        if st.button("➡️ Continue to Files", key="prog_next1", width='stretch'):
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
            if st.button("⬅️ Back", key="prog_back2", width='stretch'):
                st.session_state.progress_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", disabled=disabled, key="prog_next2", width='stretch'):
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
        labels   = ["New Errors", "Resolved Errors", "Changed Errors", "New Materials", "Views Lost"]

        tabs = st.tabs(labels)
        for tab, key in zip(tabs, labels):
            with tab:
                st.dataframe(_coerce_df_for_display(diff_dfs.get(key, pd.DataFrame())), width='stretch')

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

        with st.expander("🕑 Progress History (last 10 runs)"):
            logs = list_run_logs("progress", limit=10)
            if logs:
                st.dataframe(pd.DataFrame([
                    {"When": l["timestamp"], "Who": l["user"],
                     "New": l["metrics"]["new_errors"], "Resolved": l["metrics"]["resolved_errors"],
                     "Changed": l["metrics"]["changed_errors"], "New Mats": l["metrics"]["new_materials"],
                     "Views Lost": l["metrics"]["views_lost"]}
                    for l in logs
                ]), width='stretch')
            else:
                st.write("No history.")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Comparison", key="prog_new", width='stretch'):
                st.session_state.progress_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear all & restart", key="prog_clear4", width='stretch'):
                for k in ["progress_curr_path", "progress_prev_path",
                          "progress_manual_curr", "progress_manual_prev",
                          "progress_diff_dfs", "progress_diff_excel"]:
                    st.session_state.pop(k, None)
                st.session_state.progress_step = 1; st.rerun()

    render_footer(start_time, version="v1.2")

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
        st.caption("Upload **MARA + MARC + ZPLM MARA + ZPLM MARC + MBEW + EKPO + MARD** (all required).")

        uploads = st.file_uploader("Drop all 7 Excel files here",
                                   type=["xlsx"], accept_multiple_files=True, key="dc_multi_uploader")
        files = classify_uploaded_files(uploads)
        if uploads:
            persist_uploaded_mapping(files, "dc_files_persisted")

        persisted = st.session_state.get("dc_files_persisted") or {}
        show_file_status(required_keys, files, persisted)

        disabled = any(k not in persisted for k in required_keys)
        c1, c2 = st.columns(2)
        with c1:
            with st.expander("🧹 Clear"):
                if st.button("Clear stored files", key="dc_clear"):
                    st.session_state["dc_files_persisted"] = {}
                    st.success("Cleared."); st.rerun()
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, key="dc_next1", width='stretch'):
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
            if st.button("⬅️ Back", key="dc_back2", width='stretch'):
                st.session_state.dataclean_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="dc_next2", width='stretch'):
                st.session_state.dataclean_step = 3; st.rerun()

    elif st.session_state.dataclean_step == 3:
        st.markdown("### ⚙️ Step 3 — Run Data Cleanups")

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

        for sheet_name, df_val in results_dict.items():
            st.header(f"📄 {sheet_name}")
            st.dataframe(_coerce_df_for_display(df_val), width='stretch')

        excel_bytes = st.session_state.get("dc_excel")
        st.download_button(
            "⬇️ Download Data_Cleanups_Report.xlsx",
            data=excel_bytes.getvalue() if excel_bytes else b"",
            file_name="Data_Cleanups_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Run", key="dc_new", width='stretch'):
                st.session_state.dataclean_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear & Restart", key="dc_clear4", width='stretch'):
                st.session_state["dc_files_persisted"] = {}
                st.session_state.dataclean_step = 1; st.rerun()

    render_footer(start_time, version="v1.2")