# app.py
import streamlit as st
import pandas as pd
import re
import io
import os
import time

from utils.dashboard import (
    render_sap_run_kpis, render_sap_run_charts, render_sap_history_charts,
    render_atl_run_kpis, render_atl_run_charts, render_atl_history_charts,
    render_progress_kpis, render_progress_charts, render_mrp_history_charts,
    render_mrp_run_kpis, render_mrp_run_charts
)
from utils.identify import get_executor_identity
from sap_validation_engine import (
    run_full_validation,
    run_plm_workflow_combined,
    _apply_conditional_formatting_in_memory,
    build_sap_log_payload,
    build_atl_log_payload,
    build_progress_log_payload,
    save_run_log,
    list_run_logs,
    run_mrp_validation,
    build_mrp_log_payload
)
from utils.layout import render_footer, render_header, inject_theme_overrides
from datetime import datetime
# ==========================================================
# CSS GLOBAL INJECTION
# ==========================================================
from pathlib import Path
import streamlit as st

def inject_css(css_path: str):
    """
    Injeta um ficheiro CSS externo no Streamlit.
    Caminho relativo ao diretório onde está o app.py.
    """
    try:
        base = Path(__file__).resolve().parent
        css_file = base / css_path

        if css_file.exists():
            css = css_file.read_text(encoding="utf-8")
            st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
        else:
            st.warning(f"⚠️ CSS não encontrado: {css_file}")
    except Exception as e:
        st.error(f"Erro ao carregar CSS: {e}")

# ==========================================================
# CONFIGURAÇÃO INICIAL
# ==========================================================
st.set_page_config(page_title="SAP & ATL Validation Hub", layout="wide")
st.title("🧰 Validation Hub")
inject_css("assets/styles.css")
start_time = time.time()

# Header global (aparece em todas as páginas)
theme_name = render_header("🧪 Validation Hub — SAP & ATL", show_theme_picker=True)

# ==========================================================
# HELPERS PARA MODO MANUAL (Progress) – DETEÇÃO DE "Atual" e "Anterior"
# ==========================================================
_TS_PATTERN = re.compile(r'(\d{8}_\d{6})')  # ex.: 20260114_085435

def extract_timestamp_from_name(filename: str) -> datetime | None:
    if not filename:
        return None
    m = _TS_PATTERN.search(filename)
    if not m:
        return None
    raw = m.group(1)  # 'YYYYMMDD_HHMMSS'
    try:
        return datetime.strptime(raw, "%Y%m%d_%H%M%S")
    except Exception:
        return None

# ==========================================================
# Helper: escolher (Atual, Anterior) a partir de uploads (modo Manual)
# Usa extract_timestamp_from_name para ordenar por YYYYMMDD_HHMMSS
# Retorna: (curr_file, prev_file, diagnostics: list[str])
# ==========================================================
def pick_current_previous_from_uploads(files: list) -> tuple:
    """
    Recebe a lista de UploadedFile (qualquer nome).
    - Extrai timestamps do nome
    - Ordena desc por timestamp
    - Seleciona (Atual, Anterior)
    Retorna (curr_file, prev_file, diagnostics:list[str])
    diagnostics: mensagens para mostrar no UI (ex.: ficheiros sem timestamp).
    """
    diagnostics = []
    candidates = []

    for f in files or []:
        ts = extract_timestamp_from_name(getattr(f, "name", ""))
        if ts is None:
            diagnostics.append(f"⚠ Sem timestamp reconhecido: {getattr(f, 'name', '')}")
        candidates.append((ts, f))

    # mantém os que têm timestamp primeiro; em seguida, ordena por ts desc
    with_ts = [(ts, f) for ts, f in candidates if ts is not None]
    without_ts = [(ts, f) for ts, f in candidates if ts is None]

    with_ts.sort(key=lambda x: x[0], reverse=True)  # mais recente primeiro

    ordered = [f for _, f in with_ts] + [f for _, f in without_ts]

    curr = ordered[0] if len(ordered) >= 1 else None
    prev = ordered[1] if len(ordered) >= 2 else None

    if len(with_ts) < 2:
        diagnostics.append("ℹ️ Menos de dois ficheiros com timestamp válido; separei pela ordem de upload para completar os 2.")

    return curr, prev, diagnostics


# ==========================================================
# Helper: escolher os dois últimos relatórios arquivados (ALL/PT01/PT02)
# Usado pela Página 4 (Progress) em modo "Auto (arquivado)"
# ==========================================================
def pick_last_two_reports(plant_scope: str):
    """
    Prioridade:
    1) usar os logs 'sap' com report_path (filtrados por scope)
    2) fallback: varrer ./reports/sap/** pelo sufixo e usar dois mais recentes
    Retorna: (curr_path, prev_path, count)
    """
    import os, glob
    scope = (plant_scope or "ALL").upper()

    # 1) Tenta pelos logs
    logs = list_run_logs("sap", limit=200) or []
    cand = []
    for log in logs:
        if str(log.get("plant_filter", "ALL")).upper() != scope:
            continue
        path = log.get("report_path")
        if path and os.path.exists(path):
            # timestamp dos logs é suficiente para ordenar por ordem de criação
            cand.append((log.get("timestamp", ""), os.path.abspath(path)))

    if len(cand) >= 2:
        cand.sort(key=lambda x: x[0], reverse=True)
        return cand[0][1], cand[1][1], len(cand)

    # 2) Fallback: varrer o vault
    base = os.path.join(".", "reports", "sap")
    month_dirs = [d for d in (os.listdir(base) if os.path.exists(base) else [])]
    paths = []
    for m in month_dirs:
        paths.extend(
            glob.glob(os.path.join(base, m, f"*SAP_Validation_Adjusted_Report_{scope}.xlsx"))
        )
    # Ordenar por timestamp embutido no nome (YYYYMMDD_HHMMSS_...)
    paths = sorted(paths, reverse=True)

    if len(paths) >= 2:
        return os.path.abspath(paths[0]), os.path.abspath(paths[1]), len(paths)
    if len(paths) == 1:
        return os.path.abspath(paths[0]), None, 1

    # Nenhum encontrado
    return None, None, 0

# ==========================================================
# Persistência robusta de ficheiros (evita perder uploads)
# ==========================================================
def persist_uploaded_mapping(files_map: dict, state_key: str = "sap_files_persisted") -> dict:
    """
    Converte o dicionário de UploadedFile -> {key: UploadedFile}
    num dicionário persistente -> {key: {"name": str, "bytes": bytes}}
    e guarda no st.session_state[state_key].
    """
    if not files_map:
        st.session_state[state_key] = {}
        return {}

    persisted = {}
    for k, f in files_map.items():
        if f is None:
            continue
        # obter bytes de forma segura
        try:
            content = f.getvalue()
        except Exception:
            try:
                f.seek(0)
                content = f.read()
            except Exception:
                continue
        persisted[k] = {
            "name": getattr(f, "name", f"{k}.xlsx"),
            "bytes": content,
        }
    st.session_state[state_key] = persisted
    return persisted


def rehydrate_persisted_mapping(state_key: str = "sap_files_persisted") -> dict:
    """
    Recria objetos file-like (BytesIO) com atributo .name
    a partir do dicionário persistido em st.session_state[state_key].
    Retorna {key: bytesIO_com_name}.
    """
    stored = st.session_state.get(state_key, {}) or {}
    out = {}
    for k, meta in stored.items():
        if not meta:
            continue
        bio = io.BytesIO(meta.get("bytes", b""))
        # anexar nome para ser compatível com o resto do código (logs etc.)
        try:
            bio.name = meta.get("name", f"{k}.xlsx")
        except Exception:
            pass
        out[k] = bio
    return out

# ==========================================================
# NOVO: Classificador Universal de Ficheiros (Multi‑Upload)
# ==========================================================
def classify_uploaded_files(files):
    """
    Recebe uma lista de ficheiros carregados e classifica-os automaticamente
    pelas páginas SAP, MRP, ATL e Progress.
    """
    mapping = {
        # SAP Validation
        "sap": None,
        "mara": None,
        "marc": None,
        "mvke": None,
        "zplm_marc": None,
        "zplm_mara": None,
        "qmat": None,

        # ATL Workflow
        "auto": None,
        "plm": None,
        "timing": None,

        # Progress
        "current": None,
        "previous": None,
    }

    if not files:
        return mapping

    for f in files:
        name = f.name.lower()

        # SAP
        if "sapui5" in name or "item" in name:
            mapping["sap"] = f
            continue
        if "mara" in name and "zplm" not in name:
            mapping["mara"] = f
            continue
        if "marc" in name and "zplm" not in name:
            mapping["marc"] = f
            continue
        if "mvke" in name:
            mapping["mvke"] = f
            continue
        if "zplm" in name and "marc" in name:
            mapping["zplm_marc"] = f
            continue
        if "zplm" in name and "mara" in name:
            mapping["zplm_mara"] = f
            continue
        if "qmat" in name:
            mapping["qmat"] = f
            continue


        # ATL
        if "automotive" in name:
            mapping["auto"] = f
            continue
        if "timingletter" in name or "timing_letter" in name:
            mapping["plm"] = f
            continue
        if "timing" in name:
            mapping["timing"] = f
            continue

        # PROGRESS
        if "current" in name or "novo" in name or "atual" in name:
            mapping["current"] = f
            continue
        if "previous" in name or "anterior" in name or "old" in name:
            mapping["previous"] = f
            continue

    return mapping

# ==========================================================
# NOVO: Quais ficheiros são relevantes por página
# ==========================================================
def required_keys_for_page(page_choice: str) -> list[str]:
    if page_choice == "📦 SAP Material Validation":
        return [
            "Item Browse Report.xlsx",
            "MARA.xlsx",
            "MARC.xlsx",
            "MVKE.xlsx",
            "ZPLM MARC.xlsx",
            "ZPLM MARA.xlsx",
            "QMAT.xlsx",
        ]

    if page_choice == "🧮 MRP Validation":
        return ["MARA.xlsx", "MARC.xlsx", "ZPLM MARA.xlsx"]

    if page_choice == "Automotive + Timing Letter (A & TL)":
        return [
            "PLM Automotive Validation.xlsx",
            "PLM TimingLetter Validation.xlsx",
            "Timing Letters.xlsx",
        ]
    if page_choice == "Change Cockpit":
            # Ajusta quando definires os ficheiros desta página
            return []

    if page_choice == "📈 SAP Material Validation Progress":
        return [
            "(date)_SAP_Validation_Adjusted_Report_ALL.xlsx",
            "(Date)_SAP_Validation_Adjusted_Report_ALL.xlsx",
        ]

    return []

# ==========================================================
# função de estilização de erros (DataFrame)
# ==========================================================
def colorize_errors(df):
    def highlight_row(row):
        error_text = str(row.get("Error", "")).lower()

        if "mandatory" in error_text or "obrigatório" in error_text:
            return ['background-color: #ffcccc'] * len(row)  # red light
        if "invalid" in error_text or "inválido" in error_text:
            return ['background-color: #ffe5b4'] * len(row)  # orange
        if "view" in error_text and "missing" in error_text:
            return ['background-color: #fff7a8'] * len(row)  # yellow
        return [''] * len(row)

    return df.style.apply(highlight_row, axis=1)

# ==========================================================
# Nomes bonitos para exibir no Expander
# ==========================================================
LABELS_PT = {
    "sap": "SAPUI5 Export (Item Browse Report.xlsx)",
    "mara": "MARA.xlsx",
    "marc": "MARC.xlsx",
    "mvke": "MVKE.xlsx",
    "zplm_marc": "ZPLM MARC.xlsx",
    "zplm_mara": "ZPLM MARA.xlsx",
    "qmat": "QMAT.xlsx",

    "auto": "PLM Automotive Validation.xlsx",
    "plm": "PLM TimingLetter Validation.xlsx",
    "timing": "Timing Letters.xlsx",

    "current": "Relatório Atual (Progress)",
    "previous": "Relatório Anterior (Progress)",
}

# --- Helper: converter display name -> internal key ---
DISPLAY_TO_INTERNAL = {
    v.lower(): k for k, v in LABELS_PT.items()
}

def display_to_internal_key(display_name: str):
    """Recebe o nome bonito (ex 'MARA.xlsx') e devolve 'mara'."""
    if not display_name:
        return None
    return DISPLAY_TO_INTERNAL.get(display_name.strip().lower())

# ==========================================================
# SIDEBAR MENU + SUBMENU
# ==========================================================
st.sidebar.header("Report Selection")

main_choice = st.sidebar.radio(
    "",
    [
        "📦 SAP Material Validation",
        "🚗 PLM Workflow",
        "🧮 MRP Validation",
        "📈 SAP Material Validation Progress"
    ]
)

sub_choice = None
if main_choice == "🚗 PLM Workflow":

    # Título visual
    st.sidebar.markdown(
        "<div class='sidebar-submenu-title'>▼ Options</div>",
        unsafe_allow_html=True
    )

    # Container indentado (submenu)
    with st.sidebar.container():
        st.sidebar.markdown("<div class='sidebar-submenu-wrap'>", unsafe_allow_html=True)

        sub_choice = st.sidebar.radio(
            "",
            [
                "Automotive + Timing Letter (A & TL)",
                "Change Cockpit"
            ],
            key="submenu_plm"
        )

        st.sidebar.markdown("</div>", unsafe_allow_html=True)

# ==========================================================
# AUXILIAR DE LEITURA DE FOLHAS
# ==========================================================
def has_sheet(uploaded, sheet_name: str):
    try:
        xl = pd.ExcelFile(uploaded)
        return sheet_name in xl.sheet_names, xl.sheet_names
    except Exception as e:
        return False, [f"Erro ao abrir: {e}"]

# ==========================================================
# IDENTIDADE DO UTILIZADOR (CROSS-PAGE)
# ==========================================================
if "executor_name" not in st.session_state:
    st.session_state["executor_name"] = ""

def identity_input(label: str, key: str):
    """
    Entrada consistente para todas as páginas.
    Se já existir nome na session_state, usa-o.
    Caso contrário, campo aparece vazio.
    """
    executor_name = st.text_input(
        label,
        value=st.session_state["executor_name"],
        placeholder="CDS ID...",
        key=key
    )
    st.session_state["executor_name"] = executor_name.strip()

    if not st.session_state["executor_name"]:
        st.warning("⚠ Please enter your CDS ID to continue")
        st.stop()

    return st.session_state["executor_name"]

# ======================================================================================
# PAGE 1
# SAP MATERIAL VALIDATION (WIZARD + PERSISTÊNCIA DE UPLOADS)
# ======================================================================================
if main_choice == "📦 SAP Material Validation":

    import os
    import pandas as pd
    import streamlit as st

    st.subheader("📦 SAP Material Validation Tool")

    # CSS global (ficheiro externo)
    inject_css("assets/styles.css")

    # ------------------------------
    # IDENTIDADE E PREFERÊNCIAS
    # ------------------------------
    effective_executor = identity_input("Who performs?", "sap_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    # ------------------------------
    # ESTADO DO WIZARD
    # ------------------------------
    if "sap_step" not in st.session_state:
        st.session_state.sap_step = 1

    # Wizard visual (premium)
    
    def render_wizard(current_step: int):
        steps = ["Files", "Scope", "Validate", "Results"]
        cols = st.columns(len(steps))

        for idx, col in enumerate(cols, start=1):
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
                        <div class="step-label">{steps[idx-1]}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

        completion = (current_step - 1) / (len(steps)-1)
        st.markdown(
            f'<div class="progress-line" style="width:{completion*100}%"></div>',
            unsafe_allow_html=True
        )

    render_wizard(st.session_state.sap_step)

    # ------------------------------
    # GUARDAS / HELPERS DE ESTADO
    # ------------------------------
    # Se alguém saltar para passo > 1 sem ter ficheiros persistidos, voltamos ao 1
    required_keys = ["sap", "mara", "marc", "mvke", "zplm_marc", "zplm_mara", "qmat"]
    persisted_now = st.session_state.get("sap_files_persisted", {}) or {}
    if st.session_state.sap_step > 1 and any(k not in persisted_now for k in required_keys):
        st.session_state.sap_step = 1
        st.rerun()

    # AUXILIAR: valida se a folha SAPUI5 Export existe (se já tiveres esta função no app, podes reaproveitar)
    def has_sheet(uploaded, sheet_name: str):
        try:
            xl = pd.ExcelFile(uploaded)
            return sheet_name in xl.sheet_names, xl.sheet_names
        except Exception as e:
            return False, [f"Erro ao abrir: {e}"]

    # ==========================================================================
    # PASSO 1 — FICHEIROS
    # ==========================================================================
    if st.session_state.sap_step == 1:

        st.markdown("### 📂 Step 1 — Load Files")

        uploads = st.file_uploader(
            "Drop all necessary files here",
            type="xlsx",
            accept_multiple_files=True,
            key="sap_multi_uploader"  # chave estável evita resets visuais
        )

        # Classificar uploads mais recentes
        files = classify_uploaded_files(uploads)

        # Se houve uploads agora, persistimos (bytes + name) para não perder nos reruns
        if uploads:
            persist_uploaded_mapping(files, state_key="sap_files_persisted")

        # Mostrar deteção do upload atual
        
        with st.expander("🔍 Files identified (Current upload)"):
            for display_name in required_keys_for_page(main_choice):
                internal = display_to_internal_key(display_name)
                f = files.get(internal) if internal else None
                st.write(f"**{display_name}** → {getattr(f, 'name', '—')}")

        # Mostrar o que está persistido na sessão (o que realmente conta para avançar)
        with st.expander("💾 Files saved in the session (persisted)"):
            persisted = st.session_state.get("sap_files_persisted", {}) or {}
            if persisted:
                for k, meta in persisted.items():
                    label = LABELS_PT.get(k, k)
                    st.write(f"**{label}** → {meta.get('name', f'{k}.xlsx')}")
            else:
                st.write("There are no persisted files yet.")

        # Botões (baseado no que está PERSISTIDO, não no upload atual)
        persisted = st.session_state.get("sap_files_persisted", {}) or {}
        disabled = any(k not in persisted for k in required_keys)

        c1, c2 = st.columns([1,1])
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, width='stretch'):
                st.session_state.sap_step = 2
                st.rerun()

        # Utilitário: Limpar ficheiros persistidos (opcional)
        with st.expander("🧹 Clear uploaded files"):
            if st.button("Clear persistent uploads"):
                st.session_state["sap_files_persisted"] = {}
                st.success("Persistent files have been cleared")
                st.rerun()

    # ==========================================================================
    # PASSO 2 — ÂMBITO
    # ==========================================================================
    if st.session_state.sap_step == 2:

        st.markdown("### 🌍 Step 2 — Select Scope")

        export_scope = st.radio(
            "Scope of download:",
            ["All", "Only PT01", "Only PT02"],
            index=0, horizontal=True
        )
        plant_filter_export = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(export_scope)

        apply_filter_display = st.checkbox("Apply filter to displayed tables", value=False)
        plant_filter_display = plant_filter_export if apply_filter_display else None

        # Guardar no estado
        st.session_state.sap_export_scope = export_scope
        st.session_state.sap_plant_export = plant_filter_export
        st.session_state.sap_plant_display = plant_filter_display

        # Botões
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("⬅️ Back (Files)", width='stretch'):
                st.session_state.sap_step = 1
                st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", width='stretch'):
                st.session_state.sap_step = 3
                st.rerun()

    # ==========================================================================
    # PASSO 3 — VALIDAR
    # ==========================================================================
    if st.session_state.sap_step == 3:

        st.markdown("### ⚙️ Step 3 — Run Validation")
        st.info("Click the button below to validate the uploaded files.")

        if st.button("▶ Validate now", type="primary"):

            # Rehidratar ficheiros persistidos como BytesIO (com .name)
            files = rehydrate_persisted_mapping(state_key="sap_files_persisted")
            sap_file   = files.get("sap")
            mara_file  = files.get("mara")
            marc_file  = files.get("marc")
            mvke_file  = files.get("mvke")
            zplm_marc  = files.get("zplm_marc")
            zplm_mara  = files.get("zplm_mara")
            qmat_file  = files.get("qmat")

            plant_filter_export = st.session_state.sap_plant_export
            plant_filter_display = st.session_state.sap_plant_display

            # Sanity check da folha SAPUI5 Export
            ok_sheet, sheets = has_sheet(sap_file, "SAPUI5 Export")
            if not ok_sheet:
                st.error(f"The sheet 'SAPUI5 Export' does not exist. Sheets found: {sheets}")
                st.stop()

            progress_bar = st.progress(0, text="Initializing validation...")

            def update_progress(percent: int, message: str = ""):
                progress_bar.progress(percent, text=message)

            from sap_validation_engine import run_full_validation

            with st.spinner("Validating materials..."):

                # Execução para visualização (display)
                
                excel_display, errors_df, summary_df = run_full_validation(
                    sap_file=sap_file,
                    mara_file=mara_file,
                    marc_file=marc_file,
                    mvke_file=mvke_file,
                    qmat_file=qmat_file,
                    zplm_marc_file=zplm_marc,
                    zplm_mara_file=zplm_mara,
                    previous_report=None,
                    plant_filter=plant_filter_display,
                    progress_callback=update_progress
                )

                # Execução para exportação (se necessário)
                if plant_filter_export != plant_filter_display:
                    excel_export, _, _ = run_full_validation(
                        sap_file=sap_file,
                        mara_file=mara_file,
                        marc_file=marc_file,
                        mvke_file=mvke_file,
                        qmat_file=qmat_file,
                        zplm_marc_file=zplm_marc,
                        zplm_mara_file=zplm_mara,
                        previous_report=None,
                        plant_filter=plant_filter_export
                    )
                else:
                    excel_export = excel_display

            # Guardar resultados para o Passo 4
            st.session_state.sap_errors_df    = errors_df
            st.session_state.sap_summary_df   = summary_df
            st.session_state.sap_excel_export = excel_export

            # Guardar no vault + Logs
            suffix = plant_filter_export if plant_filter_export else "ALL"
            from sap_validation_engine import save_report_to_vault, build_sap_log_payload, save_run_log
            report_path = save_report_to_vault(excel_export, report_type="sap", plant_suffix=suffix)
            st.session_state.sap_report_path = report_path

            if save_history:
                payload = build_sap_log_payload(
                    user=effective_executor,
                    plant_filter=plant_filter_export,
                    sap_file=sap_file, mara_file=mara_file, marc_file=marc_file, mvke_file=mvke_file,
                    errors_df=errors_df, summary_df=summary_df,
                    report_path=report_path
                )
                save_run_log(payload)

            st.session_state.sap_step = 4
            st.rerun()

    # ==========================================================================
    # PASSO 4 — RESULTADOS
    # ==========================================================================
    if st.session_state.sap_step == 4:

        st.markdown("### 📊 Step 4 — Results")

        errors_df    = st.session_state.sap_errors_df
        summary_df   = st.session_state.sap_summary_df
        excel_export = st.session_state.sap_excel_export
        report_path  = st.session_state.sap_report_path

        # KPI Cards (usam classes do CSS externo)
        c = st.columns(3)
        with c[0]:
            st.markdown(
                f'<div class="card"><h4>Total Errors</h4><div class="big">{errors_df.shape[0]}</div></div>',
                unsafe_allow_html=True
            )
        with c[1]:
            total_fields = int(errors_df["Field"].nunique()) if "Field" in errors_df.columns else 0
            st.markdown(
                f'<div class="card"><h4>Distinctive Fields</h4><div class="big">{total_fields}</div></div>',
                unsafe_allow_html=True
            )
        with c[2]:
            total_mats = int(errors_df["Material"].nunique()) if "Material" in errors_df.columns else 0
            st.markdown(
                f'<div class="card"><h4>Materials with Errors</h4><div class="big">{total_mats}</div></div>',
                unsafe_allow_html=True
            )

        # Tabelas
        st.header("All Errors")
        st.dataframe(errors_df, width='stretch')

        st.header("Summary")
        st.dataframe(summary_df, width='stretch')

        # Dashboards
        st.markdown("## 📈 Dashboard — SAP Material Validation")
        render_sap_run_kpis(errors_df)
        render_sap_run_charts(errors_df)

        # Download com nome igual ao do ficheiro arquivado
        download_name = os.path.basename(report_path) if report_path else "SAP_Validation_Adjusted_Report_ALL.xlsx"
        st.download_button(
            "⬇️ Download Excel",
            data=excel_export.getvalue(),
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Histórico SAP
        with st.expander("🕘 Histórico SAP (últimas 10 execuções)"):
            logs = list_run_logs("sap", limit=10)
            if logs:
                hist_df = pd.DataFrame([
                    {
                        "When": log.get("timestamp"),
                        "Who": log.get("user"),
                        "Filter": log.get("plant_filter"),
                        "Errors": log.get("metrics", {}).get("total_errors", 0),
                    }
                    for log in logs
                ])
                st.dataframe(hist_df, width='stretch')
            else:
                st.write("No history.")

        # Nova execução
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("🔁 New Validation", width='stretch'):
                # Opcional: manter uploads persistidos ou limpar
                st.session_state.sap_step = 1
                st.rerun()
        with c2:
            if st.button("🧹 Clear persisted uploads and start over", width='stretch'):
                st.session_state["sap_files_persisted"] = {}
                st.session_state.sap_step = 1
                st.rerun()

    # app.py – no final do script (depois de renderizar a página seleccionada)
    render_footer(start_time, version="v1.1")
    
# ==========================================================
# PAGE 2 — SUBMENU - PLM ATL — PREMIUM WIZARD
# ==========================================================
if main_choice == "🚗 PLM Workflow" and sub_choice == "Automotive + Timing Letter (A & TL)":

    st.subheader("🚗 PLM Workflow (ATL)")

    inject_css("assets/styles.css")

    # ------------------------------
    # IDENTIDADE (igual Página 1)
    # ------------------------------
    effective_executor = identity_input("Who performs?", "atl_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    # ------------------------------
    # ESTADO DO WIZARD
    # ------------------------------
    if "atl_step" not in st.session_state:
        st.session_state.atl_step = 1

    # ------------------------------
    # WIZARD VISUAL
    # ------------------------------
    def render_atl_wizard(current_step: int):
        steps = ["Files", "Scope", "Validate", "Results"]
        cols = st.columns(len(steps))
        for idx, col in enumerate(cols, start=1):
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
    <div class="step-label">{steps[idx-1]}</div>
</div>
                    """,
                    unsafe_allow_html=True
                )

        completion = (current_step - 1) / (len(steps) - 1)
        st.markdown(
            f"""
<div class="progress-line" style="width:{completion * 100}%"></div>
            """,
            unsafe_allow_html=True
        )

    render_atl_wizard(st.session_state.atl_step)

    # ------------------------------
    # CLASSIFY UPLOADS
    # ------------------------------
    def classify_atl_files(files):
        mapping = {"auto": None, "plm": None, "timing": None}
        if not files:
            return mapping

        for f in files:
            name = f.name.lower()

            if "automotive" in name:
                mapping["auto"] = f
                continue

            if "timingletter" in name or "timing_letter" in name:
                mapping["plm"] = f
                continue

            if "timing" in name:
                mapping["timing"] = f
                continue

        return mapping

    # ==========================================================
    # PASSO 1 — UPLOAD
    # ==========================================================
    if st.session_state.atl_step == 1:

        st.markdown("### 📂 Step 1 — Load Files")

        uploads = st.file_uploader(
            "Drop Automotive + TimingLetter + Timing Letter",
            type="xlsx",
            accept_multiple_files=True,
            key="atl_multi_uploader"
        )

        files = classify_atl_files(uploads)

        if uploads:
            persist_uploaded_mapping(files, state_key="atl_files_persisted")
   
        with st.expander("🔍 Files identified (Current upload)"):
            for display_name in required_keys_for_page(sub_choice):
                internal = display_to_internal_key(display_name)
                f = files.get(internal) if internal else None
                st.write(f"**{display_name}** → {getattr(f, 'name', '—')}")

        with st.expander("💾 Files saved in session (persisted)"):
            persisted = st.session_state.get("atl_files_persisted", {}) or {}
            if persisted:
                for k, meta in persisted.items():
                    st.write(f"**{k}** → {meta.get('name','')}")
            else:
                st.write("There are no persisted files yet.")

        persisted = st.session_state.get("atl_files_persisted", {}) or {}
        disabled = any(k not in persisted for k in ["auto", "plm", "timing"])

        c1, c2 = st.columns([1, 1])
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, width='stretch'):
                st.session_state.atl_step = 2
                st.rerun()

        with st.expander("🧹 Clear persistent uploads"):
            if st.button("Clear ATL uploads"):
                st.session_state["atl_files_persisted"] = {}
                st.success("Cleared.")
                st.rerun()

    # ==========================================================
    # PASSO 2 — ÂMBITO
    # ==========================================================
    if st.session_state.atl_step == 2:

        st.markdown("### 🌍 Step 2 — Choose Scope")

        atl_scope = st.radio("Scope:", ["All", "Only PT01", "Only PT02"], horizontal=True)
        atl_filter_export = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(atl_scope)

        apply_filter_display = st.checkbox("Apply filter to the displayed tables", value=False)
        atl_filter_display = atl_filter_export if apply_filter_display else None

        st.session_state.atl_scope = atl_scope
        st.session_state.atl_filter_export = atl_filter_export
        st.session_state.atl_filter_display = atl_filter_display

        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("⬅️ Back (Files)", width='stretch'):
                st.session_state.atl_step = 1
                st.rerun()

        with c2:
            if st.button("➡️ Continue to Validation", width='stretch'):
                st.session_state.atl_step = 3
                st.rerun()

    # ==========================================================
    # PASSO 3 — RUN ATL
    # ==========================================================
    if st.session_state.atl_step == 3:

        st.markdown("### ⚙️ Step 3 — Run ATL")

        if st.button("▶️ Run Workflow ATL", type="primary"):

            files = rehydrate_persisted_mapping("atl_files_persisted")

            auto_file = files.get("auto")
            plm_file = files.get("plm")
            timing_file = files.get("timing")

            if any(f is None for f in [auto_file, plm_file, timing_file]):
                st.error("⚠️ Required files are missing.")
                st.stop()

            with st.spinner("Processing Workflow ATL..."):
                excel_all, df_pt01, df_pt02 = run_plm_workflow_combined(
                    auto_file=auto_file,
                    plm_file=plm_file,
                    timing_file=timing_file
                )

            st.session_state.atl_pt01 = df_pt01
            st.session_state.atl_pt02 = df_pt02
            st.session_state.atl_excel_all = excel_all

            if save_history:
                payload = build_atl_log_payload(
                    user=effective_executor,
                    plant_scope=st.session_state.atl_scope,
                    df_pt01=df_pt01,
                    df_pt02=df_pt02
                )
                save_run_log(payload)

            st.session_state.atl_step = 4
            st.rerun()

    # ==========================================================
    # PASSO 4 — RESULTADOS
    # ==========================================================
    if st.session_state.atl_step == 4:

        st.markdown("### 📊 Step 4 — Results")

        df_pt01 = st.session_state.atl_pt01
        df_pt02 = st.session_state.atl_pt02
        excel_all = st.session_state.atl_excel_all

        # ------------------------------
        # VISUALIZAÇÃO (scope filtro)
        # ------------------------------
        scope = st.session_state.atl_filter_display

        if scope == "PT01":
            st.header("PT01")
            st.dataframe(df_pt01, width='stretch')

        elif scope == "PT02":
            st.header("PT02")
            st.dataframe(df_pt02, width='stretch')

        else:
            st.header("PT01")
            st.dataframe(df_pt01, width='stretch')

            st.header("PT02")
            st.dataframe(df_pt02, width='stretch')

        # ------------------------------
        # DASHBOARD
        # ------------------------------
        st.markdown("## 📈 Dashboard — PLM Workflow (ATL)")
        render_atl_run_kpis(df_pt01, df_pt02)
        render_atl_run_charts(df_pt01, df_pt02)

        # ------------------------------
        # DOWNLOAD
        # ------------------------------
        suffix = st.session_state.atl_filter_export if st.session_state.atl_filter_export else "ALL"

        st.download_button(
            f"⬇️ Download PLM_Workflow_{suffix}.xlsx",
            data=excel_all.getvalue(),
            file_name=f"PLM_Workflow_{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ------------------------------
        # HISTÓRICO
        # ------------------------------
        with st.expander("🕑 History ATL (Last 10 runs)"):
            logs = list_run_logs("atl", limit=10)
            if logs:
                hist_df = pd.DataFrame([
                    {
                        "When": log["timestamp"],
                        "Who": log["user"],
                        "Scope": log["plant_scope"],
                        "PT01 rows": log["metrics"]["PT01_rows"],
                        "PT02 rows": log["metrics"]["PT02_rows"]
                    }
                    for log in logs
                ])
                st.dataframe(hist_df, width='stretch')
            else:
                st.write("No History.")

        # ------------------------------
        # NOVA EXECUÇÃO
        # ------------------------------
        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("🔄 New Run ATL", width='stretch'):
                st.session_state.atl_step = 1
                st.rerun()

        with c2:
            if st.button("🧹 Clear persistent uploads and return to the beginning.", width='stretch'):
                st.session_state["atl_files_persisted"] = {}
                st.session_state.atl_step = 1
                st.rerun()

    # app.py – no final do script (depois de renderizar a página seleccionada)
    render_footer(start_time, version="v1.1")

# ==========================================================
# PAGE 2 — SUBMENU - PLM CHANGE COCKPIT — PREMIUM WIZARD
# ==========================================================
if main_choice == "🚗 PLM Workflow" and (sub_choice == "Change Cockpit"):
    # --- NOVA PÁGINA: Change Cockpit ---
    st.subheader("🛠️ Change Cockpit (In Progess...)")
    inject_css("assets/styles.css")

    # Identidade (consistente com as restantes páginas)
    effective_executor = identity_input("Who performs?", "cc_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    # Estado inicial do wizard (opcional, se quiseres mesmo ter 4 passos iguais)
    if "cc_step" not in st.session_state:
        st.session_state.cc_step = 1

    # Wizard visual (exemplo; copiares o padrão que usas nas outras páginas)
    def render_cc_wizard(current_step: int):
        steps = ["Files", "Scope", "Validate", "Results"]
        cols = st.columns(len(steps))
        for idx, col in enumerate(cols, start=1):
            with col:
                klass = "active" if current_step == idx else ("done" if current_step > idx else "pending")
                st.markdown(
                    f"""
<div class="wizard-step {klass}">
  <div class="step-index">{idx}</div>
  <div class="step-label">{steps[idx-1]}</div>
</div>
""",
                    unsafe_allow_html=True,
                )
        completion = (current_step - 1) / (len(steps) - 1)
        st.markdown(f'<div class="wizard-progress"><div style="width:{completion*100:.0f}%"></div></div>', unsafe_allow_html=True)

    render_cc_wizard(st.session_state.cc_step)

    # PASSO 1 — Ficheiros (placeholder)
    if st.session_state.cc_step == 1:
        st.markdown("### 📂 Step 1 — Load Files")
        uploads = st.file_uploader(
            "Drop the Change Cockpit input(s) here",
            type=["xlsx", "csv"],
            accept_multiple_files=True,
            key="cc_multi_uploader"
        )
        # TODO: classificar ficheiros quando definires o modelo CC
        can_continue = True  # troca quando soubermos os files obrigatórios
        c1, c2 = st.columns([1, 1])
        with c2:
            if st.button("➡️ Continue to Scope", disabled=not can_continue, width='stretch'):
                st.session_state.cc_step = 2
                st.rerun()

    # PASSO 2 — Âmbito (placeholder)
    if st.session_state.cc_step == 2:
        st.markdown("### 🌍 Step 2 — Select Scope")
        cc_scope = st.radio("Scope:", ["All", "Only PT01", "Only PT02"], horizontal=True)
        st.session_state.cc_scope = cc_scope
        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("⬅️ Back (Files)", width='stretch'):
                st.session_state.cc_step = 1
                st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", width='stretch'):
                st.session_state.cc_step = 3
                st.rerun()

    # PASSO 3 — Validar/Executar (placeholder)
    if st.session_state.cc_step == 3:
        st.markdown("### ⚙️ Step 3 — Run Change Cockpit")
        st.info("Define here the processing for Change Cockpit.")
        if st.button("▶️ Run Change Cockpit", type="primary"):
            with st.spinner("Running Change Cockpit..."):
                # TODO: executar a lógica CC e preparar dataframes/Excel
                st.session_state.cc_result_df = None
                st.session_state.cc_excel = None
            st.session_state.cc_step = 4
            st.rerun()

    # PASSO 4 — Resultados (placeholder)
    if st.session_state.cc_step == 4:
        st.markdown("### 📈 Step 4 — Results")
        st.success("Results placeholder. Plug your tables and dashboards here.")

        # Exemplo de botões finais
        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("🔄 New Run", width='stretch'):
                st.session_state.cc_step = 1
                st.rerun()
        with c2:
            if st.button("🧹 Clear & Start Over", width='stretch'):
                for k in ["cc_result_df", "cc_excel"]:
                    st.session_state.pop(k, None)
                st.session_state.cc_step = 1
                st.rerun()
    # app.py – no final do script (depois de renderizar a página seleccionada)
    render_footer(start_time, version="v1.1")

# ==========================================================
# PAGE 3 — MRP VALIDATION — PREMIUM WIZARD
# ==========================================================
elif main_choice == "🧮 MRP Validation":

    st.subheader("🧮 MRP Validation")

    inject_css("assets/styles.css")

    # ------------------------------
    # IDENTIDADE (consistente com outras páginas)
    # ------------------------------
    effective_executor = identity_input("Who performs?", "mrp_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    # ------------------------------
    # ESTADO DO WIZARD
    # ------------------------------
    if "mrp_step" not in st.session_state:
        st.session_state.mrp_step = 1

    # ------------------------------
    # WIZARD VISUAL
    # ------------------------------
    def render_mrp_wizard(current_step: int):
        steps = ["Files", "Scope", "Validate", "Results"]
        cols = st.columns(len(steps))
        for idx, col in enumerate(cols, start=1):
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
    <div class="step-label">{steps[idx-1]}</div>
</div>
                    """,
                    unsafe_allow_html=True
                )

        completion = (current_step - 1) / (len(steps) - 1)
        st.markdown(
            f"""
<div class="progress-line" style="width:{completion * 100}%"></div>
            """,
            unsafe_allow_html=True
        )

    render_mrp_wizard(st.session_state.mrp_step)

    # ------------------------------
    # CLASSIFY UPLOADS
    # ------------------------------
    def classify_mrp(files):
        mapping = {"mara": None, "marc": None, "zplm_marc": None}
        if not files:
            return mapping

        for f in files:
            name = f.name.lower()

            if "mara" in name and "zplm" not in name:
                mapping["mara"] = f
                continue

            if "marc" in name and "zplm" not in name:
                mapping["marc"] = f
                continue

            if "zplm" in name and "marc" in name:
                mapping["zplm_marc"] = f
                continue

        return mapping

    # ==========================================================
    # PASSO 1 — UPLOAD
    # ==========================================================
    if st.session_state.mrp_step == 1:

        st.markdown("### 📂 Step 1 — Load Files")

        uploads = st.file_uploader(
            "Drop MARA + MARC + (optional) ZPLM MARC",
            type="xlsx",
            accept_multiple_files=True,
            key="mrp_multi_uploader"
        )

        files = classify_mrp(uploads)

        if uploads:
            persist_uploaded_mapping(files, state_key="mrp_files_persisted")

        with st.expander("🔍 Files identified (Current Upload)"):
            for display_name in required_keys_for_page(main_choice):
                internal = display_to_internal_key(display_name)
                f = files.get(internal) if internal else None
                st.write(f"**{display_name}** → {getattr(f, 'name', '—')}")

        with st.expander("💾 Files saved in session (persisted)"):
            persisted = st.session_state.get("mrp_files_persisted", {}) or {}
            if persisted:
                for k, meta in persisted.items():
                    st.write(f"**{k}** → {meta.get('name')}")
            else:
                st.write("There are no persisted files yet.")

        persisted = st.session_state.get("mrp_files_persisted", {}) or {}
        disabled = any(k not in persisted for k in ["mara", "marc"])

        c1, c2 = st.columns([1, 1])
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, width='stretch'):
                st.session_state.mrp_step = 2
                st.rerun()

        with st.expander("🧹 Clear persistent uploads"):
            if st.button("Clear MRP Uploads"):
                st.session_state["mrp_files_persisted"] = {}
                st.success("Cleared.")
                st.rerun()

    # ==========================================================
    # PASSO 2 — ÂMBITO
    # ==========================================================
    if st.session_state.mrp_step == 2:

        st.markdown("### 🌍 Step 2 — Choose Scope")

        export_scope = st.radio(
            "Scope:",
            ["All", "Only PT01", "Only PT02"],
            horizontal=True
        )

        plant_filter_export = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(export_scope)

        apply_filter_display = st.checkbox("Apply filter to the displayedd tables", value=False)
        plant_filter_display = plant_filter_export if apply_filter_display else None

        st.session_state.mrp_export_scope = export_scope
        st.session_state.mrp_plant_export = plant_filter_export
        st.session_state.mrp_plant_display = plant_filter_display

        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("⬅️ Back (Files)", width='stretch'):
                st.session_state.mrp_step = 1
                st.rerun()

        with c2:
            if st.button("➡️ Continue to Validation", width='stretch'):
                st.session_state.mrp_step = 3
                st.rerun()

    # ==========================================================
    # PASSO 3 — VALIDAR
    # ==========================================================
    if st.session_state.mrp_step == 3:

        st.markdown("### ⚙️ Step 3 — Run MRP Validation")

        st.info("Run all MRP Validation Files")

        if st.button("▶️ Execute MRP validation", type="primary"):

            files = rehydrate_persisted_mapping("mrp_files_persisted")

            mara_file = files.get("mara")
            marc_file = files.get("marc")
            zplm_marc = files.get("zplm_marc")

            plant_filter_export = st.session_state.mrp_plant_export
            plant_filter_display = st.session_state.mrp_plant_display

            with st.spinner("Validating MRP files..."):

                excel_display, errors_df, summary_df = run_mrp_validation(
                    mara_file=mara_file,
                    marc_file=marc_file,
                    zplm_marc_file=zplm_marc,
                    plant_filter=plant_filter_display
                )

                if plant_filter_export != plant_filter_display:
                    excel_export, _, _ = run_mrp_validation(
                        mara_file=mara_file,
                        marc_file=marc_file,
                        zplm_marc_file=zplm_marc,
                        plant_filter=plant_filter_export
                    )
                else:
                    excel_export = excel_display

            st.session_state.mrp_errors_df = errors_df
            st.session_state.mrp_summary_df = summary_df
            st.session_state.mrp_excel_export = excel_export

            if save_history:
                payload = build_mrp_log_payload(
                    user=effective_executor,
                    plant_filter=plant_filter_export,
                    mara_file=mara_file,
                    marc_file=marc_file,
                    zplm_marc_file=zplm_marc,
                    errors_df=errors_df,
                    summary_df=summary_df
                )
                save_run_log(payload)

            st.session_state.mrp_step = 4
            st.rerun()

    # ==========================================================
    # PASSO 4 — RESULTADOS
    # ==========================================================
    if st.session_state.mrp_step == 4:

        st.markdown("### 📊 Step 4 — Results")

        errors_df = st.session_state.mrp_errors_df
        summary_df = st.session_state.mrp_summary_df
        excel_export = st.session_state.mrp_excel_export

        # KPI CARDS
        c = st.columns(3)
        with c[0]:
            st.markdown(
                f"""
<div class="card">
<h4>Total Errors</h4>
<div class="big">{errors_df.shape[0]}</div>
</div>
                """,
                unsafe_allow_html=True
            )
        with c[1]:
            total_fields = int(errors_df["Field"].nunique()) if "Field" in errors_df else 0
            st.markdown(
                f"""
<div class="card">
<h4>Distinct Fields</h4>
<div class="big">{total_fields}</div>
</div>
                """,
                unsafe_allow_html=True
            )
        with c[2]:
            total_mats = int(errors_df["Material"].nunique()) if "Material" in errors_df else 0
            st.markdown(
                f"""
<div class="card">
<h4>Materiais with errors</h4>
<div class="big">{total_mats}</div>
</div>
                """,
                unsafe_allow_html=True
            )

        # Tabelas
        st.header("All Errors")
        st.dataframe(errors_df, width='stretch')

        st.header("Summary")
        st.dataframe(summary_df, width='stretch')

        # DASHBOARD
        st.markdown("## 📈 Dashboard — MRP Validation")
        render_mrp_run_kpis(errors_df)
        render_mrp_run_charts(errors_df)

        # HISTÓRICO
        with st.expander("🕑 MRP History (Last 10 runs)"):
            logs = list_run_logs("mrp", limit=10)
            if logs:
                hist_df = pd.DataFrame([
                    {
                        "When": log["timestamp"],
                        "Who": log["user"],
                        "Filter": log["plant_filter"],
                        "Errors": log["metrics"]["total_errors"]
                    }
                    for log in logs
                ])
                st.dataframe(hist_df, width='stretch')
            else:
                st.write("No history.")

        # DOWNLOAD
        suffix = st.session_state.mrp_plant_export if st.session_state.mrp_plant_export else "ALL"
        st.download_button(
            f"⬇️ Download MRP_Validation_Adjusted_Report_{suffix}.xlsx",
            data=excel_export.getvalue(),
            file_name=f"MRP_Validation_Adjusted_Report_{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # NOVA EXECUÇÃO
        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("🔄 New Validation", width='stretch'):
                st.session_state.mrp_step = 1
                st.rerun()

        with c2:
            if st.button("🧹 Clear persistent uploads and return to the beginning", width='stretch'):
                st.session_state["mrp_files_persisted"] = {}
                st.session_state.mrp_step = 1
                st.rerun()

    # app.py – no final do script (depois de renderizar a página seleccionada)
    render_footer(start_time, version="v1.1")

# ==========================================================
# PAGE 4 — SAP MATERIAL VALIDATION PROGRESS — PREMIUM WIZARD
# ==========================================================
elif main_choice == "📈 SAP Material Validation Progress":

    st.subheader("📈 SAP Material Validation Progress")

    inject_css("assets/styles.css")

    # ------------------------------
    # IDENTIDADE
    # ------------------------------
    effective_executor = identity_input("Who performs?", "progress_exec")
    save_history = st.checkbox("Save this exectution in history", value=True)

    # ------------------------------
    # ESTADO DO WIZARD
    # ------------------------------
    if "progress_step" not in st.session_state:
        st.session_state.progress_step = 1

    # ------------------------------
    # WIZARD VISUAL (4 passos)
    # ------------------------------
    def render_progress_wizard(current_step: int):
        steps = ["Mode", "Files", "Validate", "Results"]
        cols = st.columns(len(steps))
        for idx, col in enumerate(cols, start=1):
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
    <div class="step-label">{steps[idx-1]}</div>
</div>
                    """,
                    unsafe_allow_html=True,
                )

        completion = (current_step - 1) / (len(steps) - 1)
        st.markdown(
            f'<div class="progress-line" style="width:{completion*100}%"></div>',
            unsafe_allow_html=True,
        )

    render_progress_wizard(st.session_state.progress_step)

    # ==========================================================
    # PASSO 1 — MODO (AUTO vs MANUAL)
    # ==========================================================
    if st.session_state.progress_step == 1:

        st.markdown("### 🧭 Step 1 — Choose Mode")

        mode = st.radio("Comparison Mode:", ["Auto", "Manual"], horizontal=True)
        st.session_state.progress_mode = mode

        if mode == "Auto":
            st.markdown("### 🌍 Scope (Previous Reports)")
            prog_scope = st.radio("Scope:", ["ALL", "PT01", "PT02"], horizontal=True)
            st.session_state.prog_scope = prog_scope

        if st.button("➡️ Continue to Files", width='stretch'):
            st.session_state.progress_step = 2
            st.rerun()

    # ==========================================================
    # PASSO 2 — FICHEIROS
    # ==========================================================
    if st.session_state.progress_step == 2:

        mode = st.session_state.progress_mode
        st.markdown("### 📂 Step 2 — Files")

        # ------------------------------
        # MODO AUTO
        # ------------------------------
        if mode == "Auto":

            prog_scope = st.session_state.prog_scope

            curr_path, prev_path, avail = pick_last_two_reports(prog_scope)

            with st.expander("🔍 Reports found automatically"):
                st.write(f"**Available for {prog_scope}**: {avail}")
                st.write(f"**Most Recent:** {curr_path if curr_path else '—'}")
                st.write(f"**Previous:** {prev_path if prev_path else '—'}")

                if avail == 1:
                    st.info("**1** Report available. Need one more to compare.")

                if avail == 0:
                    st.warning("No reports available to compare.")

            st.session_state.progress_curr_path = curr_path
            st.session_state.progress_prev_path = prev_path

            disabled = not (curr_path and prev_path)

        # ------------------------------
        # MODO MANUAL
        # ------------------------------
        else:
            st.markdown("Load **2 files** (any order). I'll detect the newest automatically.")

            uploads = st.file_uploader(
                "Current + Previous (no renaming!)",
                type="xlsx",
                accept_multiple_files=True,
                key="progress_manual_uploader",
            )

            curr_file, prev_file, notes = pick_current_previous_from_uploads(uploads or [])
            
            with st.expander("🔍 Files identified (Current upload)"):
                for display_name in required_keys_for_page(main_choice):
                    internal = display_to_internal_key(display_name)
                    f = curr_file if internal == "current" else prev_file if internal == "previous" else None
                    st.write(f"**{display_name}** → {getattr(f, 'name', '—')}")

            st.session_state.manual_curr = curr_file
            st.session_state.manual_prev = prev_file

            disabled = not (curr_file and prev_file)

        # ------------------------------
        # BOTÕES
        # ------------------------------
        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("⬅️ Back", width='stretch'):
                st.session_state.progress_step = 1
                st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", disabled=disabled, width='stretch'):
                st.session_state.progress_step = 3
                st.rerun()

    # ==========================================================
    # PASSO 3 — VALIDAR
    # ==========================================================
    if st.session_state.progress_step == 3:

        st.markdown("### ⚙️ Step 3 — Compare Reports")

        if st.button("▶️ Execute Differences", type="primary"):

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

            with st.spinner("Comparing reports"):
                excel_diff, diff_dfs = compare_adjusted_reports(curr_bytes, prev_bytes)

            st.session_state.progress_diff_excel = excel_diff
            st.session_state.progress_diff_dfs = diff_dfs
            st.session_state.progress_current_name = (
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
                    diff_dfs=diff_dfs
                )
                save_run_log(payload)

            st.session_state.progress_step = 4
            st.rerun()

    # ==========================================================
    # PASSO 4 — RESULTADOS
    # ==========================================================
    if st.session_state.progress_step == 4:

        st.markdown("### 📊 Step 4 — Results")

        diff_dfs = st.session_state.progress_diff_dfs

        # ------------------------------
        # TABS
        # ------------------------------
        labels = ["New Errors", "Resolved Errors", "Changed Errors", "New Materials", "Views Lost"]
        keys = ["New Errors", "Resolved Errors", "Changed Errors", "New Materials", "Views Lost"]

        tabs = st.tabs(labels)
        for tab, key in zip(tabs, keys):
            with tab:
                st.dataframe(diff_dfs.get(key, pd.DataFrame()), width='stretch')

        # ------------------------------
        # DASHBOARDS
        # ------------------------------
        st.markdown("## 📈 Dashboard — Progress")
        render_progress_kpis(diff_dfs)
        render_progress_charts(diff_dfs)

        # ------------------------------
        # DOWNLOAD
        # ------------------------------
        st.markdown("---")
        st.download_button(
            "⬇️ Download SAP_Validation_Progress.xlsx",
            data=st.session_state.progress_diff_excel.getvalue(),
            file_name="SAP_Validation_Progress.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # ------------------------------
        # HISTÓRICO
        # ------------------------------
        with st.expander("🕑 Progress History (last 10 executions)"):
            logs = list_run_logs("progress", limit=10)
            if logs:
                hist_df = pd.DataFrame([
                    {
                        "When": log["timestamp"],
                        "Who": log["user"],
                        "New": log["metrics"]["new_errors"],
                        "Resolved": log["metrics"]["resolved_errors"],
                        "Changed": log["metrics"]["changed_errors"],
                        "New Mats": log["metrics"]["new_materials"],
                        "Views Lost": log["metrics"]["views_lost"],
                    }
                    for log in logs
                ])
                st.dataframe(hist_df, width='stretch')
            else:
                st.write("No History.")

        # ------------------------------
        # NOVA EXECUÇÃO
        # ------------------------------
        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("🔄 New comparison", width='stretch'):
                st.session_state.progress_step = 1
                st.rerun()
        with c2:
            if st.button("🧹 Clear all and restart validation", width='stretch'):
                for k in [
                    "progress_curr_path", "progress_prev_path",
                    "progress_manual_curr", "progress_manual_prev",
                    "progress_diff_dfs", "progress_diff_excel"
                ]:
                    st.session_state.pop(k, None)
                st.session_state.progress_step = 1
                st.rerun()
    # app.py – no final do script (depois de renderizar a página seleccionada)
    render_footer(start_time, version="v1.1")