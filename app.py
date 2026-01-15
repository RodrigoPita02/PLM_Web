# app.py
import streamlit as st
import pandas as pd
import io
import os

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

# ==========================================================
# HELPERS PARA MODO MANUAL (Progress) – DETEÇÃO DE "Atual" e "Anterior"
# ==========================================================
import re
from datetime import datetime

_TS_PATTERN = re.compile(r'(\d{8}_\d{6})')  # ex.: 20260114_085435

def extract_timestamp_from_name(filename: str) -> datetime | None:
    """
    Procura por 'YYYYMMDD_HHMMSS' no nome do ficheiro e devolve um datetime.
    Se não encontrar, devolve None.
    """
    if not filename:
        return None
    m = _TS_PATTERN.search(filename)
    if not m:
        return None
    raw = m.group(1)  # '20260114_085435'
    try:
        return datetime.strptime(raw, "%Y%m%d_%H%M%S")
    except Exception:
        return None


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
# Persistência robusta de ficheiros (evita perder uploads)
# ==========================================================
import io

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
    if page_choice == "SAP Material Validation":
        return ["sap", "mara", "marc", "mvke", "zplm_marc", "zplm_mara"]

    if page_choice == "MRP Validation":
        return ["mara", "marc", "zplm_marc"]

    if page_choice == "PLM Workflow (ATL)":
        return ["auto", "plm", "timing"]

    if page_choice == "SAP Material Validation Progress":
        return ["current", "previous"]

    return []


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

    "auto": "PLM Automotive Validation.xlsx",
    "plm": "PLM TimingLetter Validation.xlsx",
    "timing": "Timing Letters.xlsx",

    "current": "Relatório Atual (Progress)",
    "previous": "Relatório Anterior (Progress)",
}


# ==========================================================
# SIDEBAR MENU
# ==========================================================
st.sidebar.header("Menu")
choice = st.sidebar.radio(
    "Escolher relatório:",
    [
        "SAP Material Validation",
        "PLM Workflow (ATL)",
        "MRP Validation",
        "SAP Material Validation Progress"
    ],
    index=0
)

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
        placeholder="Introduza o seu nome…",
        key=key
    )
    st.session_state["executor_name"] = executor_name.strip()

    if not st.session_state["executor_name"]:
        st.warning("⚠ Por favor indique o seu nome para continuar.")
        st.stop()

    return st.session_state["executor_name"]

# ======================================================================================
# PAGE 1
# SAP MATERIAL VALIDATION (WIZARD + PERSISTÊNCIA DE UPLOADS)
# ======================================================================================
if choice == "SAP Material Validation":

    import os
    import pandas as pd
    import streamlit as st

    st.subheader("📦 SAP Material Validation Tool")

    # CSS global (ficheiro externo)
    inject_css("assets/styles.css")

    # ------------------------------
    # IDENTIDADE E PREFERÊNCIAS
    # ------------------------------
    effective_executor = identity_input("Quem executa?", "sap_exec")
    save_history = st.checkbox("Guardar esta execução no histórico", value=True)

    # ------------------------------
    # ESTADO DO WIZARD
    # ------------------------------
    if "sap_step" not in st.session_state:
        st.session_state.sap_step = 1

    # Wizard visual (premium)
    
    def render_wizard(current_step: int):
        steps = ["Ficheiros", "Âmbito", "Validar", "Resultados"]
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
    required_keys = ["sap", "mara", "marc", "mvke"]
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

        st.markdown("### 📂 Passo 1 — Carregar Ficheiros")

        uploads = st.file_uploader(
            "Largar todos os ficheiros necessários aqui",
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
        with st.expander("🔍 Ficheiros identificados (último upload)"):
            for key in required_keys_for_page(choice):
                f = files.get(key)
                label = LABELS_PT.get(key, key)
                st.write(f"**{label}** → {f.name if f else '—'}")

        # Mostrar o que está persistido na sessão (o que realmente conta para avançar)
        with st.expander("💾 Ficheiros guardados na sessão (persistidos)"):
            persisted = st.session_state.get("sap_files_persisted", {}) or {}
            if persisted:
                for k, meta in persisted.items():
                    label = LABELS_PT.get(k, k)
                    st.write(f"**{label}** → {meta.get('name', f'{k}.xlsx')}")
            else:
                st.write("Ainda não existem ficheiros persistidos.")

        # Botões (baseado no que está PERSISTIDO, não no upload atual)
        persisted = st.session_state.get("sap_files_persisted", {}) or {}
        disabled = any(k not in persisted for k in required_keys)

        c1, c2 = st.columns([1,1])
        with c2:
            if st.button("➡️ Continuar para Âmbito", disabled=disabled, use_container_width=True):
                st.session_state.sap_step = 2
                st.rerun()

        # Utilitário: Limpar ficheiros persistidos (opcional)
        with st.expander("🧹 Limpar ficheiros carregados"):
            if st.button("Limpar uploads persistidos"):
                st.session_state["sap_files_persisted"] = {}
                st.success("Ficheiros persistidos foram limpos.")
                st.rerun()

    # ==========================================================================
    # PASSO 2 — ÂMBITO
    # ==========================================================================
    if st.session_state.sap_step == 2:

        st.markdown("### 🌍 Passo 2 — Selecionar Âmbito")

        export_scope = st.radio(
            "Âmbito do download:",
            ["Tudo", "Apenas PT01", "Apenas PT02"],
            index=0, horizontal=True
        )
        plant_filter_export = {"Apenas PT01": "PT01", "Apenas PT02": "PT02"}.get(export_scope)

        apply_filter_display = st.checkbox("Aplicar filtro às tabelas apresentadas", value=False)
        plant_filter_display = plant_filter_export if apply_filter_display else None

        # Guardar no estado
        st.session_state.sap_export_scope = export_scope
        st.session_state.sap_plant_export = plant_filter_export
        st.session_state.sap_plant_display = plant_filter_display

        # Botões
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("⬅️ Voltar (Ficheiros)", use_container_width=True):
                st.session_state.sap_step = 1
                st.rerun()
        with c2:
            if st.button("➡️ Continuar para Validação", use_container_width=True):
                st.session_state.sap_step = 3
                st.rerun()

    # ==========================================================================
    # PASSO 3 — VALIDAR
    # ==========================================================================
    if st.session_state.sap_step == 3:

        st.markdown("### ⚙️ Passo 3 — Executar Validação")
        st.info("Clique no botão abaixo para validar os ficheiros carregados.")

        if st.button("▶ Validar agora", type="primary"):

            # Rehidratar ficheiros persistidos como BytesIO (com .name)
            files = rehydrate_persisted_mapping(state_key="sap_files_persisted")
            sap_file   = files.get("sap")
            mara_file  = files.get("mara")
            marc_file  = files.get("marc")
            mvke_file  = files.get("mvke")
            zplm_marc  = files.get("zplm_marc")
            zplm_mara  = files.get("zplm_mara")

            plant_filter_export = st.session_state.sap_plant_export
            plant_filter_display = st.session_state.sap_plant_display

            # Sanity check da folha SAPUI5 Export
            ok_sheet, sheets = has_sheet(sap_file, "SAPUI5 Export")
            if not ok_sheet:
                st.error(f"A folha 'SAPUI5 Export' não existe. Folhas encontradas: {sheets}")
                st.stop()

            from sap_validation_engine import run_full_validation

            with st.spinner("A validar materiais..."):

                # Execução para visualização (display)
                excel_display, errors_df, summary_df = run_full_validation(
                    sap_file=sap_file,
                    mara_file=mara_file,
                    marc_file=marc_file,
                    mvke_file=mvke_file,
                    zplm_marc_file=zplm_marc,
                    zplm_mara_file=zplm_mara,
                    previous_report=None,
                    plant_filter=plant_filter_display
                )

                # Execução para exportação (se necessário)
                if plant_filter_export != plant_filter_display:
                    excel_export, _, _ = run_full_validation(
                        sap_file=sap_file,
                        mara_file=mara_file,
                        marc_file=marc_file,
                        mvke_file=mvke_file,
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

        st.markdown("### 📊 Passo 4 — Resultados")

        errors_df    = st.session_state.sap_errors_df
        summary_df   = st.session_state.sap_summary_df
        excel_export = st.session_state.sap_excel_export
        report_path  = st.session_state.sap_report_path

        # KPI Cards (usam classes do CSS externo)
        c = st.columns(3)
        with c[0]:
            st.markdown(
                f'<div class="card"><h4>Total de Erros</h4><div class="big">{errors_df.shape[0]}</div></div>',
                unsafe_allow_html=True
            )
        with c[1]:
            total_fields = int(errors_df["Field"].nunique()) if "Field" in errors_df.columns else 0
            st.markdown(
                f'<div class="card"><h4>Fields distintos</h4><div class="big">{total_fields}</div></div>',
                unsafe_allow_html=True
            )
        with c[2]:
            total_mats = int(errors_df["Material"].nunique()) if "Material" in errors_df.columns else 0
            st.markdown(
                f'<div class="card"><h4>Materiais com erro</h4><div class="big">{total_mats}</div></div>',
                unsafe_allow_html=True
            )

        # Tabelas
        st.header("All Errors")
        st.dataframe(errors_df, use_container_width=True)

        st.header("Adjusted Summary")
        st.dataframe(summary_df, use_container_width=True)

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
                        "Quando": log.get("timestamp"),
                        "Quem": log.get("user"),
                        "Filtro": log.get("plant_filter"),
                        "Erros": log.get("metrics", {}).get("total_errors", 0),
                    }
                    for log in logs
                ])
                st.dataframe(hist_df, use_container_width=True)
            else:
                st.write("Sem histórico.")

        # Nova execução
        c1, c2 = st.columns([1,1])
        with c1:
            if st.button("🔁 Nova Validação", use_container_width=True):
                # Opcional: manter uploads persistidos ou limpar
                st.session_state.sap_step = 1
                st.rerun()
        with c2:
            if st.button("🧹 Limpar uploads persistidos e recomeçar", use_container_width=True):
                st.session_state["sap_files_persisted"] = {}
                st.session_state.sap_step = 1
                st.rerun()

# ======================================================================================
#                                     PAGE 2
#                               PLM WORKFLOW (ATL)
# ======================================================================================
elif choice == "PLM Workflow (ATL)":
    st.subheader("🚗 PLM Workflow (ATL)")

    # Identidade
    effective_executor = identity_input("Quem executa?", "atl_exec")
    save_history = st.checkbox("Guardar esta execução no histórico", value=True)

    # -------------------- Uploaders --------------------
    uploads = st.file_uploader(
        "Largar todos os ficheiros necessários aqui",
        type="xlsx",
        accept_multiple_files=True
    )
    files = classify_uploaded_files(uploads)

    # Expander — mostra apenas os ficheiros relevantes à página atual
    with st.expander("🔍 Ficheiros identificados automaticamente (página atual)"):
        for key in required_keys_for_page(choice):
            val = files.get(key)
            label = LABELS_PT.get(key, key)
            st.write(f"**{label}** → {val.name if val else '❌ não encontrado'}")

    # Uploads
    auto_file  = files.get("auto")
    plm_file   = files.get("plm")
    timing_file = files.get("timing")

    # Filtros
    st.markdown("### 🧭 Âmbito de Visualização / Exportação")
    atl_scope = st.radio("Âmbito:", ["Tudo", "Apenas PT01", "Apenas PT02"], horizontal=True)
    atl_filter_export = {"Apenas PT01": "PT01", "Apenas PT02": "PT02"}.get(atl_scope)
    apply_filter_display_atl = st.checkbox("Aplicar filtro às tabelas apresentadas", value=False)
    atl_filter_display = atl_filter_export if apply_filter_display_atl else None

    if st.button("▶ Run ATL", type="primary"):
        if any(f is None for f in [auto_file, plm_file, timing_file]):
            st.error("⚠ Faltam ficheiros.")
            st.stop()

        with st.spinner("A construir PLM Workflow Combined..."):
            excel_all, df_pt01, df_pt02 = run_plm_workflow_combined(
                auto_file=auto_file, plm_file=plm_file, timing_file=timing_file
            )

        st.success("ATL concluído!")

        # -------------------- Guardar histórico --------------------
        if save_history:
            payload = build_atl_log_payload(
                user=effective_executor,
                plant_scope=atl_scope,
                df_pt01=df_pt01, df_pt02=df_pt02
            )
            save_run_log(payload)
            st.info("📁 Execução ATL guardada.")

        # Display
        if atl_filter_display == "PT01":
            df_display = {"PT01": df_pt01}
        elif atl_filter_display == "PT02":
            df_display = {"PT02": df_pt02}
        else:
            df_display = {"PT01": df_pt01, "PT02": df_pt02}

        for plant, df in df_display.items():
            st.header(plant)
            st.dataframe(df, use_container_width=True)
   
        # ---- Dashboards (execução atual ATL) ----
        st.markdown("## 📈 Dashboard — PLM Workflow (ATL)")
        render_atl_run_kpis(df_pt01, df_pt02)
        render_atl_run_charts(df_pt01, df_pt02)

        # ---- Evolução (histórico ATL) ----
        with st.expander("📉 Evolução temporal (Histórico ATL)"):
            logs = list_run_logs("atl", limit=50)
            render_atl_history_charts(logs)

        # Export
        if atl_filter_export == "PT01":
            df_export_pt01 = df_pt01; df_export_pt02 = None
            sheets = ("PT01",); suffix = "PT01"
        elif atl_filter_export == "PT02":
            df_export_pt01 = None; df_export_pt02 = df_pt02
            sheets = ("PT02",); suffix = "PT02"
        else:
            df_export_pt01 = df_pt01; df_export_pt02 = df_pt02
            sheets = ("PT01", "PT02"); suffix = "ALL"

        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine="openpyxl") as writer:
            if df_export_pt01 is not None:
                df_export_pt01.to_excel(writer, sheet_name="PT01", index=False)
            if df_export_pt02 is not None:
                df_export_pt02.to_excel(writer, sheet_name="PT02", index=False)
        out_excel.seek(0)

        out_excel = _apply_conditional_formatting_in_memory(out_excel, sheet_names=sheets)

        st.download_button(
            f"Download PLM_Workflow_{suffix}.xlsx",
            data=out_excel.getvalue(),
            file_name=f"PLM_Workflow_{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Histórico
        with st.expander("🕘 Histórico ATL (últimas 10 execuções)"):
            logs = list_run_logs("atl", limit=10)
            if logs:
                hist_df = pd.DataFrame([{
                    "Quando": log["timestamp"],
                    "Quem": log["user"],
                    "Âmbito": log["plant_scope"],
                    "PT01 rows": log["metrics"]["PT01_rows"],
                    "PT02 rows": log["metrics"]["PT02_rows"],
                } for log in logs])
                st.dataframe(hist_df, use_container_width=True)
            else:
                st.write("Sem histórico.")


# ======================================================================================
#                                     PAGE 3
#                                 MRP VALIDATION
# ======================================================================================
elif choice == "MRP Validation":
    st.subheader("🧩 MRP Validation")
    # Identidade (cross-page)
    effective_executor = identity_input("Quem executa?", "mrp_exec")
    save_history = st.checkbox("Guardar esta execução no histórico", value=True)

    # -------------------- Uploaders --------------------
    uploads = st.file_uploader(
        "Largar todos os ficheiros necessários aqui",
        type="xlsx",
        accept_multiple_files=True
    )
    files = classify_uploaded_files(uploads)

    # Expander — mostra apenas os ficheiros relevantes à página atual
    with st.expander("🔍 Ficheiros identificados automaticamente (página atual)"):
        for key in required_keys_for_page(choice):
            val = files.get(key)
            label = LABELS_PT.get(key, key)
            st.write(f"**{label}** → {val.name if val else '❌ não encontrado'}")

    # Uploaders
    mara_file = files.get("mara")
    marc_file = files.get("marc")
    zplm_marc = files.get("zplm_marc")

    # Filtro de Plant (visualização/exportação)
    st.markdown("### 🧭 Âmbito de Exportação")
    export_scope = st.radio(
        "Âmbito do download:",
        ["Tudo", "Apenas PT01", "Apenas PT02"],
        index=0, horizontal=True
    )
    plant_filter_export = {"Apenas PT01": "PT01", "Apenas PT02": "PT02"}.get(export_scope)

    apply_filter_display = st.checkbox("Aplicar filtro às tabelas apresentadas", value=False)
    plant_filter_display = plant_filter_export if apply_filter_display else None

    # RUN
    if st.button("▶ Run MRP Validation", type="primary"):
        missing = [mara_file, marc_file]
        if any(f is None for f in missing):
            st.error("⚠ Faltam ficheiros (MARA/MARC).")
            st.stop()

        with st.spinner("A validar regras MRP..."):
            # Execução para visualização
            excel_display, errors_df, summary_df = run_mrp_validation(
                mara_file=mara_file,
                marc_file=marc_file,
                zplm_marc_file=zplm_marc,
                plant_filter=plant_filter_display
            )

            # Execução para exportação (se filtro diferente do display)
            if plant_filter_export != plant_filter_display:
                excel_export, _, _ = run_mrp_validation(
                    mara_file=mara_file,
                    marc_file=marc_file,
                    zplm_marc_file=zplm_marc,
                    plant_filter=plant_filter_export
                )
            else:
                excel_export = excel_display

        st.success("Validação MRP concluída!")

        # Guardar histórico
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
            st.info("📁 Execução MRP guardada no histórico.")

        # Apresentação
        st.header("All Errors")
        st.dataframe(errors_df, use_container_width=True)

        st.header("Adjusted Summary")
        st.dataframe(summary_df, use_container_width=True)

        # Dashboards (execução atual)
        st.markdown("## 📈 Dashboard — MRP Validation")
        render_mrp_run_kpis(errors_df)
        render_mrp_run_charts(errors_df)

        # Evolução temporal (histórico MRP)
        with st.expander("📉 Evolução temporal (Histórico MRP)"):
            logs = list_run_logs("mrp", limit=50)
            render_mrp_history_charts(logs)

        # Download
        suffix = plant_filter_export if plant_filter_export else "ALL"
        st.download_button(
            "Download Excel (MRP)",
            data=excel_export.getvalue(),
            file_name=f"MRP_Validation_Adjusted_Report_{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Histórico (últimas 10 execuções)
        with st.expander("🕘 Histórico MRP (últimas 10 execuções)"):
            logs = list_run_logs("mrp", limit=10)
            if logs:
                hist_df = pd.DataFrame([{
                    "Quando": log["timestamp"],
                    "Quem": log["user"],
                    "Filtro": log["plant_filter"],
                    "Erros": log["metrics"]["total_errors"]
                } for log in logs])
                st.dataframe(hist_df, use_container_width=True)
            else:
                st.write("Sem histórico.")

# ======================================================================================
#                                     PAGE 4
#                          SAP MATERIAL VALIDATION PROGRESS
# ======================================================================================
else:
    import os, io
    import pandas as pd

    st.subheader("📈 SAP Material Validation Progress")

    # Identidade
    effective_executor = identity_input("Quem executa?", "prog_exec")
    save_history = st.checkbox("Guardar esta execução no histórico", value=True)

    # 🔁 NOVO: Modo Auto
    auto_mode = st.toggle("Auto (usar últimos relatórios guardados)", value=True)
    
    # Filtro de planta para Auto (para casarmos 'ALL', 'PT01' ou 'PT02' com o relatório arquivado)
    st.markdown("### 🧭 Âmbito (Auto)")
    prog_scope = st.radio("Âmbito:", ["ALL", "PT01", "PT02"], horizontal=True, index=0)
    
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
                cand.append((log.get("timestamp", ""), os.path.abspath(path)))

        # Se há 2 ou mais logs válidos → usar
        if len(cand) >= 2:
            cand.sort(key=lambda x: x[0], reverse=True)
            return cand[0][1], cand[1][1], len(cand)

        # 2) Fallback: varrer o vault
        base = os.path.join(".", "reports", "sap")
        month_dirs = [d for d in (os.listdir(base) if os.path.exists(base) else [])]
        paths = []
        for m in month_dirs:
            paths.extend(glob.glob(os.path.join(base, m, f"*SAP_Validation_Adjusted_Report_{scope}.xlsx")))

        # Ordenar por timestamp embutido no nome
        paths = sorted(paths, reverse=True)

        if len(paths) >= 2:
            return os.path.abspath(paths[0]), os.path.abspath(paths[1]), len(paths)

        if len(paths) == 1:
            return os.path.abspath(paths[0]), None, 1

        # Nenhum encontrado
        return None, None, 0

    # -------------------- Uploaders --------------------
    if auto_mode:  
        curr_path, prev_path, avail = pick_last_two_reports(prog_scope)

        with st.expander("🔍 Relatórios selecionados automaticamente"):
            st.write(f"**Disponíveis p/ {prog_scope}**: {avail}")
            st.write(f"**Atual** → {curr_path if curr_path else '❌ não encontrado'}")
            st.write(f"**Anterior** → {prev_path if prev_path else '❌ não encontrado'}")
            if avail == 1:
                st.info("Foi encontrado **1 relatório**. Gere mais **1** para poder comparar.")
            if avail == 0:
                st.warning("Ainda **não há relatórios** arquivados para este âmbito.")

        run_disabled = not (curr_path and prev_path)
        if st.button("▶ Run Progress (Auto)", type="primary", disabled=run_disabled):

            if not curr_path or not prev_path:
                st.error("⚠ Não foi possível encontrar dois relatórios arquivados para o âmbito selecionado.")
                st.stop()

            # Abrir como BytesIO (compatível com compare_adjusted_reports)
            with open(curr_path, "rb") as f:
                curr_bytes = io.BytesIO(f.read())
            with open(prev_path, "rb") as f:
                prev_bytes = io.BytesIO(f.read())

            from sap_validation_engine import compare_adjusted_reports
            with st.spinner("A comparar relatórios (Auto)..."):
                excel_diff, diff_dfs = compare_adjusted_reports(curr_bytes, prev_bytes)

            st.success("Progresso (Auto) concluído!")

            # Guardar histórico (nome = basename dos caminhos)
            if save_history:
                payload = build_progress_log_payload(
                    user=effective_executor,
                    current_name=os.path.basename(curr_path),
                    previous_name=os.path.basename(prev_path),
                    diff_dfs=diff_dfs
                )
                save_run_log(payload)
                st.info("📁 Progresso (Auto) guardado.")

            # Tabs (mesmo bloco do Manual)
            labels = ["Erros Novos", "Erros Resolvidos", "Erros Alterados", "Novos Materiais", "Views Perdidas"]
            keys   = ["New Errors", "Resolved Errors", "Changed Errors", "New Materials", "Views Lost"]
            tabs = st.tabs(labels)
            for tab, key in zip(tabs, keys):
                with tab:
                    st.dataframe(diff_dfs.get(key, pd.DataFrame()), use_container_width=True)

            # Dashboards + Download (iguais ao Manual)
            st.markdown("## 📈 Dashboard — Progress (Diff)")
            render_progress_kpis(diff_dfs)
            render_progress_charts(diff_dfs)
            st.markdown("---")
            st.download_button(
                "Download SAP_Validation_Progress.xlsx",
                data=excel_diff.getvalue(),
                file_name="SAP_Validation_Progress.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Histórico (iguais ao Manual)
            with st.expander("🕘 Histórico Progress (últimas 10 execuções)"):
                logs = list_run_logs("progress", limit=10)
                if logs:
                    hist_df = pd.DataFrame([{
                        "Quando": log["timestamp"],
                        "Quem": log["user"],
                        "New": log["metrics"]["new_errors"],
                        "Resolved": log["metrics"]["resolved_errors"],
                        "Changed": log["metrics"]["changed_errors"],
                        "New Mats": log["metrics"]["new_materials"],
                        "Views Lost": log["metrics"]["views_lost"],
                    } for log in logs])
                    st.dataframe(hist_df, use_container_width=True)
                else:
                    st.write("Sem histórico.")
    else:
        # ============ MODO MANUAL (agora com deteção automática por timestamp) ============
        uploads = st.file_uploader(
            "Largar os ficheiros (pelo menos 2) – sem renomear; eu detecto Atual/Anterior pelo timestamp do nome",
            type="xlsx",
            accept_multiple_files=True
        )

        # Mostrar os ficheiros que o utilizador colocou
        with st.expander("🔍 Ficheiros carregados (Manual)"):
            if uploads:
                for f in uploads:
                    st.write(f"• {f.name}")
            else:
                st.write("Ainda sem ficheiros carregados.")

        # Selecionar (Atual, Anterior) automaticamente
        curr_file, prev_file, notes = pick_current_previous_from_uploads(uploads or [])

        with st.expander("🧭 Seleção automática (Manual)"):
            if notes:
                for n in notes:
                    st.info(n)
            st.write(f"**Atual (mais recente)** → {getattr(curr_file, 'name', '❌ não selecionado')}")
            st.write(f"**Anterior** → {getattr(prev_file, 'name', '❌ não selecionado')}")

        # Botão com proteção (precisa de 2)
        run_disabled_manual = not (curr_file and prev_file)
        if st.button("▶ Run Progress (Manual)", type="primary", disabled=run_disabled_manual):
            from sap_validation_engine import compare_adjusted_reports
            with st.spinner("A comparar relatórios (Manual)..."):
                excel_diff, diff_dfs = compare_adjusted_reports(curr_file, prev_file)

            st.success("Progresso (Manual) concluído!")

            # Guardar histórico
            if save_history:
                payload = build_progress_log_payload(
                    user=effective_executor,
                    current_name=getattr(curr_file, "name", "current.xlsx"),
                    previous_name=getattr(prev_file, "name", "previous.xlsx"),
                    diff_dfs=diff_dfs
                )
                save_run_log(payload)
                st.info("📁 Progresso guardado.")

            # Tabs
            labels = ["Erros Novos", "Erros Resolvidos", "Erros Alterados", "Novos Materiais", "Views Perdidas"]
            keys   = ["New Errors", "Resolved Errors", "Changed Errors", "New Materials", "Views Lost"]
            tabs = st.tabs(labels)
            for tab, key in zip(tabs, keys):
                with tab:
                    st.dataframe(diff_dfs.get(key, pd.DataFrame()), use_container_width=True)

            # Dashboards (comparação atual)
            st.markdown("## 📈 Dashboard — Progress (Diff)")
            render_progress_kpis(diff_dfs)
            render_progress_charts(diff_dfs)

            # Download
            st.markdown("---")
            st.download_button(
                "Download SAP_Validation_Progress.xlsx",
                data=excel_diff.getvalue(),
                file_name="SAP_Validation_Progress.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Histórico (últimas 10 execuções)
            with st.expander("🕘 Histórico Progress (últimas 10 execuções)"):
                logs = list_run_logs("progress", limit=10)
                if logs:
                    hist_df = pd.DataFrame([{
                        "Quando": log["timestamp"],
                        "Quem": log["user"],
                        "New": log["metrics"]["new_errors"],
                        "Resolved": log["metrics"]["resolved_errors"],
                        "Changed": log["metrics"]["changed_errors"],
                        "New Mats": log["metrics"]["new_materials"],
                        "Views Lost": log["metrics"]["views_lost"],
                    } for log in logs])
                    st.dataframe(hist_df, use_container_width=True)
                else:
                    st.write("Sem histórico.")