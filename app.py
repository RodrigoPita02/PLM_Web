
# app.py
import streamlit as st
import pandas as pd
import io

from utils.identify import get_executor_identity
from sap_validation_engine import (
    run_full_validation,
    run_plm_workflow_combined,
    _apply_conditional_formatting_in_memory,
    build_sap_log_payload,
    build_atl_log_payload,
    build_progress_log_payload,
    save_run_log,
    list_run_logs
)

# ==========================================================
# CONFIGURAÇÃO INICIAL
# ==========================================================
st.set_page_config(page_title="SAP & ATL Validation Hub", layout="wide")
st.title("🧰 Validation Hub")

# ==========================================================
# SIDEBAR MENU
# ==========================================================
st.sidebar.header("Menu")
choice = st.sidebar.radio(
    "Escolher relatório:",
    ["SAP Material Validation", "PLM Workflow (ATL)", "SAP Material Validation Progress"],
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
#                                        PAGE 1
#                             SAP MATERIAL VALIDATION
# ======================================================================================
if choice == "SAP Material Validation":
    st.subheader("📦 SAP Material Validation Tool")

    # -------------------- Identidade do utilizador --------------------
    effective_executor = identity_input("Quem executa?", "sap_exec")
    save_history = st.checkbox("Guardar esta execução no histórico", value=True)

    # -------------------- Uploaders --------------------
    sap_file  = st.file_uploader("Item Browse Report.xlsx (SAPUI5 Export)", type="xlsx", key="sap")
    mara_file = st.file_uploader("MARA.xlsx", type="xlsx", key="mara")
    marc_file = st.file_uploader("MARC.xlsx", type="xlsx", key="marc")
    mvke_file = st.file_uploader("MVKE.xlsx", type="xlsx", key="mvke")
    zplm_marc = st.file_uploader("ZPLM MARC.xlsx (opcional)", type="xlsx", key="zplm_marc")
    zplm_mara = st.file_uploader("ZPLM MARA.xlsx (opcional)", type="xlsx", key="zplm_mara")

    # -------------------- FILTRO DE PLANTA --------------------
    st.markdown("### 🧭 Âmbito de Exportação")
    export_scope = st.radio(
        "Âmbito do download:",
        ["Tudo", "Apenas PT01", "Apenas PT02"],
        index=0, horizontal=True
    )
    plant_filter_export = {"Apenas PT01": "PT01", "Apenas PT02": "PT02"}.get(export_scope)

    apply_filter_display = st.checkbox("Aplicar filtro às tabelas apresentadas", value=False)
    plant_filter_display = plant_filter_export if apply_filter_display else None

    # -------------------- RUN --------------------
    if st.button("▶ Run Validation", type="primary"):
        missing = [sap_file, mara_file, marc_file, mvke_file]
        if any(f is None for f in missing):
            st.error("⚠ Faltam ficheiros.")
            st.stop()

        ok_sheet, sheets = has_sheet(sap_file, "SAPUI5 Export")
        if not ok_sheet:
            st.error(f"A folha 'SAPUI5 Export' não existe. Folhas encontradas: {sheets}")
            st.stop()

        with st.spinner("A validar materiais..."):
            # Execução para display
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

            # Execução para exportação
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

        st.success("Validação concluída!")

        # -------------------- Guardar Histórico --------------------
        if save_history:
            payload = build_sap_log_payload(
                user=effective_executor,
                plant_filter=plant_filter_export,
                sap_file=sap_file, mara_file=mara_file, marc_file=marc_file, mvke_file=mvke_file,
                errors_df=errors_df, summary_df=summary_df
            )
            save_run_log(payload)
            st.info("📁 Execução guardada no histórico.")

        # -------------------- Apresentação --------------------
        st.header("All Errors")
        st.dataframe(errors_df, use_container_width=True)

        st.header("Adjusted Summary")
        st.dataframe(summary_df, use_container_width=True)

        # Download
        suffix = plant_filter_export if plant_filter_export else "ALL"
        st.download_button(
            "Download Excel",
            data=excel_export.getvalue(),
            file_name=f"SAP_Validation_Adjusted_Report_{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Histórico
        with st.expander("🕘 Histórico SAP (últimas 10 execuções)"):
            logs = list_run_logs("sap", limit=10)
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
#                                     PAGE 2
#                               PLM WORKFLOW (ATL)
# ======================================================================================
elif choice == "PLM Workflow (ATL)":
    st.subheader("🚗 PLM Workflow (ATL)")

    # Identidade
    effective_executor = identity_input("Quem executa?", "atl_exec")
    save_history = st.checkbox("Guardar esta execução no histórico", value=True)

    # Uploads
    auto_file  = st.file_uploader("PLM Automotive Validation.xlsx", type="xlsx")
    plm_file   = st.file_uploader("PLM TimingLetter Validation.xlsx", type="xlsx")
    timing_file = st.file_uploader("Timing Letters.xlsx", type="xlsx")

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
#                          SAP MATERIAL VALIDATION PROGRESS
# ======================================================================================
else:
    st.subheader("📈 SAP Material Validation Progress")

    # Identidade
    effective_executor = identity_input("Quem executa?", "prog_exec")
    save_history = st.checkbox("Guardar esta execução no histórico", value=True)

    diff_curr = st.file_uploader("Relatório Atual", type="xlsx", key="diff_curr")
    diff_prev = st.file_uploader("Relatório Anterior", type="xlsx", key="diff_prev")

    if st.button("▶ Run Progress", type="primary"):
        if diff_curr is None or diff_prev is None:
            st.error("⚠ Faltam ficheiros.")
            st.stop()

        from sap_validation_engine import compare_adjusted_reports

        with st.spinner("A comparar relatórios..."):
            excel_diff, diff_dfs = compare_adjusted_reports(diff_curr, diff_prev)

        st.success("Progresso concluído!")

        # Guardar histórico
        if save_history:
            payload = build_progress_log_payload(
                user=effective_executor,
                current_name=getattr(diff_curr, "name", "current.xlsx"),
                previous_name=getattr(diff_prev, "name", "previous.xlsx"),
                diff_dfs=diff_dfs
            )
            save_run_log(payload)
            st.info("📁 Progresso guardado.")

        # Tabs
        labels = ["Erros Novos", "Erros Resolvidos", "Erros Alterados", "Novos Materiais", "Views Perdidas"]
        keys =    ["New Errors",  "Resolved Errors",  "Changed Errors",  "New Materials",  "Views Lost"]

        tabs = st.tabs(labels)
        for tab, key in zip(tabs, keys):
            with tab:
                st.dataframe(diff_dfs.get(key, pd.DataFrame()), use_container_width=True)

        st.markdown("---")
        st.download_button(
            "Download SAP_Validation_Progress.xlsx",
            data=excel_diff.getvalue(),
            file_name="SAP_Validation_Progress.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

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