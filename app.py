# app.py
import streamlit as st
import pandas as pd
import io

from sap_validation_engine import (
    run_full_validation,           # Relatório SAP Material Validation
    # nova função ATL:
    run_plm_workflow_combined      # Relatório PLM Workflow (ATL)
)

st.set_page_config(page_title="SAP & ATL Validation Hub", layout="wide")
st.title("🧰 Validation Hub")

# ---------------- Sidebar Menu ----------------
st.sidebar.header("Menu")
choice = st.sidebar.radio(
    "Escolher relatório:",
    ["SAP Material Validation", "PLM Workflow (ATL)"],
    index=0
)

# Pequena função de apoio para validar existência de folha
def has_sheet(uploaded, sheet_name: str):
    try:
        xl = pd.ExcelFile(uploaded)
        return sheet_name in xl.sheet_names, xl.sheet_names
    except Exception as e:
        return False, [f"Erro ao abrir: {e}"]

# ---------------- Página: SAP Material Validation ----------------
if choice == "SAP Material Validation":
    st.subheader("📦 SAP Material Validation Tool")

    sap_file      = st.file_uploader("Item Browse Report.xlsx (folha: SAPUI5 Export)", type="xlsx", key="sap")
    mara_file     = st.file_uploader("MARA.xlsx", type="xlsx", key="mara")
    marc_file     = st.file_uploader("MARC.xlsx", type="xlsx", key="marc")
    mvke_file     = st.file_uploader("MVKE.xlsx", type="xlsx", key="mvke")
    zplm_marc     = st.file_uploader("ZPLM MARC.xlsx (opcional)", type="xlsx", key="zplm_marc")
    zplm_mara     = st.file_uploader("ZPLM MARA.xlsx (opcional)", type="xlsx", key="zplm_mara")

    if st.button("▶ Run Validation", type="primary"):
        missing = [sap_file, mara_file, marc_file, mvke_file]
        if any(f is None for f in missing):
            st.error("⚠ Faltam ficheiros obrigatórios.")
        else:
            ok_sheet, sheets = has_sheet(sap_file, "SAPUI5 Export")
            if not ok_sheet:
                st.error(
                    f"⚠ A folha 'SAPUI5 Export' não foi encontrada no ficheiro enviado. "
                    f"Folhas disponíveis: {sheets}"
                )
            else:
                with st.spinner("A validar materiais..."):
                    excel_bytes, errors_df, summary_df = run_full_validation(
                        sap_file=sap_file,     # Item Browse Report.xlsx
                        mara_file=mara_file,   # sheet index 0
                        marc_file=marc_file,   # sheet index 0
                        mvke_file=mvke_file,   # sheet index 0
                        zplm_marc_file=zplm_marc,
                        zplm_mara_file=zplm_mara,
                        previous_report=None
                    )
                st.success("Validação concluída!")
                st.header("Erros Detetados (All Errors)")
                st.dataframe(errors_df, use_container_width=True)
                st.header("Resumo Ajustado (Adjusted Summary)")
                st.dataframe(summary_df, use_container_width=True)
                st.download_button(
                    "Download Excel",
                    data=excel_bytes.getvalue(),
                    file_name="SAP_Validation_Adjusted_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ---------------- Página: PLM Workflow (ATL) ----------------
else:
    st.subheader("🚗 PLM Workflow (ATL)")

    auto_file   = st.file_uploader("PLM Automotive Validation.xlsx", type="xlsx", key="atl_auto")
    plm_file    = st.file_uploader("PLM TimingLetter Validation.xlsx", type="xlsx", key="atl_plm")
    timing_file = st.file_uploader("Timing Letters.xlsx", type="xlsx", key="atl_timing")

    if st.button("▶ Run ATL", type="primary"):
        if any(f is None for f in [auto_file, plm_file, timing_file]):
            st.error("⚠ Faltam ficheiros obrigatórios (Automotive / TimingLetter / Timing Letters).")
        else:
            with st.spinner("A construir PLM Workflow Combined..."):
                excel_bytes, df_pt01, df_pt02 = run_plm_workflow_combined(
                    auto_file=auto_file,
                    plm_file=plm_file,
                    timing_file=timing_file
                )
            st.success("ATL concluído!")
            st.header("PT01")
            st.dataframe(df_pt01, use_container_width=True)
            st.header("PT02")
            st.dataframe(df_pt02, use_container_width=True)

            st.download_button(
                "Download PLM_Workflow_Combined.xlsx",
                data=excel_bytes.getvalue(),
                file_name="PLM_Workflow_Combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
