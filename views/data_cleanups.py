"""Vista: 🧹 Data Cleanups."""

import pandas as pd
import streamlit as st
from services.uploads import (
    LABELS_PT,
    REQUIRED_KEYS,
    classify_uploaded_files,
    persist_uploaded_mapping,
    rehydrate_persisted_mapping,
    validate_uploaded_files,
)
from ui.components import WIZARD_STEPS_DEFAULT, identity_input, render_wizard, show_file_status
from engines.common.helpers import _coerce_df_for_display
from engines.common.run_logs import build_dc_log_payload, save_run_log
from engines.data_cleanups.pipeline import run_data_cleanups_validation


def render(start_time: float) -> None:

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
        <strong>📋 What files do you need? (all 8 required)</strong><br><br>
        <table style="border-collapse:collapse;width:100%">
        <tr><td style="padding:3px 8px">📄 <strong>MARA</strong></td><td style="color:var(--app-fg-muted)">General material data — Material Type, Old Material Number</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MARC</strong></td><td style="color:var(--app-fg-muted)">Plant-level data — PSMS status, Production Storage Location</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>ZPLM MARC</strong></td><td style="color:var(--app-fg-muted)">PLM plant data — Material Plant Type, EOP Date, Service Liability Date, LTB fields</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>ZPLM MARA</strong></td><td style="color:var(--app-fg-muted)">PLM general data — Enterproj Number, Product Hierarchy</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>ZPLM_TA_IFS</strong></td><td style="color:var(--app-fg-muted)">IFS interface table — validates that every FP exists in the table</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MBEW</strong></td><td style="color:var(--app-fg-muted)">Valuation data — Standard Price per plant</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>EKPO</strong></td><td style="color:var(--app-fg-muted)">Purchase order items — open POs to close validation</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MARD</strong></td><td style="color:var(--app-fg-muted)">Storage location stock — obsolete materials with no stock</td></tr>
        </table>
        </div>
        """, unsafe_allow_html=True)

        uploads = st.file_uploader("Drop all 8 Excel files here",
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
            <div><span style="color:var(--app-fg-muted)">Files loaded</span><br><strong>{len(_pers_dc)} of 8</strong></div>
        </div>
        <div style="margin-top:10px;font-size:0.8rem;color:var(--app-fg-muted)">
            Will run: FP Level Status · EOP vs Service Liability · Standard Cost · LTB fields ·
            EKPO open POs · MARD stock · Cross-existence MARA/MARC ↔ ZPLM · FP existence in ZPLM_TA_IFS · PSMS 75/99 excluded
        </div>
        </div>
        """, unsafe_allow_html=True)

        files          = rehydrate_persisted_mapping("dc_files_persisted")
        mara_file      = files.get("mara")
        marc_file      = files.get("marc")
        zplm_marc_file = files.get("zplm_marc")
        zplm_mara_file = files.get("zplm_mara")
        zplm_ta_ifs_file = files.get("zplm_ta_ifs")
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
                    zplm_ta_ifs_file=zplm_ta_ifs_file,
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
                    st.dataframe(_df_dc, width='stretch')

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
