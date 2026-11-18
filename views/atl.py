"""Vista: 🚗 PLM Workflow' and sub_choice."""

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
from ui.components import (
    WIZARD_STEPS_DEFAULT,
    identity_input,
    render_history_with_download,
    render_wizard,
    show_file_status,
)
from engines.common.helpers import _coerce_df_for_display
from engines.common.run_logs import build_atl_log_payload, save_run_log
from ui.dashboard import render_atl_run_charts, render_atl_run_kpis
from ui.layout import render_footer
from engines.plm_workflow.pipeline import run_plm_workflow_combined


def render(start_time: float) -> None:

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
        <tr><td style="padding:3px 8px;color:var(--color-ok)">🔓 <strong>C_WORKFLOWTASK</strong></td><td style="color:var(--app-fg-muted)">Optional — enables accurate group X marks from Work item text (READY/SELECTED status only)</td></tr>
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
        show_file_status(required_keys, files, persisted,
                         optional_keys=["workflow_task"])

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
            <div><span style="color:var(--app-fg-muted)">Files loaded</span><br><strong>{len(_pers_atl)} of 3{"  +  C_WORKFLOWTASK ✅" if "workflow_task" in _pers_atl else ""}</strong></div>
        </div>
        <div style="margin-top:10px;font-size:0.8rem;color:var(--app-fg-muted)">
            The workflow will match Change Numbers from Automotive and TimingLetter sources,
            filter active Timing Letters (excludes Complete/Rejected), and compute Aging for each CN.
        </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("▶️ Run Workflow ATL", type="primary", key="atl_run"):
            files = rehydrate_persisted_mapping("atl_files_persisted")
            auto_file      = files.get("auto")
            plm_file       = files.get("plm")
            timing_file    = files.get("timing")
            wft_file       = files.get("workflow_task")

            if any(f is None for f in [auto_file, plm_file, timing_file]):
                st.error("⚠️ Required files are missing.")
                st.stop()

            with st.spinner("Processing ATL Workflow..."):
                excel_all, df_pt01, df_pt02 = run_plm_workflow_combined(
                    auto_file=auto_file, plm_file=plm_file, timing_file=timing_file,
                    workflowtask_file=wft_file,
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
            st.dataframe(_df_pt01_s, width='stretch')

        with st.expander("📋 PT02 Change Numbers", expanded=False):
            _srch_pt02 = st.text_input("🔍 Filter PT02", key="plm_pt02_search")
            _df_pt02_s = _coerce_df_for_display(df_pt02) if df_pt02 is not None else pd.DataFrame()
            if _srch_pt02.strip() and not _df_pt02_s.empty:
                _mask_p02 = _df_pt02_s.apply(lambda col: col.astype(str).str.contains(_srch_pt02.strip(), case=False, na=False)).any(axis=1)
                _df_pt02_s = _df_pt02_s[_mask_p02]
            st.caption(f"{len(_df_pt02_s):,} rows")
            st.dataframe(_df_pt02_s, width='stretch')

        suffix = st.session_state.atl_filter_export or "ALL"
        st.download_button(
            f"⬇️ Download PLM_Workflow_{suffix}.xlsx",
            data=excel_all.getvalue(), file_name=f"PLM_Workflow_{suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        render_history_with_download(
            "atl", limit=10, label="ATL History",
            col_map={"When": "timestamp", "Who": "user",
                     "PT01": "metrics.PT01_rows", "PT02": "metrics.PT02_rows"}
        )

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

    render_footer(start_time, version="v1.4")
