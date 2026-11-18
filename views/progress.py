"""Vista: 📊 SAP Validation Progress."""

import io
import os
import pandas as pd
import streamlit as st
from ui.components import (
    WIZARD_STEPS_PROGRESS,
    identity_input,
    render_history_with_download,
    render_wizard,
)
from engines.common.helpers import _coerce_df_for_display
from engines.common.run_logs import build_progress_log_payload, save_run_log
from services.uploads import pick_current_previous_from_uploads, pick_last_two_reports
from ui.layout import render_footer
from ui.dashboard import render_progress_charts, render_progress_kpis


def render(start_time: float) -> None:

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
            from engines.progress.compare import compare_adjusted_reports

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
                st.dataframe(_df_prog, width='stretch')

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
            if st.button("🔄 New Comparison", key="prog_new", width='stretch'):
                st.session_state.progress_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear all & restart", key="prog_clear4", width='stretch'):
                for k in ["progress_curr_path", "progress_prev_path",
                          "progress_manual_curr", "progress_manual_prev",
                          "progress_diff_dfs", "progress_diff_excel"]:
                    st.session_state.pop(k, None)
                st.session_state.progress_step = 1; st.rerun()

    render_footer(start_time, version="v1.4")
