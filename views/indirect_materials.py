"""Vista: 📦 Indirect Materials Report."""

import os
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
from engines.indirect_materials.pipeline import (
    build_im_log_payload,
    run_indirect_materials_validation,
)
from ui.layout import render_footer
from ui.dashboard import render_sap_run_charts, render_sap_run_kpis
from engines.common.run_logs import save_report_to_vault, save_run_log


def render(start_time: float) -> None:

    st.subheader("📦 Indirect Materials Report")

    effective_executor = identity_input("Who performs?", "im_exec")
    save_history = st.checkbox("Save this execution in history", value=True, key="im_save_history")

    if "im_step" not in st.session_state:
        st.session_state.im_step = 1

    render_wizard(st.session_state.im_step, WIZARD_STEPS_DEFAULT)

    required_keys = REQUIRED_KEYS["📦 Indirect Materials Report"]
    persisted_now = st.session_state.get("im_files_persisted") or {}

    if st.session_state.im_step > 2 and any(k not in persisted_now for k in required_keys):
        st.session_state.im_step = 1
        st.rerun()

    # --- STEP 1: FILES ---
    if st.session_state.im_step == 1:
        st.markdown("### 📂 Step 1 — Load Files")
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--accent);border-radius:6px;padding:12px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>📋 Indirect Materials scope: ERSA · HIBE · VERP</strong><br><br>
        <table style="border-collapse:collapse;width:100%">
        <tr><td style="padding:3px 8px">📄 <strong>MARA</strong></td><td style="color:var(--app-fg-muted)">General material data (Material Type, weights, packaging)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MARC</strong></td><td style="color:var(--app-fg-muted)">Plant-level material data (MRP, purchasing, storage)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MVKE</strong></td><td style="color:var(--app-fg-muted)">Sales org data — VERP tax and item category</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MBEW</strong></td><td style="color:var(--app-fg-muted)">Valuation data (valuation class, price control)</td></tr>
        </table>
        </div>
        """, unsafe_allow_html=True)

        uploads = st.file_uploader(
            "Drop all necessary files here",
            type="xlsx",
            accept_multiple_files=True,
            key="im_multi_uploader",
        )

        files = classify_uploaded_files(uploads)
        if uploads:
            persist_uploaded_mapping(files, "im_files_persisted")

        persisted = st.session_state.get("im_files_persisted") or {}
        show_file_status(required_keys, files, persisted)

        _col_errs_im = validate_uploaded_files(persisted, required_keys)
        if _col_errs_im:
            for _k, _msg in _col_errs_im.items():
                st.warning(f"⚠️ **{LABELS_PT.get(_k, _k)}** — {_msg}")

        disabled = any(k not in persisted for k in required_keys) or bool(_col_errs_im)
        c1, c2 = st.columns(2)
        with c1:
            with st.expander("🧹 Clear uploaded files"):
                if st.button("Clear persistent uploads", key="im_clear"):
                    st.session_state["im_files_persisted"] = {}
                    st.success("Cleared.")
                    st.rerun()
        with c2:
            if st.button("➡️ Continue to Scope", disabled=disabled, key="im_next1", width='stretch'):
                st.session_state.im_step = 2
                st.rerun()

    # --- STEP 2: SCOPE ---
    elif st.session_state.im_step == 2:
        st.markdown("### 🌍 Step 2 — Select Scope")
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--color-ok);border-radius:6px;padding:10px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>ℹ️ Plant filter</strong> restricts validation to PT01, PT02, TN01 or TN02.
        Leave <em>All plants</em> to validate every plant present in the upload.
        </div>
        """, unsafe_allow_html=True)

        export_scope = st.radio(
            "Plant scope:",
            ["All plants", "Only PT01", "Only PT02", "Only TN01", "Only TN02"],
            horizontal=True,
            key="im_export_scope",
        )
        _plant_map = {
            "Only PT01": "PT01", "Only PT02": "PT02",
            "Only TN01": "TN01", "Only TN02": "TN02",
        }
        plant_filter_export = _plant_map.get(export_scope)
        apply_filter_display = st.checkbox(
            "Apply filter to displayed tables", value=False, key="im_apply_filter_display"
        )
        plant_filter_display = plant_filter_export if apply_filter_display else None

        st.session_state.im_plant_export = plant_filter_export
        st.session_state.im_plant_display = plant_filter_display

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="im_back2", width='stretch'):
                st.session_state.im_step = 1
                st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="im_next2", width='stretch'):
                st.session_state.im_selected_export_scope = export_scope
                st.session_state.im_selected_plant_filter_export = plant_filter_export
                st.session_state.im_selected_plant_filter_display = plant_filter_display
                st.session_state.im_step = 3
                st.rerun()

    # --- STEP 3: VALIDATE ---
    elif st.session_state.im_step == 3:
        st.markdown("### ⚙️ Step 3 — Run Validation")

        _scope_im = st.session_state.get(
            "im_selected_export_scope",
            st.session_state.get("im_export_scope", "All plants"),
        )
        _exec_im = st.session_state.get("executor_name", "—")
        _n_files = len(persisted_now)

        st.markdown(f"""
        <div style="background:var(--app-bg-secondary);border:1px solid var(--app-border);border-radius:10px;padding:16px 20px;margin-bottom:16px">
        <div style="font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--app-fg-muted);margin-bottom:10px">
            ✅ Ready to validate — Indirect Materials (ERSA / HIBE / VERP)
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;font-size:0.87rem;color:var(--app-fg)">
            <div><span style="color:var(--app-fg-muted)">Executor</span><br><strong>{_exec_im}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Plant Scope</span><br><strong>{_scope_im}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Files loaded</span><br><strong>{_n_files}</strong></div>
        </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("▶️ Validate now", type="primary", key="im_run"):
            files = rehydrate_persisted_mapping("im_files_persisted")
            mara_file = files.get("mara")
            marc_file = files.get("marc")
            mvke_file = files.get("mvke")
            mbew_file = files.get("mbew")

            plant_filter_export = st.session_state.get(
                "im_selected_plant_filter_export",
                st.session_state.get("im_plant_export"),
            )
            plant_filter_display = st.session_state.get(
                "im_selected_plant_filter_display",
                st.session_state.get("im_plant_display"),
            )

            _pb_slot = st.empty()
            progress_bar = _pb_slot.progress(0, text="Initializing validation...")

            def update_progress(percent: int, message: str = ""):
                progress_bar.progress(min(percent, 100), text=message)

            with st.spinner("Validating indirect materials..."):
                excel_display, errors_df, summary_df = run_indirect_materials_validation(
                    mara_file=mara_file,
                    marc_file=marc_file,
                    mvke_file=mvke_file,
                    mbew_file=mbew_file,
                    plant_filter=plant_filter_display,
                    progress_callback=update_progress,
                )

                if plant_filter_export != plant_filter_display:
                    excel_export, _, _ = run_indirect_materials_validation(
                        mara_file=mara_file,
                        marc_file=marc_file,
                        mvke_file=mvke_file,
                        mbew_file=mbew_file,
                        plant_filter=plant_filter_export,
                    )
                else:
                    excel_export = excel_display

            suffix = plant_filter_export or "ALL"
            report_path = save_report_to_vault(excel_export, report_type="indirect", plant_suffix=suffix)
            st.session_state.im_report_path = report_path
            st.session_state.im_errors_df = errors_df
            st.session_state.im_summary_df = summary_df
            st.session_state.im_excel_export = excel_export

            if save_history:
                payload = build_im_log_payload(
                    user=effective_executor,
                    plant_filter=plant_filter_export or suffix,
                    mara_file=mara_file,
                    marc_file=marc_file,
                    mvke_file=mvke_file,
                    mbew_file=mbew_file,
                    errors_df=errors_df,
                    summary_df=summary_df,
                    report_path=report_path,
                )
                save_run_log(payload)

            _pb_slot.empty()
            st.session_state.im_step = 4
            st.rerun()

    # --- STEP 4: RESULTS ---
    elif st.session_state.im_step == 4:
        st.markdown("### 📊 Step 4 — Results")

        errors_df = st.session_state.im_errors_df
        summary_df = st.session_state.im_summary_df
        excel_export = st.session_state.im_excel_export
        report_path = st.session_state.im_report_path

        _fe_cnt = int(errors_df.shape[0])
        _f_cnt = int(errors_df["Field"].nunique()) if "Field" in errors_df.columns else 0
        _m_cnt = int(errors_df["Material"].nunique()) if "Material" in errors_df.columns else 0

        c = st.columns(3)
        with c[0]:
            st.markdown(
                f'<div class="card"><h4>Field Errors</h4><div class="big">{_fe_cnt:,}</div></div>',
                unsafe_allow_html=True,
            )
        with c[1]:
            st.markdown(
                f'<div class="card"><h4>Distinct Fields</h4><div class="big">{_f_cnt}</div></div>',
                unsafe_allow_html=True,
            )
        with c[2]:
            st.markdown(
                f'<div class="card"><h4>Materials with Errors</h4><div class="big">{_m_cnt:,}</div></div>',
                unsafe_allow_html=True,
            )

        _top_field = (
            errors_df["Field"].value_counts().index[0]
            if (not errors_df.empty and "Field" in errors_df.columns)
            else "—"
        )
        _top_plant = (
            errors_df["Plant"].astype(str).str.strip().replace("", "General")
            .value_counts().index[0]
            if (not errors_df.empty and "Plant" in errors_df.columns)
            else "—"
        )
        st.markdown(f"""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--accent);border-radius:8px;
                    padding:12px 16px;margin-bottom:16px;font-size:0.9rem;line-height:1.6;color:var(--app-fg)">
            <strong>📋 Indirect Materials Summary</strong><br>
            Found <strong>{_fe_cnt:,} errors</strong> across <strong>{_m_cnt:,} materials</strong>
            in <strong>{_f_cnt} distinct fields</strong>.
            Most affected field: <strong>{_top_field}</strong> · Most affected plant: <strong>{_top_plant}</strong>.
        </div>
        """, unsafe_allow_html=True)

        st.markdown("## 📈 Dashboard — Indirect Materials")
        render_sap_run_kpis(errors_df)
        render_sap_run_charts(errors_df)

        download_name = (
            os.path.basename(report_path)
            if report_path
            else "Indirect_Materials_Report_ALL.xlsx"
        )
        st.download_button(
            "⬇️ Download Excel",
            data=excel_export.getvalue(),
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("📋 All Errors", expanded=False):
            st.dataframe(_coerce_df_for_display(errors_df), width="stretch")

        with st.expander("📊 Adjusted Summary", expanded=False):
            st.dataframe(_coerce_df_for_display(summary_df), width="stretch")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Validation", key="im_new", width='stretch'):
                st.session_state.im_step = 1
                st.rerun()
        with c2:
            if st.button("🧹 Clear & Restart", key="im_clear4", width='stretch'):
                st.session_state["im_files_persisted"] = {}
                st.session_state.im_step = 1
                st.rerun()

    render_footer(start_time, version="v1.6")
