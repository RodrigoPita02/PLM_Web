"""Vista: 📋 SSA Validation."""

import pandas as pd
import streamlit as st
from ui.components import WIZARD_STEPS_DEFAULT, identity_input, render_wizard
from engines.common.helpers import _coerce_df_for_display
from services.uploads import persist_uploaded_mapping, rehydrate_persisted_mapping
from ui.layout import render_footer
from engines.ssa.pipeline import run_ssa_validation


def render(start_time: float) -> None:

    st.subheader("📋 Schedule Agreement (SSA) Validation")
    effective_executor = identity_input("Who performs?", "ssa_exec")

    if "ssa_step" not in st.session_state:
        st.session_state.ssa_step = 1

    render_wizard(st.session_state.ssa_step, WIZARD_STEPS_DEFAULT)

    ssa_required = ["vbak", "vbap", "vbkd", "kna1"]
    ssa_optional = ["marc", "vbpa", "vbfa", "vbep", "mvke"]

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
        <tr><td style="padding:3px 8px;color:var(--color-ok)">🔓 <strong>MARC</strong></td><td style="color:var(--app-fg-muted)">Optional — Material Status 70/75/99, MRP Type checks</td></tr>
        <tr><td style="padding:3px 8px;color:var(--color-ok)">🔓 <strong>VBPA</strong></td><td style="color:var(--app-fg-muted)">Optional — Partner country override for DDP+GB rule, Unloading Point</td></tr>
        <tr><td style="padding:3px 8px;color:var(--color-ok)">🔓 <strong>VBFA</strong></td><td style="color:var(--app-fg-muted)">Optional — Sales Document Flow, gates the Weights rule on OBD existence for non-ZLZ/ZLZM documents</td></tr>
        <tr><td style="padding:3px 8px;color:var(--color-ok)">🔓 <strong>VBEP</strong></td><td style="color:var(--app-fg-muted)">Optional — Schedule lines, validates active demand for ZLZ/ZLZM using Release Type = 1</td></tr>
        <tr><td style="padding:3px 8px;color:var(--color-ok)">🔓 <strong>MVKE</strong></td><td style="color:var(--app-fg-muted)">Optional — Acct Assmt Grp Mat. lookup for S-materials in Mandatory Fields</td></tr>
        </table>
        </div>
        """, unsafe_allow_html=True)

        uploads = st.file_uploader(
            "Drop VBAK, VBAP, VBKD, KNA1 + optional MARC, VBPA, VBFA, VBEP and MVKE",
            type="xlsx", accept_multiple_files=True, key="ssa_uploader",
        )

        def classify_ssa_files(files):
            mapping = {k: None for k in ["vbak","vbap","vbkd","vbpa","kna1","marc","vbfa","vbep","mvke"]}
            for f in (files or []):
                name = f.name.lower()
                if "vbak" in name:   mapping["vbak"] = f
                elif "vbap" in name: mapping["vbap"] = f
                elif "vbkd" in name: mapping["vbkd"] = f
                elif "vbpa" in name: mapping["vbpa"] = f
                elif "kna1" in name: mapping["kna1"] = f
                elif "marc" in name: mapping["marc"] = f
                elif "vbfa" in name: mapping["vbfa"] = f
                elif "vbep" in name: mapping["vbep"] = f
                elif "mvke" in name: mapping["mvke"] = f
            return mapping

        files_map = classify_ssa_files(uploads)
        if uploads:
            persist_uploaded_mapping(files_map, "ssa_files_persisted")

        persisted = st.session_state.get("ssa_files_persisted") or {}

        with st.expander("📋 File Status", expanded=True):
            label_map = {
                "vbak": "VBAK.xlsx", "vbap": "VBAP.xlsx",
                "vbkd": "VBKD.xlsx", "kna1": "KNA1.xlsx",
                "marc": "MARC.xlsx (optional)", "vbpa": "VBPA.xlsx (optional)",
                "vbfa": "VBFA.xlsx (optional)", "vbep": "VBEP.xlsx (optional)",
                "mvke": "MVKE.xlsx (optional)"
            }
            for key in ["vbak","vbap","vbkd","kna1","marc","vbpa","vbfa","vbep","mvke"]:
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
            if st.button("➡️ Continue to Scope", disabled=disabled, key="ssa_next1", width='stretch'):
                st.session_state.ssa_step = 2; st.rerun()

    elif st.session_state.ssa_step == 2:
        st.markdown("### 🌍 Step 2 — Choose Scope")
        ssa_scope = st.radio("Filter by Plant:", ["All Plants", "Only PT01", "Only PT02"], horizontal=True)
        st.session_state.ssa_plant_filter = {"Only PT01": "PT01", "Only PT02": "PT02"}.get(ssa_scope)
        st.session_state.ssa_scope = ssa_scope

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="ssa_back2", width='stretch'):
                st.session_state.ssa_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue", key="ssa_next2", width='stretch'):
                st.session_state.ssa_step = 3; st.rerun()

    elif st.session_state.ssa_step == 3:
        st.markdown("### ⚙️ Step 3 — Run SSA Validation")

        _exec_ssa  = st.session_state.get("executor_name", "—")
        _scope_ssa = st.session_state.get("ssa_scope", "All Plants")
        _pers_ssa  = st.session_state.get("ssa_files_persisted") or {}
        _has_marc  = "marc" in _pers_ssa
        _has_vbpa  = "vbpa" in _pers_ssa
        _has_vbfa  = "vbfa" in _pers_ssa
        _has_vbep  = "vbep" in _pers_ssa
        _has_mvke  = "mvke" in _pers_ssa

        st.markdown(f"""
        <div style="background:var(--app-bg-secondary);border:1px solid var(--app-border);border-radius:10px;padding:16px 20px;margin-bottom:16px">
        <div style="font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--app-fg-muted);margin-bottom:10px">
            ✅ Ready to validate
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;font-size:0.87rem;color:var(--app-fg)">
            <div><span style="color:var(--app-fg-muted)">Executor</span><br><strong>{_exec_ssa}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Scope</span><br><strong>{_scope_ssa}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Optional files</span><br><strong>{"MARC ✅" if _has_marc else "MARC ⬜"} &nbsp; {"VBPA ✅" if _has_vbpa else "VBPA ⬜"} &nbsp; {"VBFA ✅" if _has_vbfa else "VBFA ⬜"} &nbsp; {"VBEP ✅" if _has_vbep else "VBEP ⬜"} &nbsp; {"MVKE ✅" if _has_mvke else "MVKE ⬜"}</strong></div>
        </div>
        <div style="margin-top:10px;font-size:0.8rem;color:var(--app-fg-muted)">
            Rules: Mandatory Fields{" (S-material Acct Assmt Grp Mat. checked against MVKE)" if _has_mvke else ""} · Incoterms · Incompletion Log · Terms of Payment · Shipping Point ·
            Unloading Point Consistency · Customer Reference · Weights{" (ZLZ/ZLZM demand-gated via VBEP Release Type = 1)" if _has_vbep else " (ZLZ/ZLZM skipped without usable VBEP)"}{"; other types OBD-gated via VBFA" if _has_vbfa else ""} ·
            Order Without Items · Customer Material{"· Rejected Status 70/75/99 · MRP Type" if _has_marc else ""}
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
                    vbfa_file=files.get("vbfa"),
                    vbep_file=files.get("vbep"),
                    mvke_file=files.get("mvke"),
                    plant_filter=st.session_state.get("ssa_plant_filter"),
                    apply_styles=False,
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
                    st.dataframe(df_show, width='stretch')

        with st.expander("📋 Summary", expanded=False):
            st.dataframe(_coerce_df_for_display(summary_df), width='stretch')

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
            if st.button("🔄 New Validation", key="ssa_new", width='stretch'):
                st.session_state.ssa_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear & Restart", key="ssa_clear4", width='stretch'):
                st.session_state["ssa_files_persisted"] = {}
                st.session_state.ssa_step = 1; st.rerun()

    render_footer(start_time, version="v1.3")
    render_footer(start_time, version="v1.4")
