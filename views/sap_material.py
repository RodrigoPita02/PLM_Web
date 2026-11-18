"""Vista: 📦 SAP Material Validation."""

import os
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
from engines.common.rolesviews import get_rolesviews_bytesio as _safe_rolesviews_bytesio
from engines.assignments.history import build_owner_dashboard_df, track_error_count
from engines.common.run_logs import build_sap_log_payload, save_report_to_vault, save_run_log
from engines.assignments.store import merge_assignments, save_open_assignments
from ui.assignment_panel import render_assignment_panel
from ui.layout import render_footer
from ui.dashboard import render_sap_run_charts, render_sap_run_kpis
from engines.sap_material.pipeline import run_full_validation
from engines.tunisia.pipeline import run_tn_validation


def render(start_time: float) -> None:

    st.subheader("📦 SAP Material Validation Tool")

    effective_executor = identity_input("Who performs?", "sap_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    if "sap_step" not in st.session_state:
        st.session_state.sap_step = 1

    render_wizard(st.session_state.sap_step, WIZARD_STEPS_DEFAULT)

    required_keys = REQUIRED_KEYS["📦 SAP Material Validation"]
    persisted_now = st.session_state.get("sap_files_persisted") or {}

    # Guard: se ficheiros em falta volta ao passo 1
    if st.session_state.sap_step > 2 and any(k not in persisted_now for k in required_keys):
        st.session_state.sap_step = 1
        st.rerun()

    # --- STEP 1: FILES ---
    if st.session_state.sap_step == 1:
        st.markdown("### 📂 Step 1 — Load Files")
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--accent);border-radius:6px;padding:12px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>📋 What files do you need?</strong><br><br>
        <table style="border-collapse:collapse;width:100%">
        <tr><td style="padding:3px 8px">📄 <strong>MARA</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → MARA (general material data)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MARC</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → MARC (plant-level material data)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MVKE</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → MVKE (sales organisation data)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>ZPLM MARC / MARA</strong></td><td style="color:var(--app-fg-muted)">PLM custom tables (ZPLM_MARC / ZPLM_MARA export)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MARM</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → MARM (units of measure)</td></tr>
        <tr><td style="padding:3px 8px">📄 <strong>MBEW</strong></td><td style="color:var(--app-fg-muted)">SAP table SE16 → MBEW (valuation data)</td></tr>
        <tr><td style="padding:3px 8px; color:var(--color-ok)">🔓 <strong>CSR Item / Head</strong></td><td style="color:var(--app-fg-muted)">Optional — enables CSR exemption rules for MRP Type</td></tr>
        <tr><td style="padding:3px 8px; color:var(--color-ok)">🔓 <strong>MATPLSTA</strong></td><td style="color:var(--app-fg-muted)">Optional — enables Planning Status = 9 exemption for ROH Per.Lot Sz.Plng Cal. rule</td></tr>
        <tr><td style="padding:3px 8px; color:var(--color-ok)">🔓 <strong>MARD</strong></td><td style="color:var(--app-fg-muted)">Optional — enables stock-aware MRP Controller rule (PSMS 70/75 only require OB1/OB2 when there's no stock)</td></tr>
        </table>
        </div>
        """, unsafe_allow_html=True)

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
                         optional_keys=["csr_item", "csr_head", "matplsta", "mard"])

        # Quick column validation
        _col_errs_sap = validate_uploaded_files(persisted, required_keys)
        if _col_errs_sap:
            for _k, _msg in _col_errs_sap.items():
                st.warning(f"⚠️ **{LABELS_PT.get(_k, _k)}** — {_msg}")

        disabled = any(k not in persisted for k in required_keys) or bool(_col_errs_sap)
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
        st.markdown("""
        <div style="background:var(--app-bg-secondary);border-left:4px solid var(--color-ok);border-radius:6px;padding:10px 16px;margin-bottom:12px;font-size:0.85rem;color:var(--app-fg)">
        <strong>ℹ️ What does scope mean?</strong><br>
        <strong>Site</strong> controls which factory is validated — Palmela (PT01/PT02), Tunisia (TN01/TN02), or both.<br>
        <strong>Plant filter</strong> further restricts within the site — select a specific plant or keep All to include both plants of the chosen site.<br>
        <strong>Apply filter to displayed tables</strong> also filters the results shown here in the app — leave unchecked to see all plants in the UI while exporting only the chosen scope.
        </div>
        """, unsafe_allow_html=True)

        # Site selector — on_change resets plant scope to avoid stale values
        def _on_site_change():
            if "sap_export_scope" in st.session_state:
                del st.session_state["sap_export_scope"]

        site_scope = st.radio(
            "Site:",
            ["Palmela (PT01 / PT02)", "Tunisia (TN01 / TN02)", "All Sites"],
            horizontal=True,
            key="sap_site_scope",
            on_change=_on_site_change,
        )

        # Plant filter within selected site
        if site_scope == "Palmela (PT01 / PT02)":
            plant_options = ["All Palmela plants", "Only PT01", "Only PT02"]
        elif site_scope == "Tunisia (TN01 / TN02)":
            plant_options = ["All Tunisia plants", "Only TN01", "Only TN02"]
        else:
            plant_options = ["All plants", "Only PT01", "Only PT02", "Only TN01", "Only TN02"]

        export_scope = st.radio("Plant scope for download:", plant_options, horizontal=True, key="sap_export_scope")

        _plant_map = {
            "Only PT01": "PT01", "Only PT02": "PT02",
            "Only TN01": "TN01", "Only TN02": "TN02",
        }
        plant_filter_export = _plant_map.get(export_scope)

        apply_filter_display = st.checkbox("Apply filter to displayed tables", value=False, key="sap_apply_filter_display")
        plant_filter_display = plant_filter_export if apply_filter_display else None

        # Only store derived (non-widget) values — sap_site_scope and sap_export_scope
        # are already managed by Streamlit via their widget keys.
        st.session_state.sap_plant_export  = plant_filter_export
        st.session_state.sap_plant_display = plant_filter_display

        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="sap_back2", width='stretch'):
                st.session_state.sap_step = 1
                st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="sap_next2", width='stretch'):
                st.session_state["sap_selected_site_scope"]   = site_scope
                st.session_state["sap_selected_export_scope"] = export_scope
                st.session_state["sap_selected_plant_filter_export"] = plant_filter_export
                st.session_state["sap_selected_plant_filter_display"] = plant_filter_display
                st.session_state.sap_step = 3
                st.rerun()

    # --- STEP 3: VALIDATE ---
    elif st.session_state.sap_step == 3:
        st.markdown("### ⚙️ Step 3 — Run Validation")

        # Confirmation summary
        _persisted3 = st.session_state.get("sap_files_persisted") or {}
        _scope3     = st.session_state.get(
            "sap_selected_export_scope",
            st.session_state.get("sap_export_scope", "All")
        )
        _site3      = st.session_state.get(
            "sap_selected_site_scope",
            st.session_state.get("sap_site_scope", "Palmela (PT01 / PT02)")
        )
        _exec3      = st.session_state.get("executor_name", "—")
        _n_files    = len(_persisted3)
        _optional3  = [k for k in ["csr_item","csr_head","matplsta"] if k in _persisted3]

        # Determine which engine will run
        _is_tn   = "Tunisia" in _site3
        _is_both = "All Sites" in _site3
        _engine_label = "Tunisia Engine (TN01/TN02)" if _is_tn else ("Both Engines (PT + TN)" if _is_both else "Palmela Engine (PT01/PT02)")

        st.markdown(f"""
        <div style="background:var(--app-bg-secondary);border:1px solid var(--app-border);border-radius:10px;padding:16px 20px;margin-bottom:16px">
        <div style="font-size:0.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--app-fg-muted);margin-bottom:10px">
            ✅ Ready to validate — confirm before running
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:10px;font-size:0.87rem;color:var(--app-fg)">
            <div><span style="color:var(--app-fg-muted)">Executor</span><br><strong>{_exec3}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Site</span><br><strong>{_site3}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Plant Scope</span><br><strong>{_scope3}</strong></div>
            <div><span style="color:var(--app-fg-muted)">Engine</span><br><strong>{_engine_label}</strong></div>
        </div>
        <div style="margin-top:8px;font-size:0.8rem;color:var(--app-fg-muted)">Files loaded: <strong>{_n_files} required{"  +  " + str(len(_optional3)) + " optional" if _optional3 else ""}</strong></div>
        {"<div style='margin-top:6px;font-size:0.8rem;color:var(--color-ok)'>🔓 CSR exemption rules <strong>enabled</strong> — active CSRs will exempt FP materials from MRP = ND errors.</div>" if _optional3 else "<div style='margin-top:6px;font-size:0.8rem;color:var(--app-fg-muted)'>🔓 CSR files not uploaded — CSR exemption rules disabled.</div>"}
        {"<div style='margin-top:4px;font-size:0.8rem;color:var(--color-ok)'>🔓 MATPLSTA loaded — ROH Planning Status = 9 exemption <strong>enabled</strong>.</div>" if "matplsta" in _persisted3 else "<div style='margin-top:4px;font-size:0.8rem;color:var(--app-fg-muted)'>🔓 MATPLSTA not uploaded — all ROH with Per.Lot Sz.Plng Cal. filled will be flagged.</div>"}
        {"<div style='margin-top:4px;font-size:0.8rem;color:var(--color-ok)'>🔓 MARD loaded — MRP Controller rule is <strong>stock-aware</strong> (PSMS 70/75 only require OB1/OB2 when there's no stock).</div>" if "mard" in _persisted3 else "<div style='margin-top:4px;font-size:0.8rem;color:var(--app-fg-muted)'>🔓 MARD not uploaded — MRP Controller rule uses the legacy behaviour (PSMS 70 always requires OB1, PSMS 75 always requires OB2).</div>"}
        {"<div style='margin-top:6px;font-size:0.8rem;color:var(--accent)'>⚠️ Both Sites mode runs the Palmela engine (PT01/PT02) + Tunisia engine (TN01/TN02), both sourced from MARC/MARA/MVKE tables, and merges results.</div>" if _is_both else ""}
        </div>
        """, unsafe_allow_html=True)

        if st.button("▶️ Validate now", type="primary", key="sap_run"):
            files = rehydrate_persisted_mapping("sap_files_persisted")
            mara_file = files.get("mara")
            marc_file = files.get("marc")
            mvke_file = files.get("mvke")
            zplm_marc = files.get("zplm_marc")
            zplm_mara = files.get("zplm_mara")
            marm_file = files.get("marm")
            mbew_file = files.get("mbew")
            csr_item  = files.get("csr_item")
            csr_head  = files.get("csr_head")
            matplsta  = files.get("matplsta")
            mard_file = files.get("mard")

            plant_filter_export = st.session_state.get(
                "sap_selected_plant_filter_export",
                st.session_state.get("sap_plant_export")
            )
            plant_filter_display = st.session_state.get(
                "sap_selected_plant_filter_display",
                st.session_state.get("sap_plant_display")
            )
            site_scope = st.session_state.get(
                "sap_selected_site_scope",
                st.session_state.get("sap_site_scope", "Palmela (PT01 / PT02)")
            )

            run_palmela = site_scope in ("Palmela (PT01 / PT02)", "All Sites")
            run_tn = site_scope in ("Tunisia (TN01 / TN02)", "All Sites")

            _pb_slot = st.empty()
            progress_bar = _pb_slot.progress(0, text="Initializing validation...")

            def update_progress(percent: int, message: str = ""):
                progress_bar.progress(min(percent, 100), text=message)

            with st.spinner("Validating materials..."):

                # ── TUNISIA ONLY ───────────────────────────────────────────────
                if run_tn and not run_palmela:
                    excel_display, errors_df, summary_df = run_tn_validation(
                        mara_file=mara_file, marc_file=marc_file, mvke_file=mvke_file,
                        marm_file=marm_file, mbew_file=mbew_file,
                        zplm_marc_file=zplm_marc, zplm_mara_file=zplm_mara,
                        previous_report=None,
                        plant_filter=plant_filter_display,
                        progress_callback=update_progress,
                        csr_item_file=csr_item, csr_head_file=csr_head,
                        matplsta_file=matplsta, mard_file=mard_file,
                    )
                    excel_export = excel_display

                # ── BOTH SITES: Palmela + Tunisia, merged ──────────────────────
                elif run_tn and run_palmela:
                    update_progress(2, "Running Palmela engine...")

                    _pt_filter = plant_filter_export if plant_filter_export in ("PT01", "PT02") else None
                    excel_pt, errors_pt, _ = run_full_validation(
                        mara_file=mara_file, marc_file=marc_file,
                        mvke_file=mvke_file,
                        zplm_marc_file=zplm_marc, zplm_mara_file=zplm_mara,
                        marm_file=marm_file, previous_report=None,
                        plant_filter=_pt_filter,
                        progress_callback=lambda p, m: update_progress(int(p * 0.45), m),
                        csr_item_file=csr_item, csr_head_file=csr_head,
                        matplsta_file=matplsta, mbew_file=mbew_file,
                        mard_file=mard_file,
                    )

                    # Re-use files already parsed by Palmela engine via cache
                    update_progress(50, "Running Tunisia engine (reusing loaded files)...")
                    from engines.common.fast_loader import load_xlsx_fast as _lxf2
                    _shared = dict(
                        _mara_df      = _lxf2(mara_file,  sheet_name=0),
                        _marc_df      = _lxf2(marc_file,  sheet_name=0),
                        _mvke_df      = _lxf2(mvke_file,  sheet_name=0),
                        _zplm_marc_df = _lxf2(zplm_marc,  sheet_name=0) if zplm_marc  else None,
                        _zplm_mara_df = _lxf2(zplm_mara,  sheet_name=0) if zplm_mara  else None,
                        _marm_df      = _lxf2(marm_file,  sheet_name=0) if marm_file  else None,
                        _mbew_df      = _lxf2(mbew_file,  sheet_name=0) if mbew_file  else None,
                        _csr_item_df  = _lxf2(csr_item,   sheet_name=0) if csr_item   else None,
                        _csr_head_df  = _lxf2(csr_head,   sheet_name=0) if csr_head   else None,
                        _matplsta_df  = _lxf2(matplsta,   sheet_name=0) if matplsta   else None,
                        _mard_df      = _lxf2(mard_file,  sheet_name=0) if mard_file  else None,
                    )

                    _tn_filter = plant_filter_export if plant_filter_export in ("TN01", "TN02") else None
                    excel_tn, errors_tn, _ = run_tn_validation(
                        mara_file=None, marc_file=None, mvke_file=None,
                        previous_report=None, plant_filter=_tn_filter,
                        progress_callback=lambda p, m: update_progress(50 + int(p * 0.45), m),
                        **_shared,
                    )

                    update_progress(96, "Merging results...")

                    errors_pt_clean = errors_pt if errors_pt is not None else pd.DataFrame()
                    errors_tn_clean = errors_tn if errors_tn is not None else pd.DataFrame()

                    _error_frames = [
                        df for df in [errors_pt_clean, errors_tn_clean]
                        if df is not None and not df.empty
                    ]
                    errors_df = (
                        pd.concat(_error_frames, ignore_index=True)
                        if _error_frames
                        else pd.DataFrame(columns=['Line', 'Field', 'Error', 'Material', 'Plant'])
                    )

                    # Deduplicate: drop exact duplicates (Material+Plant+Field+Error).
                    # For plant-agnostic errors (Plant=''), if same Material+Field+Error
                    # exists with a specific plant, drop the plant-less duplicate.
                    if not errors_df.empty and 'Field' in errors_df.columns:

                        _key = ['Material', 'Field', 'Error']

                        if 'Plant' in errors_df.columns:
                            errors_df['Plant'] = errors_df['Plant'].fillna('').astype(str).str.strip()

                            # 1. remover duplicados exatos
                            errors_df = errors_df.drop_duplicates(subset=_key + ['Plant']).reset_index(drop=True)

                            # 2. separar
                            _with_plant    = errors_df[errors_df['Plant'] != '']
                            _without_plant = errors_df[errors_df['Plant'] == '']

                            # 3. dedupe global (PT+TN duplicados)
                            if not _without_plant.empty:
                                _without_plant = _without_plant.drop_duplicates(subset=_key)

                            # 4. remover global se existir versão com plant
                            if not _with_plant.empty and not _without_plant.empty:
                                _without_plant = _without_plant[
                                    ~_without_plant.set_index(_key).index.isin(
                                        _with_plant.set_index(_key).index
                                    )
                                ]

                            # 5. juntar
                            errors_df = pd.concat([_with_plant, _without_plant], ignore_index=True)

                    # Load RolesViews + combined MARC/MARA for postprocessing
                    from engines.common.fast_loader import load_xlsx_fast as _lxf3
                    from engines.material_rules.sap_like import ensure_basic_data_k
                    from engines.sap_material.postprocess import (
                        postprocess_with_views_roles_and_classification as _postprocess,
                    )
                    from engines.common.excel_report import _style_sap_report as _style
                    _rv = _lxf3(_safe_rolesviews_bytesio(), sheet_name=None)
                    _roles_df_m = _rv.get("Roles")
                    _views_df_m = _rv.get("Views")

                    # Build combined MARC — must apply ensure_basic_data_k so that
                    # Maintenance status contains 'K' for all rows, otherwise all
                    # materials appear as "Missing view K - Basic Data"
                    _marc_pt_df = _lxf3(marc_file, sheet_name=0) if marc_file else pd.DataFrame()
                    _marc_tn_df = _shared.get('_marc_df')
                    if _marc_tn_df is None:
                        _marc_tn_df = pd.DataFrame()
                    _marc_combined = pd.concat(
                        [df for df in [_marc_pt_df, _marc_tn_df] if not df.empty],
                        ignore_index=True
                    )
                    _marc_combined = ensure_basic_data_k(_marc_combined)

                    _mara_df_m = _shared.get('_mara_df')
                    if _mara_df_m is None:
                        _mara_df_m = _lxf3(mara_file, sheet_name=0)

                    # Run postprocess on merged errors — builds contagem_final_df and
                    # coverage_df with all columns (View, Role, View_Description, etc.)
                    (
                        errors_df, _, views_missing_output_df,
                        contagem_final_df, coverage_df, view_tabs_dict
                    ) = _postprocess(
                        errors_df, _mara_df_m, _marc_combined,
                        _roles_df_m, _views_df_m, plant_filter=None
                    )

                    summary_df = (
                        errors_df["Field"].value_counts().rename_axis("Field").reset_index(name="Adjusted Nº of Errors")
                        if "Field" in errors_df.columns and not errors_df.empty
                        else pd.DataFrame({"Field": [], "Adjusted Nº of Errors": []})
                    )

                    # Build merged Excel with pandas + _style_sap_report
                    import io as _io
                    _file_bytes = _io.BytesIO()

                    if not contagem_final_df.empty and 'Adjusted Nº of Errors' in contagem_final_df.columns:
                        contagem_final_df = contagem_final_df.sort_values(
                            'Adjusted Nº of Errors', ascending=False
                        ).reset_index(drop=True)

                    def _cat_missing_m(vtabs):
                        cols = ['Material', 'Plant', 'Missing_View', 'View_Description', 'Fields_Expected']
                        frames = [dfv[cols] for dfv in vtabs.values() if dfv is not None and not dfv.empty]
                        if not frames:
                            return pd.DataFrame(columns=cols)
                        return pd.concat(frames, ignore_index=True).drop_duplicates()

                    view_is_missing_df = _cat_missing_m(view_tabs_dict)

                    with pd.ExcelWriter(_file_bytes, engine='openpyxl') as _writer:
                        pd.DataFrame().to_excel(_writer, sheet_name='Executive Summary', index=False)
                        _base_row = 1
                        if not contagem_final_df.empty:
                            contagem_final_df.to_excel(_writer, sheet_name='Adjusted Summary',
                                                        index=False, startrow=1)
                            _base_row = len(contagem_final_df) + 5
                        if not coverage_df.empty:
                            coverage_df.to_excel(_writer, sheet_name='Adjusted Summary',
                                                  index=False, startrow=_base_row)
                        views_missing_output_df.to_excel(_writer, sheet_name='Field With Error', index=False)
                        errors_df.to_excel(_writer, sheet_name='All Errors', index=False)
                        view_is_missing_df.to_excel(_writer, sheet_name='View is Missing', index=False)
                        for _vc, _vdf in view_tabs_dict.items():
                            _vdf.to_excel(_writer, sheet_name=f'Missing_{_vc}'[:31], index=False)

                    _file_bytes.seek(0)
                    _file_bytes = _style(
                        _file_bytes,
                        errors_df=errors_df,
                        views_missing_output_df=views_missing_output_df,
                        contagem_final_df=contagem_final_df,
                        coverage_df=coverage_df,
                        view_tabs_dict=view_tabs_dict,
                        sap_df=None,
                        plant_filter=None,
                    )

                    excel_display = _file_bytes
                    excel_export  = _file_bytes
                    st.session_state.pop("tn_excel_export", None)

                # ── PALMELA ONLY (default) ─────────────────────────────────────
                else:
                    excel_display, errors_df, summary_df = run_full_validation(
                        mara_file=mara_file, marc_file=marc_file,
                        mvke_file=mvke_file,
                        zplm_marc_file=zplm_marc, zplm_mara_file=zplm_mara,
                        marm_file=marm_file,
                        previous_report=None,
                        plant_filter=plant_filter_display,
                        progress_callback=update_progress,
                        csr_item_file=csr_item, csr_head_file=csr_head,
                        matplsta_file=matplsta, mbew_file=mbew_file,
                        mard_file=mard_file,
                    )

                    try:
                        st.session_state["sap_view_missing_df"] = pd.read_excel(excel_display, sheet_name="View is Missing")
                    except Exception:
                        st.session_state["sap_view_missing_df"] = pd.DataFrame()

                    # Avoid re-running the engine when export scope differs from display scope
                    if plant_filter_export != plant_filter_display:
                        from engines.common.excel_report import _style_sap_report
                        from engines.sap_material.postprocess import (
                            postprocess_with_views_roles_and_classification,
                        )
                        from engines.common.fast_loader import load_xlsx_fast as _lxf
                        import io as _io
                        _rv = _lxf(_safe_rolesviews_bytesio(), sheet_name=None)
                        _roles = _rv.get("Roles"); _views = _rv.get("Views")
                        _mara_df = _lxf(mara_file, sheet_name=0)
                        _marc_df = _lxf(marc_file, sheet_name=0)
                        (_edf, _, _vmout, _cdf, _covdf, _vtabs) = \
                            postprocess_with_views_roles_and_classification(
                                errors_df, _mara_df, _marc_df, _roles, _views,
                                plant_filter=plant_filter_export)
                        _vis_df = pd.concat(
                            [df for df in _vtabs.values() if df is not None and not df.empty],
                            ignore_index=True
                        ) if _vtabs else pd.DataFrame()
                        _buf = _io.BytesIO()
                        with pd.ExcelWriter(_buf, engine="openpyxl") as _w:
                            pd.DataFrame().to_excel(_w, sheet_name="Executive Summary", index=False)
                            _cdf.to_excel(_w, sheet_name="Adjusted Summary", index=False, startrow=1)
                            _vmout.to_excel(_w, sheet_name="Field With Error", index=False)
                            _edf.to_excel(_w, sheet_name="All Errors", index=False)
                            _vis_df.to_excel(_w, sheet_name="View is Missing", index=False)
                            for _vc, _vdf in _vtabs.items():
                                _vdf.to_excel(_w, sheet_name=f"Missing_{_vc}"[:31], index=False)
                            _marc_df.to_excel(
                                _w, sheet_name="Original Data (MARC)", index=False)
                        _buf.seek(0)
                        excel_export = _style_sap_report(
                            _buf, errors_df=_edf, views_missing_output_df=_vmout,
                            contagem_final_df=_cdf, coverage_df=_covdf,
                            view_tabs_dict=_vtabs, sap_df=None,
                            plant_filter=plant_filter_export)
                    else:
                        excel_export = excel_display

            # Build suffix for filename
            _site_tag = "TN" if _is_tn else ("ALL_SITES" if _is_both else "PT")
            suffix = plant_filter_export or _site_tag
            report_path = save_report_to_vault(excel_export, report_type="sap", plant_suffix=suffix)
            st.session_state.sap_report_path  = report_path
            st.session_state.sap_errors_df    = errors_df
            st.session_state.sap_summary_df   = summary_df
            st.session_state.sap_excel_export = excel_export

            if save_history:
                payload = build_sap_log_payload(
                    user=effective_executor, plant_filter=plant_filter_export or suffix,
                    mara_file=mara_file, marc_file=marc_file,
                    mvke_file=mvke_file, errors_df=errors_df, summary_df=summary_df,
                    report_path=report_path,
                )
                save_run_log(payload)

            _pb_slot.empty()
            st.session_state.sap_step = 4
            st.rerun()

    # --- STEP 4: RESULTS ---
    elif st.session_state.sap_step == 4:
        st.markdown("### 📊 Step 4 — Results")

        errors_df    = st.session_state.sap_errors_df
        summary_df   = st.session_state.sap_summary_df
        excel_export = st.session_state.sap_excel_export
        report_path  = st.session_state.sap_report_path
        
        # 2. dashboard owners ✅ (o que adicionaste)
        _enriched = merge_assignments(errors_df)
        save_open_assignments(_enriched)
        track_error_count(_enriched)

        # FIX 4: count fields/materials consistently with Adjusted Summary
        # (Field Errors only — excludes Missing View errors)
        _fe_only = (errors_df[errors_df["Error_Type"] == "Field Error"]
                    if "Error_Type" in errors_df.columns else errors_df)
        _fe_cnt   = int(_fe_only.shape[0])
        _f_cnt    = int(_fe_only["Field"].nunique())    if "Field"    in _fe_only.columns else 0
        _m_cnt    = int(_fe_only["Material"].nunique()) if "Material" in _fe_only.columns else 0

        c = st.columns(3)
        with c[0]:
            st.markdown(f'<div class="card"><h4>Field Errors (Actionable)</h4><div class="big">{_fe_cnt:,}</div></div>', unsafe_allow_html=True)
        with c[1]:
            st.markdown(f'<div class="card"><h4>Distinct Fields</h4><div class="big">{_f_cnt}</div></div>', unsafe_allow_html=True)
        with c[2]:
            st.markdown(f'<div class="card"><h4>Materials with Errors</h4><div class="big">{_m_cnt:,}</div></div>', unsafe_allow_html=True)

        # Auto-insight — a plain-language summary of the results
        _top_field = _fe_only['Field'].value_counts().index[0] if (not _fe_only.empty and 'Field' in _fe_only.columns) else "—"
        _top_plant = (
            _fe_only['Plant'].astype(str).str.strip().replace('', 'General')
            .value_counts().index[0]
            if (not _fe_only.empty and 'Plant' in _fe_only.columns) else "—"
        )
        _mv_count = int((errors_df['Error_Type'] == 'Missing View').sum()) if 'Error_Type' in errors_df.columns else 0
        _insight_color = "var(--app-bg-secondary)" if _fe_cnt > 1000 else ("var(--app-bg-secondary)" if _fe_cnt > 200 else "var(--app-bg-secondary)")
        _insight_border = "var(--color-error)" if _fe_cnt > 1000 else ("var(--color-warn)" if _fe_cnt > 200 else "var(--color-ok)")
        st.markdown(f"""
        <div style="background:{_insight_color};border-left:4px solid {_insight_border};border-radius:8px;
                    padding:12px 16px;margin-bottom:16px;font-size:0.9rem;line-height:1.6;color:var(--app-fg)">
            <strong>📋 Validation Summary</strong><br>
            Found <strong>{_fe_cnt:,} actionable errors</strong> across <strong>{_m_cnt:,} materials</strong>
            in <strong>{_f_cnt} distinct fields</strong>.
            The most affected field is <strong>{_top_field}</strong> and the most affected plant is <strong>{_top_plant}</strong>.
            {"Additionally, " + str(_mv_count) + " errors were suppressed because the required SAP View is not yet active." if _mv_count else ""}
        </div>
        """, unsafe_allow_html=True)

        # FIX 7: Charts first, then tables in expanders
        st.markdown("## 📈 Dashboard — SAP Material Validation")
        render_sap_run_kpis(errors_df)
        render_sap_run_charts(errors_df)

        owners_df = build_owner_dashboard_df(_enriched)
        
        st.markdown("## 👥 Owner Dashboard")

        if owners_df.empty:
            st.info("No owners yet")
        else:
            st.dataframe(owners_df, width="stretch")


        download_name = os.path.basename(report_path) if report_path else "SAP_Validation_Adjusted_Report_ALL.xlsx"
        st.download_button("⬇️ Download Excel", data=excel_export.getvalue(),
                           file_name=download_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # ── Assignment panel ──────────────────────────────────────────────
        render_assignment_panel(errors_df)

        with st.expander("📋 All Errors", expanded=False):
            # Quick search filter
            search = st.text_input("🔍 Filter by Material, Field or Error", key="sap_err_search")
            df_show = _coerce_df_for_display(errors_df)
            if search.strip():
                mask = df_show.apply(
                    lambda col: col.astype(str).str.contains(search.strip(), case=False, na=False)
                ).any(axis=1)
                df_show = df_show[mask]
            st.caption(f"Showing {len(df_show):,} of {len(errors_df):,} rows")
            st.dataframe(df_show, width='stretch')

        with st.expander("📋 Adjusted Summary", expanded=False):
            st.dataframe(_coerce_df_for_display(summary_df), width='stretch')

        render_history_with_download(
            "sap", limit=10, label="SAP History",
            col_map={"When": "timestamp", "Who": "user",
                     "Plant": "plant_filter", "Errors": "metrics.total_errors"}
        )

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

    render_footer(start_time, version="v1.4")
