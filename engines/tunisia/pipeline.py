"""Pipeline: Tunisia validation (reutiliza as regras de Palmela + exceções TN)."""

import io
import pandas as pd

from engines.assignments.store import (
    build_assignments_sheet,
    merge_assignments,
    save_open_assignments,
)
from engines.common.excel_report import _style_sap_report
from engines.common.fast_loader import load_xlsx_fast
from engines.common.helpers import _coerce_df_for_display, _normalize_df_columns
from engines.common.rolesviews import get_rolesviews_bytesio
from engines.material_rules.batch_utils import (
    apply_display_field_renames,
    consolidar_batch_management_errors,
)
from engines.material_rules.cross_file import (
    validar_consumption_per,
    validar_costing_lot_size,
    validar_country_region_origin_fp,
    validar_enterproj_number,
    validar_item_category_group,
    validar_lifo_fifo_relevant_mbew,
    validar_marm_pc_ea_roh_halb,
    validar_material_price_determination_control,
    validar_packaging_material_type_mara,
    validar_product_hierarchy_vs_enterproj,
    validar_safety_time_marc,
    validar_stock_determination_group_marc,
    validar_version_indicator_marc,
)
from engines.material_rules.enrichment import (
    add_old_material_number_errors,
    fill_plant_and_validate_enterproj_from_zplm,
    validar_neg_stocks_in_plant,
)
from engines.material_rules.fields import (
    validar_batch_management,
    validar_batch_management_mara_global,
    validar_campo_obrigatorio,
    validar_division,
    validar_export_import_group,
    validar_general_item_category_group,
    validar_gross_weight,
    validar_loading_group,
    validar_lot_sizing_procedure,
    validar_material_group,
    validar_material_group_appoint,
    validar_mrp_controller,
    validar_net_weight,
    validar_per_lot_sz_plng_cal,
    validar_period_type,
    validar_plant_specific_material_status,
    validar_price_control,
    validar_price_unit,
    validar_product_hierarchy,
    validar_product_hierarchy_mvke,
    validar_product_related_origin,
    validar_purchasing_group_dxx,
    validar_valuation_class,
    validar_with_qty_structure,
)
from engines.material_rules.sap_like import (
    _exclude_material_plant_status_99,
    _filter_valid_material_types,
    _get_mps99_material_plant_pairs,
    build_sap_like_df,
    ensure_basic_data_k,
)
from engines.sap_material.postprocess import postprocess_with_views_roles_and_classification
from engines.tunisia.rules import (
    TN_PLANTS,
    tn_validar_account_assignment_group,
    tn_validar_advanced_planning_fields,
    tn_validar_automatic_purchase_order,
    tn_validar_availability_check,
    tn_validar_backflush,
    tn_validar_consumption_mode,
    tn_validar_gr_processing_time,
    tn_validar_individual_collective_requirements,
    tn_validar_jit_delivery_sched,
    tn_validar_mrp_group,
    tn_validar_mrp_type,
    tn_validar_planning_calendar,
    tn_validar_planning_time_fence,
    tn_validar_procurement_type,
    tn_validar_production_scheduling_profile,
    tn_validar_production_supervisor,
    tn_validar_profit_center,
    tn_validar_purchasing_group,
    tn_validar_repetitive_manufacturing,
    tn_validar_repetitive_manufacturing_profile,
    tn_validar_source_list_required,
    tn_validar_special_procurement_type,
    tn_validar_storage_location,
    tn_validar_storage_location_ext_procurement,
    tn_validar_strategy_group,
)

def run_tn_validation(
    mara_file,
    marc_file,
    mvke_file,
    marm_file=None,
    mbew_file=None,
    zplm_marc_file=None,
    zplm_mara_file=None,
    previous_report=None,
    plant_filter=None,
    progress_callback=None,
    csr_item_file=None,
    csr_head_file=None,
    matplsta_file=None,
    mard_file=None,
    # Optional pre-loaded DataFrames — if provided, skips file loading for those
    _mara_df=None,
    _marc_df=None,
    _mvke_df=None,
    _marm_df=None,
    _mbew_df=None,
    _zplm_marc_df=None,
    _zplm_mara_df=None,
    _csr_item_df=None,
    _csr_head_df=None,
    _matplsta_df=None,
    _mard_df=None,
):
    """
    Runs the complete Tunisia validation pipeline.
    Returns: (excel_bytes, errors_df, summary_df)
    """
    def _cb(pct, msg):
        if progress_callback:
            progress_callback(min(pct, 100), msg)

    # ── 1. Load files (skip if pre-loaded DataFrames provided) ───────────────
    _cb(3,  "Loading MARA...")
    mara_df = _mara_df if _mara_df is not None else load_xlsx_fast(mara_file, sheet_name=0)
    _cb(6,  "Loading MARC...")
    marc_df = _marc_df if _marc_df is not None else load_xlsx_fast(marc_file, sheet_name=0)
    _cb(9,  "Loading MVKE...")
    mvke_df = _mvke_df if _mvke_df is not None else load_xlsx_fast(mvke_file, sheet_name=0)
    _cb(11, "Loading ZPLM MARC...")
    zplm_marc_df = _zplm_marc_df if _zplm_marc_df is not None else (load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None)
    _cb(13, "Loading ZPLM MARA...")
    zplm_mara_df = _zplm_mara_df if _zplm_mara_df is not None else (load_xlsx_fast(zplm_mara_file, sheet_name=0) if zplm_mara_file else None)
    _cb(15, "Loading optional files...")
    csr_item_df = _csr_item_df if _csr_item_df is not None else (load_xlsx_fast(csr_item_file, sheet_name=0) if csr_item_file else None)
    csr_head_df = _csr_head_df if _csr_head_df is not None else (load_xlsx_fast(csr_head_file, sheet_name=0) if csr_head_file else None)
    matplsta_df = _matplsta_df if _matplsta_df is not None else (load_xlsx_fast(matplsta_file, sheet_name=0) if matplsta_file else None)
    marm_df     = _marm_df if _marm_df is not None else (load_xlsx_fast(marm_file, sheet_name=0) if marm_file else None)
    mbew_df     = _mbew_df if _mbew_df is not None else (load_xlsx_fast(mbew_file, sheet_name=0) if mbew_file else None)
    mard_df     = _mard_df if _mard_df is not None else (load_xlsx_fast(mard_file, sheet_name=0) if mard_file else None)

    if mara_df is None:
        raise ValueError("MARA.xlsx not loaded.")
    if marc_df is None:
        raise ValueError("MARC.xlsx not loaded.")
    if mvke_df is None:
        raise ValueError("MVKE.xlsx not loaded.")

    marc_df = _normalize_df_columns(marc_df, {
        'plant': 'Plant',
        'werks': 'Plant',
        'intrastat group': 'Export/Import Group',
        'period indicator': 'Period Type',
        'material plant type': 'Plant Type',
        'purchasing group': 'Purchasing Group',
        'source list': 'Source list',
        'source list required': 'Source list required',
        'automatic po': 'Automatic Purchase Order Allowed',
        'automatic purchase order': 'Automatic Purchase Order Allowed',
        'issue storage location': 'Issue Storage Location',
        'production storage location': 'Issue Storage Location',
        'planning strategy group': 'Planning Strategy Group',
        'mrp group': 'MRP Group',
        'disgr': 'MRP Group',
        'plant-specific material status': 'Plant Specific Material Status',
    })
    mbew_df = _normalize_df_columns(mbew_df, {
        'plant': 'Plant',
        'werks': 'Plant',
        'valuation area': 'Plant',
        'bwkey': 'Plant',
        'price control': 'Price Control',
        'price unit': 'Price Unit',
        'valuation class': 'Valuation Class',
        'cost est. with qty structure': 'With Qty Structure',
    })
    # ✅ 'Valuation Area' alias — validar_lifo_fifo_relevant_mbew() hard-requires
    # a column literally named 'Valuation Area' (MBEW's own SAP field name),
    # while everything else here was normalized to 'Plant'. Keep both.
    if mbew_df is not None and 'Plant' in mbew_df.columns:
        mbew_df['Valuation Area'] = mbew_df['Plant']

    zplm_marc_df = _normalize_df_columns(zplm_marc_df, {
        'plant': 'Plant',
        'werks': 'Plant',
        'material plant type': 'Plant Type',
        'enterproj number': 'Enterproj Number',
        'material transfer': 'Material Transfer',
        'eop date': 'EOP Date',
        'service liability date': 'Service Liability Date',
        'ltb quantity': 'LTB Quantity',
        'ltb start date': 'LTB Start Date',
    })

    # ✅ DEDUPLICATE ZPLM MARC to one row per (Material, Plant) — BEFORE any
    # consumer touches it. See sap_validation_engine.py's run_full_validation
    # for the full explanation: ZPLM MARC's change-history columns suggest
    # it can hold multiple historical rows per (Material, Plant), which
    # would otherwise multiply every validation that merges against it
    # directly (validar_batch_management_plant_marc, validar_version_indicator_marc,
    # validar_country_region_origin_fp, tn_validar_storage_location_ext_procurement).
    if zplm_marc_df is not None and not zplm_marc_df.empty and {'Material', 'Plant'}.issubset(zplm_marc_df.columns):
        _sort_col = next(
            (c for c in ['Changed On', 'Change Number', 'Created On'] if c in zplm_marc_df.columns),
            None
        )
        if _sort_col:
            # Convert to string to handle mixed datetime/time types
            zplm_marc_df[_sort_col] = zplm_marc_df[_sort_col].fillna('').astype(str)
            zplm_marc_df = zplm_marc_df.sort_values(_sort_col, ascending=False, na_position='last')
        zplm_marc_df = zplm_marc_df.drop_duplicates(subset=['Material', 'Plant'], keep='first')

    mvke_df = _normalize_df_columns(mvke_df, {'plant': 'Plant', 'werks': 'Plant', 'material': 'Material'})
    marm_df = _normalize_df_columns(marm_df, {'plant': 'Plant', 'werks': 'Plant', 'material': 'Material'})
    mard_df = _normalize_df_columns(mard_df, {'plant': 'Plant', 'werks': 'Plant', 'material': 'Material'})

    rolesviews_dict = load_xlsx_fast(get_rolesviews_bytesio(), sheet_name=None)
    if "Roles" not in rolesviews_dict or "Views" not in rolesviews_dict:
        raise ValueError("RolesViews.xlsx must contain sheets 'Roles' and 'Views'.")
    roles_df = rolesviews_dict["Roles"]
    views_df = rolesviews_dict["Views"]

    # ── 2. Filter to TN plants ────────────────────────────────────────────────
    tn_scope = set(TN_PLANTS)
    if plant_filter:
        pf = str(plant_filter).strip().upper()
        if pf in TN_PLANTS:
            tn_scope = {pf}

    def _ftn(df, col='Plant'):
        if df is None or df.empty or col not in df.columns:
            return df
        return df[df[col].astype(str).str.strip().str.upper().isin(tn_scope)].copy()

    marc_tn      = _ftn(marc_df)
    zplm_marc_tn = _ftn(zplm_marc_df) if zplm_marc_df is not None else None
    mard_tn      = _ftn(mard_df) if mard_df is not None else None
    marc_tn      = ensure_basic_data_k(marc_tn)

    tn_materials = set(
        marc_tn['Material'].astype(str).str.strip()
        .loc[marc_tn['Material'].notna()]
    )

    def _filter_to_tn_materials(df):
        if df is None or df.empty or 'Material' not in df.columns:
            return df

        out = df[df['Material'].astype(str).str.strip().isin(tn_materials)].copy()

        plant_col = next(
            (c for c in ['Plant', 'Valuation Area', 'BWKEY', 'WERKS']
             if c in out.columns),
            None
        )
        if plant_col:
            out = out[
                out[plant_col].astype(str).str.strip().str.upper().isin(tn_scope)
            ].copy()

        return out

    # mara_df is material-level (no Plant column) — do NOT filter by tn_materials.
    # Filtering it would exclude materials that exist in MARA but have no TN marc
    # entry yet, causing MARA-based validations to silently miss them.
    # Plant-level tables are filtered normally.
    #
    # marm_df (Alternative UoM) is ALSO material-level, not plant-level, and its
    # only consumer is validar_marm_pc_ea_roh_halb, which is called with the full
    # (unfiltered) mara_df. Filtering marm_df down to TN-only materials here would
    # create a mismatch: materials that legitimately exist only in Palmela plants
    # would still appear in mara_relevant (built from the global mara_df) but be
    # missing from marm_groups (built from the TN-restricted marm_df), producing
    # false "Alternative UoM found: none" errors for every non-TN material when
    # running the combined/general report. So marm_df is intentionally left
    # unfiltered here, same as mara_df.
    mbew_df      = _filter_to_tn_materials(mbew_df)
    mvke_df      = _filter_to_tn_materials(mvke_df)
    # zplm_mara_df: kept UNFILTERED for the same reason as marm_df above.
    # validar_product_hierarchy_vs_enterproj() iterates the full (unfiltered)
    # mara_df looking up each FERT material's Enterproj Number in zplm_mara_df.
    # If zplm_mara_df were restricted to TN materials, every Palmela-only FERT
    # material would be missing from that lookup, be treated as Enterproj="000...",
    # and be falsely flagged as a "PH mismatch" when running the combined report.

    # ── 3. Build sap_like_df ──────────────────────────────────────────────────
    _cb(18, "Building SAP-like dataframe from MARC + MARA + ZPLM MARC...")
    sap_like_df = build_sap_like_df(marc_tn, mara_df, zplm_marc_tn)

    # ✅ GLOBAL FILTER — ONLY FERT / HALB / ROH
    sap_like_df = _filter_valid_material_types(sap_like_df, mara_df)
    advanced_planning_df = sap_like_df.copy()

    # aplicar também às tabelas usadas em validações diretas
    marc_tn = _filter_valid_material_types(marc_tn, mara_df)
    marc_tn_full = marc_tn.copy()

    exclude_mps99_pairs = _get_mps99_material_plant_pairs(marc_tn_full)
    if exclude_mps99_pairs:
        sap_like_df = _exclude_material_plant_status_99(sap_like_df, exclude_mps99_pairs)
        marc_tn = _exclude_material_plant_status_99(marc_tn, exclude_mps99_pairs)
        if zplm_marc_tn is not None:
            zplm_marc_tn = _exclude_material_plant_status_99(
                zplm_marc_tn,
                exclude_mps99_pairs,
            )
        if mvke_df is not None:
            mvke_df = _exclude_material_plant_status_99(
                mvke_df,
                exclude_mps99_pairs,
            )
        if mbew_df is not None:
            if 'Plant' in mbew_df.columns:
                mbew_df = _exclude_material_plant_status_99(
                    mbew_df,
                    exclude_mps99_pairs,
                    plant_col='Plant',
                )
            elif 'Valuation Area' in mbew_df.columns:
                mbew_df = _exclude_material_plant_status_99(
                    mbew_df,
                    exclude_mps99_pairs,
                    plant_col='Valuation Area',
                )

    if mvke_df is not None:
        mvke_df = _filter_valid_material_types(mvke_df, mara_df)

    if mbew_df is not None:
        mbew_df = _filter_valid_material_types(mbew_df, mara_df)

    if mbew_df is not None and not mbew_df.empty:
        mbew_cols = [c for c in [
            'Material', 'Plant', 'Price Control', 'Price Unit',
            'Valuation Class', 'With Qty Structure'
        ] if c in mbew_df.columns]
        if 'Material' in mbew_cols:
            mbew_small = mbew_df[mbew_cols].copy()
            if 'Plant' in mbew_small.columns and 'Plant' in sap_like_df.columns:
                mbew_small = mbew_small.drop_duplicates(subset=['Material', 'Plant'], keep='first')
                sap_like_df = sap_like_df.merge(mbew_small, on=['Material', 'Plant'], how='left')
            else:
                mbew_small = mbew_small.drop_duplicates(subset=['Material'], keep='first')
                sap_like_df = sap_like_df.merge(mbew_small, on='Material', how='left')

    resultados = []

    # ── 4A. Mandatory fields ──────────────────────────────────────────────────
    _cb(20, "Validating mandatory fields...")
    for col in ['Material', 'Material Type', 'Description', 'Base UOM',
                'Plant', 'Plant Type']:
        validar_campo_obrigatorio(sap_like_df, col, resultados)

    # ── 4B. Fields WITHOUT TN exceptions — Palmela functions directly ─────────
    _cb(22, "Validating Division, Export/Import Group, Gen. Item Category...")
    validar_division(sap_like_df, resultados)
    validar_export_import_group(sap_like_df, resultados)
    validar_general_item_category_group(sap_like_df, resultados)
    validar_batch_management(sap_like_df, resultados)  # uses Batch Management Required Ind. from MARA
    validar_batch_management_mara_global(sap_like_df, mara_df, resultados)

    _cb(25, "Validating Gross/Net Weight...")
    validar_gross_weight(sap_like_df, resultados)
    validar_net_weight(sap_like_df, resultados)

    _cb(26, "Validating Loading Group, Lot Sizing Procedure...")
    validar_loading_group(sap_like_df, resultados)
    validar_lot_sizing_procedure(sap_like_df, resultados)

    _cb(27, "Validating Material Group...")
    validar_material_group(sap_like_df, resultados)
    validar_material_group_appoint(sap_like_df, resultados)

    _cb(28, "Validating Period Type...")
    validar_period_type(sap_like_df, resultados)

    _cb(29, "Validating Price Control, Price Unit...")
    # Only run MBEW-derived checks when there are rows in the selected TN scope.
    # This avoids false "column not found" warnings when All Sites is selected
    # but the uploaded files contain no TN01/TN02 records.
    if sap_like_df is not None and not sap_like_df.empty:
        validar_price_control(sap_like_df, resultados)
        validar_price_unit(sap_like_df, resultados)

    _cb(30, "Validating Product Hierarchy...")
    validar_product_hierarchy(sap_like_df, resultados, mvke_df, mara_df)
    validar_product_hierarchy_mvke(sap_like_df, resultados, mvke_df)

    _cb(31, "Validating Valuation Class, With Qty Structure...")
    if sap_like_df is not None and not sap_like_df.empty:
        validar_valuation_class(sap_like_df, resultados)
        validar_with_qty_structure(sap_like_df, resultados)

    _cb(32, "Validating Product-Related Origin, MRP Controller, PSMS...")
    validar_product_related_origin(sap_like_df, resultados)
    validar_mrp_controller(sap_like_df, resultados, mard_tn)
    validar_plant_specific_material_status(sap_like_df, resultados)

    _cb(33, "Validating Enterproj Number, Purchasing Group DXX...")
    validar_enterproj_number(sap_like_df, resultados)
    validar_purchasing_group_dxx(sap_like_df, resultados)

    _cb(34, "Validating Item Category Group (MVKE)...")
    validar_item_category_group(sap_like_df, resultados, mvke_df)

    # ── 4C. Fields WITH TN exceptions — TN functions ─────────────────────────
    _cb(40, "Validating Backflush (TN)...")
    tn_validar_backflush(sap_like_df, resultados)

    _cb(41, "Validating Strategy Group (TN)...")
    tn_validar_strategy_group(sap_like_df, mara_df, resultados)

    _cb(42, "Validating Availability Check (TN)...")
    tn_validar_availability_check(sap_like_df, resultados)

    _cb(43, "Validating Procurement Type (TN)...")
    tn_validar_procurement_type(sap_like_df, mara_df, resultados)

    _cb(44, "Validating Automatic Purchase Order (TN)...")
    tn_validar_automatic_purchase_order(sap_like_df, resultados)

    _cb(45, "Validating Consumption Mode (TN)...")
    tn_validar_consumption_mode(sap_like_df, mara_df, resultados)

    _cb(46, "Validating JIT Delivery Sched. (TN)...")
    tn_validar_jit_delivery_sched(sap_like_df, resultados)

    _cb(47, "Validating MRP Type (TN)...")
    tn_validar_mrp_type(sap_like_df, mara_df, resultados, zplm_marc_tn,
                         csr_item_df, csr_head_df)

    _cb(47.5, "Validating MRP Group (TN02 FERT)...")
    tn_validar_mrp_group(marc_tn, mara_df, resultados)

    _cb(48, "Validating Planning Calendar (TN)...")
    tn_validar_planning_calendar(sap_like_df, resultados)

    _cb(49, "Validating Planning Time Fence (TN)...")
    tn_validar_planning_time_fence(sap_like_df, mara_df, resultados)

    _cb(50, "Validating Production Scheduling Profile (TN)...")
    tn_validar_production_scheduling_profile(sap_like_df, mara_df, resultados)

    _cb(51, "Validating Profit Center (TN)...")
    tn_validar_profit_center(sap_like_df, resultados)

    _cb(52, "Validating Purchasing Group (TN)...")
    tn_validar_purchasing_group(sap_like_df, resultados)

    _cb(53, "Validating Repetitive Manufacturing (TN)...")
    tn_validar_repetitive_manufacturing(sap_like_df, mara_df, resultados)

    _cb(54, "Validating Repetitive Manufacturing Profile (TN)...")
    tn_validar_repetitive_manufacturing_profile(sap_like_df, mara_df, resultados)

    _cb(55, "Validating Source List Required (TN)...")
    tn_validar_source_list_required(sap_like_df, resultados)

    _cb(56, "Validating Issue Storage Location (TN)...")
    tn_validar_storage_location(sap_like_df, mara_df, resultados)

    # ── 4D. MARC-based cross-file validations (all from Palmela) ─────────────
    _cb(57, "Validating Bwd/Fwd Consumption Per. (MARC)...")
    validar_consumption_per(marc_tn, mara_df, resultados)

    _cb(58, "Validating Per.Lot Sz.Plng Cal. (MARC)...")
    if marc_tn is not None and not marc_tn.empty:
        validar_per_lot_sz_plng_cal(advanced_planning_df, resultados, mara_df, matplsta_df)

    _cb(59, "Validating Stock Determination Group (MARC)...")
    validar_stock_determination_group_marc(marc_tn, mara_df, resultados)

    _cb(60, "Validating Packaging Mat. Type (MARA)...")
    validar_packaging_material_type_mara(mara_df, resultados)

    # Batch Management per-plant check removed to avoid duplicate counting
    # (Batch Management Required Ind. is validated earlier via validar_batch_management)

    _cb(62, "Validating Version Indicator (MARC)...")
    if zplm_marc_tn is not None:
        validar_version_indicator_marc(marc_tn, zplm_marc_tn, mara_df, resultados)

    _cb(63, "Validating Country/Region of Origin (MARC)...")
    if zplm_marc_tn is not None:
        validar_country_region_origin_fp(marc_tn, zplm_marc_tn, mara_df, resultados)

    _cb(64, "Validating Product Hierarchy vs Enterproj (MARA)...")
    if zplm_mara_df is not None:
        validar_product_hierarchy_vs_enterproj(mara_df, zplm_mara_df, resultados)

    _cb(65, "Validating Alternative UoM PC (MARM)...")
    if marm_df is not None:
        validar_marm_pc_ea_roh_halb(marm_df=marm_df, mara_df=mara_df, resultados=resultados)

    _cb(66, "Validating LIFO/FIFO-relevant (MBEW)...")
    if mbew_df is not None:
        validar_lifo_fifo_relevant_mbew(mbew_df, mara_df, resultados)

    # ── 4E. TN-specific MARC cross-file rules ─────────────────────────────────
    _cb(67, "Validating Special Procurement Type (TN02 FP)...")
    tn_validar_special_procurement_type(marc_tn, zplm_marc_tn, resultados)

    _cb(68, "Validating Storage Location Ext. Procurement (TN)...")
    tn_validar_storage_location_ext_procurement(
        marc_tn, zplm_marc_tn, mara_df, resultados
    )

    _cb(69, "Validating Individual/Collective Requirements (TN)...")
    tn_validar_individual_collective_requirements(marc_tn, mara_df, resultados)

    _cb(70, "Validating Production Supervisor (TN)...")
    tn_validar_production_supervisor(marc_tn, mara_df, resultados)

    _cb(71, "Validating Account Assignment Group (TN MVKE)...")
    tn_validar_account_assignment_group(mara_df, marc_tn, mvke_df, resultados)



    _cb(73, "Validating Costing Lot Size (MARC)...")
    validar_costing_lot_size(marc_tn, mara_df, resultados)

    _cb(74, "Validating Safety Time (MARC)...")
    validar_safety_time_marc(marc_tn, mara_df, resultados)

    _cb(74.2, "Validating GR processing time (TN02 MARC)...")
    tn_validar_gr_processing_time(marc_tn, resultados)

    _cb(74.4, "Validating Material Price Determination Control (MBEW)...")
    if mbew_df is not None:
        validar_material_price_determination_control(mbew_df, mara_df, resultados)

    _cb(75, "Validating Advanced Planning fields (TN)...")
    tn_validar_advanced_planning_fields(advanced_planning_df, mara_df, resultados, matplsta_df)

    _cb(76, "Validating Neg. stocks in plant (MARC)...")
    neg_errors = validar_neg_stocks_in_plant(marc_tn, mara_df)
    if neg_errors:
        resultados.extend(neg_errors)

    # Batch Management: plant error in all plants + global error -> only global
    consolidar_batch_management_errors(resultados, sap_like_df)

    # ── 5. Build errors DataFrame ─────────────────────────────────────────────
    _cb(78, "Building errors DataFrame...")
    errors_df = _coerce_df_for_display(pd.DataFrame(resultados))

    # Keep a stable schema even when the selected site has zero errors/rows.
    # Without this, pd.DataFrame([]) has no columns and the Roles check below
    # raises KeyError: 'Field'.
    _error_columns = ['Line', 'Field', 'Error', 'Material', 'Plant']
    for _col in _error_columns:
        if _col not in errors_df.columns:
            errors_df[_col] = pd.Series(dtype='object')

    # ✅ Safety-net dedup: remove exact duplicates (same Material, Plant,
    # Field and Error). This is acceptable/expected to differ across
    # different plants for the same material (e.g. PT01 vs TN01), but the
    # exact same (Material, Plant, Field, Error) combination should never
    # be reported more than once — it would otherwise inflate error counts
    # if any upstream merge or source table happens to contain duplicate
    # rows for the same key.
    _dedup_key = [c for c in ['Material', 'Plant', 'Field', 'Error'] if c in errors_df.columns]
    if len(_dedup_key) == 4 and not errors_df.empty:
        for _c in _dedup_key:
            errors_df[_c] = errors_df[_c].fillna('').astype(str).str.strip()
        errors_df = errors_df.drop_duplicates(subset=_dedup_key).reset_index(drop=True)

    
    # ------------------------------------------------------------
    # ✅ DETETAR ERROS NÃO MAPEADOS NO ROLES
    # ------------------------------------------------------------
    roles_bytes = get_rolesviews_bytesio()
    roles_dict = load_xlsx_fast(roles_bytes, sheet_name=None)
    roles_df = roles_dict.get("Roles", None)

    if (roles_df is not None and not roles_df.empty
            and not errors_df.empty and 'Field' in errors_df.columns
            and 'Field' in roles_df.columns):
        errors_df['Field_clean'] = errors_df['Field'].astype(str).str.strip()
        roles_set = set(roles_df['Field'].astype(str).str.strip())

        missing_role_df = errors_df[~errors_df['Field_clean'].isin(roles_set)].copy()

        # DEBUG terminal
        if not missing_role_df.empty:
            print("\n--- ⚠️ ERROS NÃO MAPEADOS NO ROLES ---\n")
            for f in sorted(missing_role_df['Field_clean'].unique()):
                print(f)
            print("\n--------------------------------------\n")

        errors_df = errors_df.drop(columns=['Field_clean'], errors='ignore')
    else:
        missing_role_df = pd.DataFrame()

    if previous_report and zplm_marc_file and zplm_mara_file:
        errors_df = fill_plant_and_validate_enterproj_from_zplm(
            errors_df, latest_report=previous_report,
            zplm_marc_file=zplm_marc_file, zplm_mara_file=zplm_mara_file
        )

    if not errors_df.empty:
        errors_df = add_old_material_number_errors(mara_df, errors_df)

    # ── 6. Post-processing ────────────────────────────────────────────────────
    _cb(82, "Post-processing (Views / Roles / Classification)...")
    pf_post = plant_filter if (plant_filter and str(plant_filter).upper() in TN_PLANTS) else None

    (
        errors_df, errors_filtered_df, views_missing_output_df,
        contagem_final_df, coverage_df, view_tabs_dict
    ) = postprocess_with_views_roles_and_classification(
        errors_df, mara_df, marc_tn, roles_df, views_df, plant_filter=pf_post
    )

    # ── 7. Write Excel ────────────────────────────────────────────────────────
    _cb(90, "Generating Excel report...")
    file_bytes = io.BytesIO()

    if not contagem_final_df.empty and 'Adjusted Nº of Errors' in contagem_final_df.columns:
        contagem_final_df = contagem_final_df.sort_values(
            'Adjusted Nº of Errors', ascending=False).reset_index(drop=True)

    def _cat_missing(vtabs):
        cols = ['Material', 'Plant', 'Missing_View', 'View_Description', 'Fields_Expected']
        frames = [dfv[cols] for dfv in vtabs.values() if dfv is not None and not dfv.empty]
        if not frames:
            return pd.DataFrame(columns=cols)
        out = pd.concat(frames, ignore_index=True).drop_duplicates()
        out.sort_values(['Material', 'Plant', 'Missing_View'], inplace=True)
        return out

    view_is_missing_df = _cat_missing(view_tabs_dict)

    # Final guard: field names exactly as displayed in the report
    errors_df               = apply_display_field_renames(errors_df)
    views_missing_output_df = apply_display_field_renames(views_missing_output_df)
    contagem_final_df       = apply_display_field_renames(contagem_final_df)

    with pd.ExcelWriter(file_bytes, engine='openpyxl') as writer:

        pd.DataFrame().to_excel(writer, sheet_name='Executive Summary', index=False)

        base_row = 1
        if not contagem_final_df.empty:
            contagem_final_df.to_excel(
                writer,
                sheet_name='Adjusted Summary',
                index=False,
                startrow=1
            )
            base_row = len(contagem_final_df) + 5

        if not coverage_df.empty:
            coverage_df.to_excel(
                writer,
                sheet_name='Adjusted Summary',
                index=False,
                startrow=base_row
            )

        views_missing_output_df.to_excel(writer, sheet_name='Field With Error', index=False)
        errors_df.to_excel(writer, sheet_name='All Errors', index=False)
        view_is_missing_df.to_excel(writer, sheet_name='View is Missing', index=False)

        # ------------------------------------------------------------
        # ✅ NOVA SHEET → ERROS NÃO MAPEADOS
        # ------------------------------------------------------------
        if 'missing_role_df' in locals() and not missing_role_df.empty:
            missing_role_df[['Field', 'Material', 'Plant', 'Error']].to_excel(
                writer,
                sheet_name='Missing Role Mapping',
                index=False
            )

        # ------------------------------------------------------------
        # ✅ VIEWS MISSING (já existente)
        # ------------------------------------------------------------
        for view_code, df_view in view_tabs_dict.items():
            df_view.to_excel(writer, sheet_name=f'Missing_{view_code}'[:31], index=False)

        # ------------------------------------------------------------
        # ✅ Assignments sheet — enriched with owner/status from master.json
        # ------------------------------------------------------------
        try:
            _enriched = merge_assignments(errors_df)
            _asgn_df  = build_assignments_sheet(_enriched)
            _asgn_df.to_excel(writer, sheet_name='Assignments', index=False)
            save_open_assignments(_enriched)
        except Exception:
            pass

        marc_tn.to_excel(writer, sheet_name='Original Data (MARC)', index=False)

    file_bytes.seek(0)
    file_bytes = _style_sap_report(
        file_bytes,
        errors_df=errors_df,
        views_missing_output_df=views_missing_output_df,
        contagem_final_df=contagem_final_df,
        coverage_df=coverage_df,
        view_tabs_dict=view_tabs_dict,
        sap_df=marc_tn,
        plant_filter=plant_filter,
    )

    _cb(100, "Done!")

    summary_df = (
        errors_df['Field'].value_counts().rename_axis('Field').reset_index(name='Adjusted Nº of Errors')
        if 'Field' in errors_df.columns and not errors_df.empty
        else pd.DataFrame({'Field': [], 'Adjusted Nº of Errors': []})
    )

    return file_bytes, errors_df, summary_df
