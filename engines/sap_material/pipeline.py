"""Pipeline principal: SAP Material Validation (Palmela)."""

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
from engines.material_rules.advanced_planning import validar_advanced_planning_fields
from engines.material_rules.batch_utils import (
    apply_display_field_renames,
    consolidar_batch_management_errors,
)
from engines.material_rules.cross_file import (
    validar_account_assignment_group,
    validar_commodity_code_fp,
    validar_consumption_per,
    validar_country_region_origin_fp,
    validar_enterproj_number,
    validar_individual_collective_requirements_marc,
    validar_item_category_group,
    validar_lifo_fifo_relevant_mbew,
    validar_maintenance_status_c_roh_mara,
    validar_marm_pc_ea_roh_halb,
    validar_packaging_material_type_mara,
    validar_product_hierarchy_vs_enterproj,
    validar_production_supervisor_marc,
    validar_reference_matl_for_packing_mara,
    validar_special_procurement_type_pt02_transfer,
    validar_stock_determination_group_marc,
    validar_storage_location_ext_procurement,
    validar_version_indicator_marc,
    validar_wh_material_group_mara,
)
from engines.material_rules.enrichment import (
    add_old_material_number_errors,
    fill_plant_and_validate_enterproj_from_zplm,
    validar_neg_stocks_in_plant,
)
from engines.material_rules.fields import (
    validar_automatic_purchase_order,
    validar_availability_check,
    validar_backflush,
    validar_batch_management,
    validar_batch_management_mara_global,
    validar_campo_obrigatorio,
    validar_consumption_mode,
    validar_division,
    validar_export_import_group,
    validar_general_item_category_group,
    validar_gross_weight,
    validar_jit_delivery_sched,
    validar_loading_group,
    validar_lot_sizing_procedure,
    validar_material_group,
    validar_material_group_appoint,
    validar_mrp_controller,
    validar_mrp_type,
    validar_net_weight,
    validar_per_lot_sz_plng_cal,
    validar_period_type,
    validar_planning_calendar,
    validar_planning_time_fence,
    validar_plant_specific_material_status,
    validar_price_control,
    validar_price_unit,
    validar_procurement_type,
    validar_product_hierarchy,
    validar_product_hierarchy_mvke,
    validar_product_related_origin,
    validar_production_scheduling_profile,
    validar_profit_center,
    validar_purchasing_group,
    validar_purchasing_group_dxx,
    validar_rem_profile,
    validar_repetitive_manufacturing,
    validar_source_list_required_marc,
    validar_storage_location,
    validar_strategy_group,
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

def run_full_validation(
    mara_file,
    marc_file,
    mvke_file,
    marm_file=None,
    mbew_file=None,
    roles_file=None,
    zplm_marc_file=None,
    zplm_mara_file=None,
    previous_report=None,
    plant_filter: str | None = None,
    progress_callback=None,
    csr_item_file=None,
    csr_head_file=None,
    matplsta_file=None,
    mard_file=None,
):
    """
    Runs the complete Palmela validation pipeline:
      1. Load files
      2. Pre-normalisation + build sap_like_df (MARC + MARA + ZPLM MARC)
      3. All field validations
      4. Post-processing (Views / Roles / Summary)
      5. Excel generation in-memory
    Returns: (excel_bytes, errors_df, summary_df)

    NOTE: This engine no longer depends on the Item Browse Report / SAPUI5
    Export. Like the Tunisia engine, it builds an equivalent flat dataframe
    (`sap_like_df`) from MARC + MARA + ZPLM MARC via build_sap_like_df(),22
    and runs the exact same validar_* functions against it.
    """
    def _cb(pct, msg):
        if progress_callback:
            progress_callback(pct, msg)

    # ------------------------------------------------------------------
    # 1) Load DataFrames
    # ------------------------------------------------------------------
    _cb(7,  "Loading MARA...")
    mara_df = load_xlsx_fast(mara_file, sheet_name=0)

    _cb(10, "Loading MARC...")
    marc_df = load_xlsx_fast(marc_file, sheet_name=0)

    _cb(12, "Loading MVKE...")
    mvke_df = load_xlsx_fast(mvke_file, sheet_name=0)

    if roles_file is None:
        roles_file = get_rolesviews_bytesio()
    rolesviews_dict = load_xlsx_fast(roles_file, sheet_name=None)

    _cb(14, "Loading ZPLM data...")
    zplm_marc_df = load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None

    _cb(16, "Preparing ZPLM data...")
    zplm_mara_df = load_xlsx_fast(zplm_mara_file, sheet_name=0) if zplm_mara_file else None

    _cb(18, "Loading CSR data...")
    csr_item_df = load_xlsx_fast(csr_item_file, sheet_name=0) if csr_item_file else None
    csr_head_df = load_xlsx_fast(csr_head_file, sheet_name=0) if csr_head_file else None

    _cb(19, "Loading MATPLSTA data...")
    matplsta_df = load_xlsx_fast(matplsta_file, sheet_name=0) if matplsta_file else None

    _cb(20, "Loading MARM data...")
    marm_df = load_xlsx_fast(marm_file, sheet_name=0) if marm_file else None

    _cb(21, "Loading MBEW data...")
    mbew_df = load_xlsx_fast(mbew_file, sheet_name=0) if mbew_file else None

    _cb(21, "Loading MARD data...")
    mard_df = load_xlsx_fast(mard_file, sheet_name=0) if mard_file else None

    # ------------------------------------------------------------------
    # 2) Basic sanity checks
    # ------------------------------------------------------------------
    if mara_df is None:
        raise ValueError("MARA.xlsx not loaded.")
    if marc_df is None:
        raise ValueError("MARC.xlsx not loaded.")
    if mvke_df is None:
        raise ValueError("MVKE.xlsx not loaded.")
    if marm_df is None:
        raise ValueError("MARM.xlsx not loaded.")
    if mbew_df is None:
        raise ValueError("MBEW.xlsx not loaded.")
    if "Roles" not in rolesviews_dict or "Views" not in rolesviews_dict:
        raise ValueError("RolesViews.xlsx must contain sheets 'Roles' and 'Views'.")

    roles_df = rolesviews_dict["Roles"]
    views_df = rolesviews_dict["Views"]

    # ------------------------------------------------------------------
    # 2b) Column normalisation — same rename maps used by the TN engine,
    # since raw MARC/MBEW/ZPLM MARC/MVKE/MARM exports may use column names
    # that differ slightly from what the validar_* functions expect (these
    # used to be normalised for free by the Item Browse Report's own column
    # layout; now that we build sap_like_df ourselves, we normalise here).
    # ------------------------------------------------------------------
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
        'plant-specific material status': 'Plant Specific Material Status',
        'special procurement type': 'Special Procurement Type',
    })
    # ✅ 'Strategy Group' alias — validar_strategy_group() (generic, Palmela-only)
    # expects this exact name (matching the old Item Browse Report column),
    # while everything else (including TN's own tn_validar_strategy_group)
    # was written against 'Planning Strategy Group'. Keep both.
    if 'Planning Strategy Group' in marc_df.columns:
        marc_df['Strategy Group'] = marc_df['Planning Strategy Group']

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
    # ✅ Keep a 'Material Plant Type' alias alongside the normalized 'Plant Type'.
    # validar_repetitive_manufacturing() and validar_rem_profile() read
    # zplm_marc_df directly (not sap_like_df) and expect the original raw
    # column name; every other zplm_marc_df consumer expects 'Plant Type'.
    if zplm_marc_df is not None and 'Plant Type' in zplm_marc_df.columns:
        zplm_marc_df['Material Plant Type'] = zplm_marc_df['Plant Type']

    # ✅ DEDUPLICATE ZPLM MARC to one row per (Material, Plant) — BEFORE any
    # consumer touches it (build_sap_like_df AND the several validar_*
    # functions that merge marc_df with zplm_marc_df directly: Storage
    # Location for Ext. Procurement, Version Indicator, Batch Management
    # (Plant), Country/Region of Origin, Repetitive Manufacturing, REM
    # Profile). ZPLM MARC carries change-history columns (Original/Change
    # Number, Created/Changed By/On), which strongly suggests it can
    # legitimately hold more than one historical row per (Material, Plant).
    # A left.merge(zplm_marc_df, on=['Material','Plant']) multiplies every
    # matching row, so without this, every one of those validations would
    # inflate massively. We keep the most recently changed row per
    # (Material, Plant) when a date/change column is available, else just
    # the first occurrence.
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

    # ------------------------------------------------------------------
    # 2c) Hard-scope this engine to Palmela plants only (PT01/PT02), and
    # narrow further to a single plant if `plant_filter` asks for one.
    #
    # marc_file / mbew_file / zplm_marc_file may be *shared, global*
    # exports (e.g. when app.py runs the combined "All Sites" report and
    # hands the very same MARC file to both this engine and
    # run_tn_validation()). Without this guard, this engine would validate
    # TN01/TN02 rows too — using Palmela's generic rules instead of the
    # TN-exception-aware rules — producing spurious errors for Tunisia
    # materials that don't match (and therefore aren't removed by) the
    # app.py dedup step, since the error text differs from what the TN
    # engine legitimately produces for the same Material/Plant/Field.
    #
    # For a standalone Palmela-only run this is a no-op, since the
    # uploaded files already only contain PT01/PT02.
    # ------------------------------------------------------------------
    PALMELA_PLANTS = {"PT01", "PT02"}
    pt_scope = set(PALMELA_PLANTS)
    if plant_filter:
        pf = str(plant_filter).strip().upper()
        if pf in PALMELA_PLANTS:
            pt_scope = {pf}

    def _fpt(df, col='Plant'):
        if df is None or df.empty or col not in df.columns:
            return df
        return df[df[col].astype(str).str.strip().str.upper().isin(pt_scope)].copy()

    marc_df      = _fpt(marc_df)
    mbew_df      = _fpt(mbew_df)
    mard_df      = _fpt(mard_df) if mard_df is not None else None
    zplm_marc_df = _fpt(zplm_marc_df) if zplm_marc_df is not None else None

    # ------------------------------------------------------------------
    # 3) Pre-normalisation
    # ------------------------------------------------------------------
    marc_df = ensure_basic_data_k(marc_df)
    marc_df_full = marc_df.copy()

    # ------------------------------------------------------------------
    # 3b) Build sap_like_df (replaces the Item Browse Report / SAPUI5 Export)
    # ------------------------------------------------------------------
    _cb(18, "Building SAP-like dataframe from MARC + MARA + ZPLM MARC...")
    sap_like_df = build_sap_like_df(marc_df, mara_df, zplm_marc_df)

    # ✅ GLOBAL FILTER — ONLY FERT / HALB / ROH (Indirect Materials — ERSA,
    # HIBE, VERP — are handled by the separate Indirect Materials engine)
    sap_like_df = _filter_valid_material_types(sap_like_df, mara_df)
    advanced_planning_df = sap_like_df.copy()

    # Exclude Material Plant Status 99 rows from all validations except
    # Advanced Planning.
    exclude_mps99_pairs = _get_mps99_material_plant_pairs(marc_df_full)
    if exclude_mps99_pairs:
        sap_like_df = _exclude_material_plant_status_99(sap_like_df, exclude_mps99_pairs)
        marc_df = _exclude_material_plant_status_99(marc_df, exclude_mps99_pairs)
        if zplm_marc_df is not None:
            zplm_marc_df = _exclude_material_plant_status_99(
                zplm_marc_df,
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

    marc_df     = _filter_valid_material_types(marc_df, mara_df)

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

    

    # ------------------------------------------------------------------
    # 4) Run all field validations
    # ------------------------------------------------------------------
    resultados = []

    # Mandatory fields
    for coluna in ['Material', 'Material Type', 'Description', 'Base UOM', 'Plant', 'Plant Type']:
        _cb(20, f"Validating mandatory field: {coluna}...")
        validar_campo_obrigatorio(sap_like_df, coluna, resultados)

    validations = [
        (22, lambda: validar_procurement_type(sap_like_df, resultados)),
        (24, lambda: validar_account_assignment_group(sap_like_df, resultados, mvke_df)),
        (26, lambda: validar_automatic_purchase_order(sap_like_df, resultados)),
        (28, lambda: validar_batch_management(sap_like_df, resultados)),
        (28, lambda: validar_batch_management_mara_global(sap_like_df, mara_df, resultados)),
        (30, lambda: validar_consumption_mode(sap_like_df, resultados)),
        (32, lambda: validar_consumption_per(marc_df, mara_df, resultados)),
        (33, lambda: validar_division(sap_like_df, resultados)),
        (34, lambda: validar_export_import_group(sap_like_df, resultados)),
        (36, lambda: validar_general_item_category_group(sap_like_df, resultados)),
        (38, lambda: validar_gross_weight(sap_like_df, resultados)),
        (40, lambda: validar_item_category_group(sap_like_df, resultados, mvke_df)),
        (42, lambda: validar_jit_delivery_sched(sap_like_df, resultados)),
        (44, lambda: validar_loading_group(sap_like_df, resultados)),
        (46, lambda: validar_lot_sizing_procedure(sap_like_df, resultados)),
        (47, lambda: validar_per_lot_sz_plng_cal(advanced_planning_df, resultados, mara_df, matplsta_df)),
        (48, lambda: validar_material_group(sap_like_df, resultados)),
        (50, lambda: validar_mrp_type(sap_like_df, resultados, zplm_marc_df, csr_item_df, csr_head_df)),
        (51, lambda: validar_individual_collective_requirements_marc(marc_df, mara_df, resultados)),
        (52, lambda: validar_net_weight(sap_like_df, resultados)),
        (54, lambda: validar_period_type(sap_like_df, resultados)),
        (55, lambda: validar_storage_location_ext_procurement(marc_df, zplm_marc_df, mara_df, resultados)),
        (56, lambda: validar_planning_calendar(sap_like_df, resultados)),
        (57, lambda: validar_stock_determination_group_marc(marc_df, mara_df, resultados)),
        (58, lambda: validar_planning_time_fence(sap_like_df, resultados)),
        (59, lambda: validar_version_indicator_marc(marc_df, zplm_marc_df, mara_df, resultados)),
        (60, lambda: validar_price_control(sap_like_df, resultados)),
        (62, lambda: validar_price_unit(sap_like_df, resultados)),
        (63, lambda: validar_product_hierarchy(sap_like_df, resultados, mvke_df, mara_df)),
        (63, lambda: validar_product_hierarchy_mvke(sap_like_df, resultados, mvke_df)),
        (64, lambda: validar_production_supervisor_marc(marc_df, mara_df, resultados)),
        (66, lambda: validar_product_hierarchy_vs_enterproj(mara_df, zplm_mara_df, resultados)),
        (67, lambda: validar_packaging_material_type_mara(mara_df, resultados)),
        (67, lambda: validar_wh_material_group_mara(mara_df, resultados)),
        (67, lambda: validar_reference_matl_for_packing_mara(mara_df, resultados)),
        (67, lambda: validar_maintenance_status_c_roh_mara(mara_df, resultados, sap_like_df)),
        (67, lambda: validar_special_procurement_type_pt02_transfer(marc_df, zplm_marc_df, resultados)),
        (68, lambda: validar_production_scheduling_profile(sap_like_df, resultados)),
        (69, lambda: validar_country_region_origin_fp(marc_df, zplm_marc_df, mara_df, resultados)),
        (69, lambda: validar_commodity_code_fp(marc_df, zplm_marc_df, mara_df, resultados)),
        (70, lambda: validar_profit_center(sap_like_df, resultados)),
        (71, lambda: validar_lifo_fifo_relevant_mbew(mbew_df, mara_df, resultados)),
        (72, lambda: validar_purchasing_group(sap_like_df, resultados)),
        (74, lambda: validar_repetitive_manufacturing(marc_df, zplm_marc_df, mara_df, resultados)),
        (76, lambda: validar_rem_profile(marc_df, zplm_marc_df, mara_df, resultados)),
        (78, lambda: validar_source_list_required_marc(marc_df, mara_df, resultados)),
        (80, lambda: validar_storage_location(sap_like_df, resultados)),
        (82, lambda: validar_valuation_class(sap_like_df, resultados)),
        (84, lambda: validar_with_qty_structure(sap_like_df, resultados)),
        (86, lambda: validar_enterproj_number(sap_like_df, resultados)),
        (87, lambda: validar_mrp_controller(sap_like_df, resultados, mard_df)),
        (88, lambda: validar_plant_specific_material_status(sap_like_df, resultados)),
        (89, lambda: validar_backflush(sap_like_df, resultados)),
        (90, lambda: validar_strategy_group(sap_like_df, resultados)),
        (91, lambda: validar_product_related_origin(sap_like_df, resultados)),
        (92, lambda: validar_availability_check(sap_like_df, resultados)),
        (93, lambda: validar_material_group_appoint(sap_like_df, resultados)),
        (94, lambda: validar_purchasing_group_dxx(sap_like_df, resultados)),
    ]

    for pct, fn in validations:
        _cb(pct, "Running field validations...")
        fn()

    # ------------------------------------------------------------------
    # 4B) Cross-file validations (non-SAPUI5 tables)
    # ------------------------------------------------------------------
    _cb(95, "Validating MARM Alternative UoM rules...")
    validar_marm_pc_ea_roh_halb(
        marm_df=marm_df,
        mara_df=mara_df,
        resultados=resultados
    )

    _cb(96, "Validating Advanced Planning fields (MARC)...")
    validar_advanced_planning_fields(
        marc_df=advanced_planning_df,
        mara_df=mara_df,
        resultados=resultados,
        matplsta_df=matplsta_df,
    )

    _cb(97, "Validating Neg. stocks in plant (MARC)...")
    neg_errors_early = validar_neg_stocks_in_plant(marc_df, mara_df)
    if neg_errors_early:
        resultados.extend(neg_errors_early)

    # Batch Management: plant error in all plants + global error -> only global
    consolidar_batch_management_errors(resultados, sap_like_df)

    # ------------------------------------------------------------------
    # 5) Build errors DataFrame
    # ------------------------------------------------------------------
    errors_df = _coerce_df_for_display(pd.DataFrame(resultados))

    # ✅ Safety-net dedup: remove exact duplicates (same Material, Plant,
    # Field and Error). Different plants for the same material are fine
    # and expected (e.g. PT01 vs PT02) — only an exact repeat of the same
    # (Material, Plant, Field, Error) combination is removed here.
    _dedup_key = [c for c in ['Material', 'Plant', 'Field', 'Error'] if c in errors_df.columns]
    if len(_dedup_key) == 4 and not errors_df.empty:
        for _c in _dedup_key:
            errors_df[_c] = errors_df[_c].fillna('').astype(str).str.strip()
        errors_df = errors_df.drop_duplicates(subset=_dedup_key).reset_index(drop=True)

    # ------------------------------------------------------------
    # ✅ DETECT ERRORS NOT MAPPED IN ROLES (parity with TN engine)
    # ------------------------------------------------------------
    if roles_df is not None and not roles_df.empty:
        errors_df['Field_clean'] = errors_df['Field'].astype(str).str.strip()
        roles_set = set(roles_df['Field'].astype(str).str.strip())

        missing_role_df = errors_df[~errors_df['Field_clean'].isin(roles_set)].copy()

        if not missing_role_df.empty:
            print("\n--- ⚠️ ERRORS NOT MAPPED IN ROLES ---\n")
            for f in sorted(missing_role_df['Field_clean'].unique()):
                print(f)
            print("\n--------------------------------------\n")

        errors_df = errors_df.drop(columns=['Field_clean'], errors='ignore')
    else:
        missing_role_df = pd.DataFrame()

    # Enterproj enrichment (optional — needs previous report)
    if previous_report:
        errors_df = fill_plant_and_validate_enterproj_from_zplm(
            errors_df, latest_report=previous_report,
            zplm_marc_file=zplm_marc_file, zplm_mara_file=zplm_mara_file
        )

    # Additional enrichments
    if not errors_df.empty:
        errors_df = add_old_material_number_errors(mara_df, errors_df)

    # ------------------------------------------------------------------
    # 6) Post-processing
    # ------------------------------------------------------------------
    (
        errors_df, errors_filtered_df, views_missing_output_df,
        contagem_final_df, coverage_df, view_tabs_dict
    ) = postprocess_with_views_roles_and_classification(
        errors_df, mara_df, marc_df, roles_df, views_df, plant_filter=plant_filter
    )

    # Consolidate Missing_<View> tabs
    def _concat_missing_tabs(vtabs):
        cols = ['Material', 'Plant', 'Missing_View', 'View_Description', 'Fields_Expected']
        frames = [dfv[cols] for dfv in vtabs.values() if dfv is not None and not dfv.empty]
        if not frames:
            return pd.DataFrame(columns=cols)
        out = pd.concat(frames, ignore_index=True).drop_duplicates()
        out.sort_values(['Material', 'Plant', 'Missing_View'], inplace=True)
        return out

    view_is_missing_df = _concat_missing_tabs(view_tabs_dict)

    # ------------------------------------------------------------------
    # 7) Write Excel in-memory  (sheet order: Exec Summary first, raw data last)
    # ------------------------------------------------------------------
    _cb(100, "Generating final Excel report...")
    file_bytes = io.BytesIO()

    # Sort Adjusted Summary before writing (% of Total removed per user request)
    if not contagem_final_df.empty and 'Adjusted Nº of Errors' in contagem_final_df.columns:
        contagem_final_df = contagem_final_df.sort_values(
            'Adjusted Nº of Errors', ascending=False
        ).reset_index(drop=True)

    # Final guard: field names exactly as displayed in the report
    errors_df               = apply_display_field_renames(errors_df)
    views_missing_output_df = apply_display_field_renames(views_missing_output_df)
    contagem_final_df       = apply_display_field_renames(contagem_final_df)

    with pd.ExcelWriter(file_bytes, engine="openpyxl") as writer:
        pd.DataFrame().to_excel(writer, sheet_name="Executive Summary", index=False)

        base_row = 1
        if not contagem_final_df.empty:
            contagem_final_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=1)
            base_row = len(contagem_final_df) + 5
        if not coverage_df.empty:
            coverage_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=base_row)

        views_missing_output_df.to_excel(writer, sheet_name="Field With Error", index=False)
        errors_df.to_excel(writer, sheet_name="All Errors", index=False)
        view_is_missing_df.to_excel(writer, sheet_name="View is Missing", index=False)

        # Missing Role Mapping sheet (parity with TN engine)
        if not missing_role_df.empty:
            missing_role_df[['Field', 'Material', 'Plant', 'Error']].to_excel(
                writer, sheet_name="Missing Role Mapping", index=False
            )

        for view_code, df_view in view_tabs_dict.items():
            df_view.to_excel(writer, sheet_name=f"Missing_{view_code}"[:31], index=False)

        # Assignments sheet — enriched with owner/status from master.json
        try:
            _enriched = merge_assignments(errors_df, roles_df=roles_df)
            _asgn_df  = build_assignments_sheet(_enriched)
            _asgn_df.to_excel(writer, sheet_name="Assignments", index=False)
            save_open_assignments(_enriched)
        except Exception:
            pass

        marc_df.to_excel(writer, sheet_name="Original Data (MARC)", index=False)

    # Apply full professional styling
    file_bytes.seek(0)
    file_bytes = _style_sap_report(
        file_bytes,
        errors_df=errors_df,
        views_missing_output_df=views_missing_output_df,
        contagem_final_df=contagem_final_df,
        coverage_df=coverage_df,
        view_tabs_dict=view_tabs_dict,
        sap_df=marc_df,
        plant_filter=plant_filter,
    )

    # ------------------------------------------------------------------
    # 8) Build summary for Streamlit display
    # ------------------------------------------------------------------
    if 'Field' in errors_df.columns and not errors_df.empty:
        summary_df = (
            errors_df['Field'].value_counts()
            .rename_axis('Field').reset_index(name='Adjusted Nº of Errors')
        )
    else:
        summary_df = pd.DataFrame({"Field": [], "Adjusted Nº of Errors": []})

    return file_bytes, errors_df, summary_df
