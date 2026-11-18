"""Pós-processamento: Views, Roles e classificação dos erros."""

import pandas as pd

from engines.material_rules.batch_utils import apply_display_field_renames
from engines.material_rules.constants import FIELD_DISPLAY_RENAMES, MARA_MAINT_STATUS_C_FIELD

def postprocess_with_views_roles_and_classification(
    errors_df: pd.DataFrame,
    mara_df:   pd.DataFrame,
    marc_df:   pd.DataFrame,
    roles_df:  pd.DataFrame,
    views_df:  pd.DataFrame,
    plant_filter: str | None = None,
):
    """
    Full post-processing pipeline:
      - Maps Field -> View (via Roles)
      - Computes Letters_Present / Missing per (Material, Plant) from MARC 'Maintenance status'
      - Generates views_missing_df (per Material/Plant)
      - Flags Justified_By_Missing_View and classifies Error_Type (Missing View / Field Error)
      - Builds coverage_df (how many materials are missing each required view)
      - Builds view_tabs_dict (per-view detail tabs: Missing_<View>)
      - Constructs contagem_final_df (Adjusted Summary) with Role/View_Description overrides
      - Renames 'Export/Import Group' -> 'Intrastat Group'

    Returns:
        errors_df, errors_filtered_df, views_missing_output_df,
        contagem_final_df, coverage_df, view_tabs_dict
    """
    if errors_df is None:
        errors_df = pd.DataFrame()
    if roles_df is None or views_df is None:
        empty = pd.DataFrame()
        return errors_df, empty, empty, empty, empty, {}

    roles_df  = roles_df.copy()
    views_df  = views_df.copy()
    marc_df   = marc_df.copy()
    mara_df   = mara_df.copy()
    errors_df = errors_df.copy()

    # -----------------------------------------------
    # 1) Apply plant filter early
    # -----------------------------------------------
    if plant_filter is not None:
        plant_filter = str(plant_filter).strip().upper()

        mats_in_plant = set()
        if 'Material' in marc_df.columns and 'Plant' in marc_df.columns:
            mats_in_plant = set(
                marc_df.loc[marc_df['Plant'].astype(str).str.strip().str.upper() == plant_filter, 'Material']
                       .astype(str).str.strip()
            )

        if 'Plant' in marc_df.columns:
            marc_df = marc_df[marc_df['Plant'].astype(str).str.strip().str.upper() == plant_filter].copy()

        if 'Plant' in errors_df.columns and 'Material' in errors_df.columns:
            plant_col = errors_df['Plant'].astype(str).str.strip().str.upper()
            mat_col   = errors_df['Material'].astype(str).str.strip()
            errors_df = errors_df[(plant_col == plant_filter) |
                                  ((plant_col == '') & mat_col.isin(mats_in_plant))].copy()

    # Normalise headers
    for df in [roles_df, views_df, marc_df, mara_df, errors_df]:
        df.columns = df.columns.astype(str).str.strip()

    # -----------------------------------------------
    # 2) Base mappings: Field -> View, View -> Description
    # -----------------------------------------------
    if 'Field' in roles_df.columns and 'View' in roles_df.columns:
        roles_df['Field'] = roles_df['Field'].astype(str).str.strip()
        roles_df['View']  = roles_df['View'].astype(str).str.strip()
        field_to_view = roles_df.set_index('Field')['View'].to_dict()
    else:
        field_to_view = {}

    if 'Codigo' in views_df.columns and 'Descricao' in views_df.columns:
        views_df['Codigo']    = views_df['Codigo'].astype(str).str.strip()
        views_df['Descricao'] = views_df['Descricao'].astype(str).str.strip()
        view_to_desc     = views_df.set_index('Codigo')['Descricao'].to_dict()
        expected_letters = set(views_df['Codigo'].dropna().unique())
    else:
        view_to_desc, expected_letters = {}, set()

    if 'Field' in errors_df.columns:
        errors_df['View'] = errors_df['Field'].map(field_to_view).fillna('Not applicable')
        # MARA Maintenance status 'C' rule -> view C (Classification)
        errors_df.loc[
            errors_df['Field'].astype(str).str.strip() == MARA_MAINT_STATUS_C_FIELD, 'View'
        ] = 'C'
    else:
        errors_df['View'] = 'Not applicable'

    # -----------------------------------------------
    # 3) Extract Letters_Present per (Material, Plant)
    # -----------------------------------------------
    if 'Maintenance status' not in marc_df.columns:
        marc_df['Maintenance status'] = ''

    marc_df['Letters_Present'] = marc_df['Maintenance status'].apply(
        lambda s: [] if pd.isna(s) else list(str(s).strip())
    )

    mt_map = dict(zip(
        mara_df.get('Material', pd.Series(dtype=str)).astype(str).str.strip(),
        mara_df.get('Material Type', pd.Series(dtype=str)).astype(str).str.strip()
    ))

    if 'Plant Type' in marc_df.columns:
        marc_df['_KEY'] = list(zip(
            marc_df['Material'].astype(str).str.strip(),
            marc_df['Plant'].astype(str).str.strip()
        ))
        pt_map = dict(zip(marc_df['_KEY'], marc_df['Plant Type'].astype(str).str.strip()))
    else:
        pt_map = {}

    # -----------------------------------------------
    # 4) Views Missing per (Material, Plant)
    # -----------------------------------------------
    valid_types  = {'FERT', 'HALB', 'ROH'}
    view_to_fields = roles_df.groupby('View')['Field'].apply(list).to_dict()

    valid_mats = mara_df.loc[
        mara_df['Material Type'].astype(str).str.strip().isin(valid_types), 'Material'
    ].astype(str).str.strip().unique()

    marc_filtered = marc_df[marc_df['Material'].astype(str).str.strip().isin(valid_mats)].copy()

    missing_rows = []
    for _, row in marc_filtered.iterrows():
        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()
        present  = set(row.get('Letters_Present', []) or [])
        missing  = expected_letters - present

        fields_missing       = []
        descriptions_missing = []
        for letter in missing:
            fields_missing.extend(view_to_fields.get(letter, []))
            descriptions_missing.append(view_to_desc.get(letter, ''))

        missing_rows.append({
            'Material':              material,
            'Plant':                 plant,
            'Letters_Present':       ','.join(sorted(present)),
            'Letters_Missing':       ','.join(sorted(missing)),
            'Fields_Missing':        ','.join(sorted(fields_missing)),
            'Descriptions_Missing':  ','.join(sorted(descriptions_missing)),
        })

    views_missing_df = pd.DataFrame(missing_rows)

    # -----------------------------------------------
    # 5) Classify Error_Type + Justified_By_Missing_View
    # Pre-build a dict (Material, Plant) -> set of missing view letters for O(1) lookup
    # -----------------------------------------------
    _missing_lookup: dict[tuple, set] = {}
    if not views_missing_df.empty:
        for _, _vmrow in views_missing_df.iterrows():
            _key = (_vmrow['Material'], _vmrow['Plant'])
            _raw = _vmrow.get('Letters_Missing', '')
            _missing_lookup[_key] = set(_raw.split(',')) if isinstance(_raw, str) and _raw else set()

    def classify_error(row):
        # The validation itself proves view 'C' is not maintained (MARA level)
        if str(row.get('Field', '')).strip() == MARA_MAINT_STATUS_C_FIELD:
            return 'Missing View'
        view = str(row.get('View', '')).strip()
        if view not in ('', 'Not applicable'):
            _key = (str(row.get('Material', '')).strip(), str(row.get('Plant', '')).strip())
            if view in _missing_lookup.get(_key, set()):
                return 'Missing View'
        return 'Field Error'

    if not errors_df.empty:
        errors_df['Error_Type']              = errors_df.apply(classify_error, axis=1)
        errors_df['Justified_By_Missing_View'] = errors_df['Error_Type'] == 'Missing View'
    else:
        errors_df['Error_Type']              = pd.Series(dtype=str)
        errors_df['Justified_By_Missing_View'] = pd.Series(dtype=bool)

    # -----------------------------------------------
    # 6) Split by type
    # -----------------------------------------------
    errors_filtered_df      = errors_df[errors_df['Justified_By_Missing_View']].copy()
    views_missing_output_df = errors_df[~errors_df['Justified_By_Missing_View']].copy()

    # -----------------------------------------------
    # 7) Adjusted Summary (Field Errors only)
    # -----------------------------------------------
    if not views_missing_output_df.empty and 'Field' in views_missing_output_df.columns:
        contagem_final_df = (
            views_missing_output_df['Field']
            .value_counts()
            .rename_axis('Field')
            .reset_index(name='Adjusted Nº of Errors')
        )
    else:
        contagem_final_df = pd.DataFrame(columns=['Field', 'Adjusted Nº of Errors'])

    if not contagem_final_df.empty:
        contagem_final_df['View']             = contagem_final_df['Field'].map(field_to_view).fillna('Not applicable')
        contagem_final_df['View_Description'] = contagem_final_df['View'].map(view_to_desc).fillna('Not applicable')

        if 'Role' in roles_df.columns:
            roles_df['Role']  = roles_df['Role'].astype(str).str.strip()
            field_to_role     = roles_df.set_index('Field')['Role'].to_dict()
            contagem_final_df['Role'] = contagem_final_df['Field'].map(field_to_role).fillna('Not applicable')

        # Specific overrides
        mask_override = contagem_final_df['Field'].isin(['Enterproj Number', 'Plant Type'])
        contagem_final_df.loc[mask_override, 'View_Description'] = 'Product Master Visteon Data'

    # -----------------------------------------------
    # 8) Field renames (Export/Import Group -> Intrastat Group)
    # -----------------------------------------------
    def apply_field_renames(df_):
        return apply_display_field_renames(df_)

    errors_df               = apply_field_renames(errors_df)
    errors_filtered_df      = apply_field_renames(errors_filtered_df)
    views_missing_output_df = apply_field_renames(views_missing_output_df)
    contagem_final_df       = apply_field_renames(contagem_final_df)

    for col in ['View', 'View_Description', 'Role']:
        if col in contagem_final_df.columns:
            contagem_final_df[col] = contagem_final_df[col].astype(str).replace(
                {'nan': 'Not applicable', 'None': 'Not applicable', '': 'Not applicable'}
            )

    # -----------------------------------------------
    # 9) Coverage per View
    # -----------------------------------------------
    field_rules = {
        'Production Scheduling Profile':         lambda mt, pt: mt in ['FERT', 'HALB'],
        'REM Profile':                           lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Repetitive Manufacturing':              lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Enterproj Number':                      lambda mt, pt: mt == 'FERT',
        'Procurement Type':                      lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'] and pt in ['PP', 'MP'],
        'Account Assignment Group':              lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Automatic Purchase Order Allowed':      lambda mt, pt: pt == 'PP',
        'Batch Management Required Ind.':        lambda mt, pt: pt in ['PP', 'COI'], 
        'Source List Required':                  lambda mt, pt: pt == 'PP',
        'Purchasing Group':                      lambda mt, pt: pt == 'PP',
        'MRP Type':                              lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Planning Time Fence':                   lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Lot Sizing Procedure':                  lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Per.Lot Sz.Plng Cal.':                  lambda mt, pt: mt in ['FERT', 'HALB'],
        'With Qty Structure':                    lambda mt, pt: mt in ['FERT', 'HALB'],
        'Item Category Group':                   lambda mt, pt: True,
        'General Item Category Group':           lambda mt, pt: True,
        'Item Category Group (MVKE)':            lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Loading Group':                         lambda mt, pt: True,
        'Division':                              lambda mt, pt: True,
        'Export/Import Group':                   lambda mt, pt: True,
        'Intrastat Group':                       lambda mt, pt: True,
        'Product Hierarchy':                     lambda mt, pt: mt == 'FERT',
        'Price Control':                         lambda mt, pt: True,
        'Price Unit':                            lambda mt, pt: True,
        'Valuation Class':                       lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Gross Weight':                          lambda mt, pt: pt == 'FP',
        'Net Weight':                            lambda mt, pt: pt == 'FP',
        'Planning Calendar':                     lambda mt, pt: True,
        'Period Type':                           lambda mt, pt: True,
        'Strategy Group':                        lambda mt, pt: True,
        'MRP Controller':                        lambda mt, pt: True,
        'Plant Specific Material Status':        lambda mt, pt: True,
        'Backflush':                             lambda mt, pt: True,
        'Product-Related Origin':                lambda mt, pt: True,
        'Availability Check':                    lambda mt, pt: True,
        'Neg. stocks in plant':                  lambda mt, pt: mt in ['ROH', 'HALB'],
        'Consumption Mode':                      lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Bwd consumption per.':                  lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Fwd consumption per.':                  lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Profit Center':                         lambda mt, pt: True,
        'Material Group':                        lambda mt, pt: True,
        'JIT Delivery Sched.':                   lambda mt, pt: pt == 'PP',
        'Issue Storage Location':                lambda mt, pt: True,
        'Cntry/Reg of Origin':                   lambda mt, pt: pt == 'FP',
        #'Commodity Code':                        lambda mt, pt: pt == 'FP',
        'LIFO/FIFO-relevant':                    lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Stock Determination Group':             lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Storage Location for Ext. Procurement': lambda mt, pt: pt == 'PP',
        'Individual/ Collective Requirements':   lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Packaging Mat. Type':                   lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],   
        'WH Material Group':                     lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Production Supervisor':                 lambda mt, pt: mt in ['FERT', 'HALB'],   
        'Batch Management(Plant)':               lambda mt, pt: pt in ['PP', 'COI'],
        'Advanced Planning':                     lambda mt, pt: True,
        'Heuristic':                             lambda mt, pt: True,
        'PP Plng Procedure':                     lambda mt, pt: True,
        'Proposed Strategy':                     lambda mt, pt: True,
        'Version Indicator':                     lambda mt, pt: pt in ['FP', 'MP'],
        'Costing Lot Size':                      lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Safety Time':                           lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
    }

    def is_view_required(material_type, plant_type, view_code):
        for field in view_to_fields.get(view_code, []):
            rule = field_rules.get(field)
            if rule and rule(material_type, plant_type):
                return True
        return False

    view_missing_count = {letter: 0 for letter in expected_letters}
    for _, r in marc_filtered.iterrows():
        mt      = str(mt_map.get(str(r['Material']).strip(), '')).strip().upper()
        pt      = str(pt_map.get((str(r['Material']).strip(), str(r['Plant']).strip()), '')).strip().upper()
        present = set(r.get('Letters_Present', []) or [])
        for v in expected_letters:
            if v not in present and is_view_required(mt, pt, v):
                view_missing_count[v] += 1

    coverage_df = pd.DataFrame([
        {'View': v, 'View_Description': view_to_desc.get(v, ''),
         'Materiais without Active View': cnt,
         'Link': f'=HYPERLINK("#Missing_{v}!A1","Check Details")'}
        for v, cnt in view_missing_count.items() if cnt > 0
    ])

    # -----------------------------------------------
    # 10) Missing_<View> detail tabs
    # -----------------------------------------------
    view_tabs_dict = {}
    for v in expected_letters:
        rows = []
        for _, r in marc_filtered.iterrows():
            mt      = str(mt_map.get(str(r['Material']).strip(), '')).strip().upper()
            pt      = str(pt_map.get((str(r['Material']).strip(), str(r['Plant']).strip()), '')).strip().upper()
            present = set(r.get('Letters_Present', []) or [])
            if v not in present and is_view_required(mt, pt, v):
                rows.append({
                    'Material':         str(r['Material']).strip(),
                    'Plant':            str(r['Plant']).strip(),
                    'Missing_View':     v,
                    'View_Description': view_to_desc.get(v, ''),
                    'Fields_Expected':  ','.join(FIELD_DISPLAY_RENAMES.get(_f, _f) for _f in view_to_fields.get(v, [])),
                })
        if rows:
            df_view = pd.DataFrame(rows)
            if not df_view.empty:
                df_view = df_view.drop_duplicates(subset=['Material', 'Plant', 'Missing_View']).reset_index(drop=True)
            view_tabs_dict[v] = df_view

    # -------------------------------------------------
    # View 'C' (Classification) missing at MARA level
    # (ROH without letter 'C' in MARA Maintenance status)
    # -> same output pattern as every other missing view:
    #    Missing_C tab, "View is Missing" sheet and coverage table.
    # -------------------------------------------------
    if not errors_df.empty and 'Field' in errors_df.columns:
        _c_errs = errors_df[errors_df['Field'].astype(str).str.strip() == MARA_MAINT_STATUS_C_FIELD]
        if not _c_errs.empty:
            _c_mats = set(_c_errs['Material'].astype(str).str.strip())
            _plants_by_mat = (
                marc_filtered.assign(
                    _M=marc_filtered['Material'].astype(str).str.strip(),
                    _P=marc_filtered['Plant'].astype(str).str.strip(),
                )
                .groupby('_M')['_P'].apply(lambda x: sorted(set(x))).to_dict()
            )
            _c_rows = []
            for _m in sorted(_c_mats):
                for _pl in (_plants_by_mat.get(_m) or ['']):
                    _c_rows.append({
                        'Material':         _m,
                        'Plant':            _pl,
                        'Missing_View':     'C',
                        'View_Description': view_to_desc.get('C', 'Classification'),
                        'Fields_Expected':  ','.join(FIELD_DISPLAY_RENAMES.get(_f, _f) for _f in view_to_fields.get('C', [])),
                    })
            _c_df = pd.DataFrame(_c_rows)
            if 'C' in view_tabs_dict and not view_tabs_dict['C'].empty:
                _c_df = pd.concat([view_tabs_dict['C'], _c_df], ignore_index=True)
            view_tabs_dict['C'] = _c_df.drop_duplicates(
                subset=['Material', 'Plant', 'Missing_View']
            ).reset_index(drop=True)

            _c_count = len(view_tabs_dict['C'])
            if not coverage_df.empty and (coverage_df['View'] == 'C').any():
                coverage_df.loc[coverage_df['View'] == 'C', 'Materiais without Active View'] = _c_count
            else:
                coverage_df = pd.concat([coverage_df, pd.DataFrame([{
                    'View': 'C',
                    'View_Description': view_to_desc.get('C', 'Classification'),
                    'Materiais without Active View': _c_count,
                    'Link': '=HYPERLINK("#Missing_C!A1","Check Details")',
                }])], ignore_index=True)

    # -------------------------------------------------
    # ADD MRP Controller from MARC (Material + Plant)
    # -------------------------------------------------
    if marc_df is not None and not marc_df.empty and "MRP Controller" in marc_df.columns:

        marc_tmp = marc_df[["Material", "Plant", "MRP Controller"]].copy()

        marc_tmp["Material"] = marc_tmp["Material"].astype(str).str.strip()
        marc_tmp["Plant"] = marc_tmp["Plant"].astype(str).str.strip()

        # Avoid repeated merges when MARC contains duplicate Material+Plant rows.
        marc_tmp = marc_tmp.drop_duplicates(subset=["Material", "Plant"], keep="first")

        views_missing_output_df["Material"] = views_missing_output_df["Material"].astype(str).str.strip()
        views_missing_output_df["Plant"] = views_missing_output_df["Plant"].astype(str).str.strip()

        views_missing_output_df = views_missing_output_df.merge(
            marc_tmp,
            on=["Material", "Plant"],
            how="left"
        )

        if not views_missing_output_df.empty:
            dedup_cols = [c for c in ['Material', 'Plant', 'Field', 'Error'] if c in views_missing_output_df.columns]
            if dedup_cols:
                views_missing_output_df = views_missing_output_df.drop_duplicates(subset=dedup_cols).reset_index(drop=True)
    else:
        views_missing_output_df["MRP Controller"] = ""

    # -------------------------------------------------
    # REMOVE "View" only from Field With Error dataset
    # -------------------------------------------------
    if "View" in views_missing_output_df.columns:
        views_missing_output_df = views_missing_output_df.drop(columns=["View"])

    return errors_df, errors_filtered_df, views_missing_output_df, contagem_final_df, coverage_df, view_tabs_dict
