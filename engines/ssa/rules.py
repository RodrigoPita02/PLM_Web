"""Regras de validação do SSA."""

import pandas as pd

from engines.ssa.constants import (
    SSA_ORDER_TYPES_ALL,
    SSA_ORDER_TYPES_CUSTOMER_GROUP_2,
    SSA_ORDER_TYPES_ZLZ_GROUP,
    SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP,
    SSA_ORDER_TYPES_ZTAM_GROUP,
    SSA_S_MATERIAL_LIST,
)
from engines.ssa.helpers import (
    _ssa_build_active_demand_set,
    _ssa_build_obd_set,
    _ssa_build_s_material_mvke_map,
    _ssa_doc_type_in,
    _ssa_doc_type_map,
    _ssa_first_nonempty,
    _ssa_float_equal,
    _ssa_fully_rejected_docs,
    _ssa_has_obd,
    _ssa_is_zero_or_blank,
    _ssa_item_is_done,
    _ssa_item_key,
    _ssa_load,
    _ssa_norm,
    _ssa_pick,
)
from engines.ssa.state import _SSA_ACTIVE_DOC_TYPE_MAP

def ssa_validar_mandatory_fields(vbak_df, vbap_df, marc_df=None, mvke_df=None):
    """
    Validate mandatory VBAK and VBAP fields.

    Rule applicability is based on Sales Document Type (AUART) only, through
    the SSA_ORDER_TYPES_ZLZ_GROUP / SSA_ORDER_TYPES_ZTAM_GROUP /
    SSA_ORDER_TYPES_ZDIC_GROUP / SSA_ORDER_TYPES_ZMD_GROUP sets. The Sales
    Document number is never inspected.

    VBAK:
      Customer Group checks are skipped when VBAK Overall Header (UVALL)
      is equal to "C".

      - Customer Group 1:
            Mandatory for Sales Document Types in SSA_ORDER_TYPES_ALL
            (ZLZ/ZLZM, ZTAM/ZOR/ZPRO/ZFD/ZINI/ZORA/ZFD2/ZORM,
            ZDIC/ZMC/ZRE/ZDSC/ZRE2/ZMC2, ZMD/ZDSD/ZDID/ZMD2).

      - Customer Group 2:
            Mandatory only for Sales Document Types ZLZ, ZLZM, ZOR, ZTAM
            (SSA_ORDER_TYPES_CUSTOMER_GROUP_2).

      - Customer Group 3:
            Mandatory for Sales Document Types in SSA_ORDER_TYPES_ZLZ_GROUP,
            with the following refinement:
              * ZLZ:
                    Mandatory only when VBAK-ABDIS ("MRP for DlvSchType") is not "E".
              * ZLZM:
                    Always mandatory.

    VBAP:
      Items where Reason for Rejection (VBAP-ABGRU) is filled are skipped.

      - Route:
            Mandatory for Sales Document Types in SSA_ORDER_TYPES_ZLZ_GROUP.

      - IFS ID:
            Mandatory for Sales Document Types in
            SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP.

      - Supplier_code:
            Mandatory for Sales Document Types in
            SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP when Sales Document Type
            is one of:
                ZOR, ZTAM, ZORA
            (ZPRO is explicitly excluded from this check.)

      - Plant:
            Mandatory for Sales Document Types in
            SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP.

      - Acct Assmt Grp Mat. (KTGRM):

            Applies only to active items, meaning VBAP-ABGRU is blank.

            KTGRM must not be blank for the following Sales Document Types:
                ZLZ, ZLZM, ZRE, ZPRO, ZOR, ZORM, ZFD, ZFD2, ZRE2,
                ZMD, ZMD2, ZMC, ZMC2, ZDID, ZDIC

            Additional validations:

              1. For ZLZ, ZLZM and ZRE:
                   If the Material Number does not start with "S-",
                   KTGRM must be equal to "Z1".

                   If the Material Number starts with "S-", only the
                   mandatory non-blank validation applies.

              2. For ZPRO, ZOR, ZORM, ZFD, ZFD2, ZRE2, ZMD, ZMD2,
                 ZMC, ZMC2, ZDID and ZDIC:

                   If the Material Number starts with "S-" and is one of the
                   S-materials in SSA_S_MATERIAL_LIST, KTGRM must match the
                   value found for that material in MVKE (field
                   "Acct Assmt Grp Mat.").

                   If the Material Number starts with "S-" but is not
                   included in SSA_S_MATERIAL_LIST (or MVKE carries no value
                   for it), KTGRM must only be non-blank.

                   For materials that do not start with "S-", KTGRM must
                   only be non-blank.

    Note:
      marc_df is retained in the function signature for backward compatibility
      with the existing orchestrator, but it is no longer used by the KTGRM
      validation. mvke_df is used only for the S-material comparison
      described above; when it is not supplied, S-materials fall back to
      the plain non-blank check.
    """
    errors = []
    doc_type_map = _ssa_doc_type_map(vbak_df)

    # ── VBAK ──────────────────────────────────────────────────────────────────
    if vbak_df is not None and not vbak_df.empty:
        vbak = _ssa_load(vbak_df)

        doc_col = _ssa_pick(
            vbak,
            ["Sales Document", "VBELN"],
        )

        abdis_col = _ssa_pick(
            vbak,
            ["MRP for DlvSchType", "MRPforDS", "ABDIS"],
        )

        header_status_col = _ssa_pick(
            vbak,
            [
                "Overall Header",
                "UVALL",
            ],
        )

        def is_completed_document(row):
            """Skip Customer Group checks when VBAK Overall Header = 'C'."""
            if header_status_col is None:
                return False

            return _ssa_norm(row[header_status_col]).upper() == "C"

        # ── Customer Group 1 / Customer Group 2 ───────────────────────────────
        field_order_types = {
            "Customer Group 1": SSA_ORDER_TYPES_ALL,
            "Customer Group 2": SSA_ORDER_TYPES_CUSTOMER_GROUP_2,
        }

        for field, allowed_types in field_order_types.items():
            col = _ssa_pick(vbak, [field])

            # If the column does not exist in the input file,
            # this specific validation cannot be performed.
            if col is None:
                continue

            for _, row in vbak.iterrows():
                # Skip completed sales documents.
                if is_completed_document(row):
                    continue

                doc = (
                    _ssa_norm(row[doc_col])
                    if doc_col
                    else ""
                )

                # Skip document types outside the scope of the field.
                if not _ssa_doc_type_in(doc, allowed_types):
                    continue

                # Generate an error when the applicable field is blank.
                if _ssa_norm(row[col]) == "":
                    errors.append({
                        "Table": "VBAK",
                        "Sales Document": doc,
                        "Item": "",
                        "Material": "",
                        "Plant": "",
                        "Field": field,
                        "Value": "",
                        "Error": (
                            f"Mandatory field blank: '{field}'"
                        ),
                    })

        # ── Customer Group 3 ──────────────────────────────────────────────────
        cg3_col = _ssa_pick(
            vbak,
            ["Customer Group 3"],
        )

        if cg3_col:
            for _, row in vbak.iterrows():
                # Skip completed sales documents.
                if is_completed_document(row):
                    continue

                doc = (
                    _ssa_norm(row[doc_col])
                    if doc_col
                    else ""
                )

                # Customer Group 3 applies only to the ZLZ document-type group.
                if not _ssa_doc_type_in(
                    doc,
                    SSA_ORDER_TYPES_ZLZ_GROUP,
                ):
                    continue

                doc_type = doc_type_map.get(doc, "").upper()

                # For ZLZ, Customer Group 3 is not mandatory when
                # MRP for DlvSchType is equal to E.
                if doc_type == "ZLZ":
                    abdis_val = (
                        _ssa_norm(row[abdis_col]).upper()
                        if abdis_col
                        else ""
                    )

                    if abdis_val == "E":
                        continue

                if _ssa_norm(row[cg3_col]) == "":
                    errors.append({
                        "Table": "VBAK",
                        "Sales Document": doc,
                        "Item": "",
                        "Material": "",
                        "Plant": "",
                        "Field": "Customer Group 3",
                        "Value": "",
                        "Error": (
                            "Mandatory field blank: "
                            "'Customer Group 3'"
                        ),
                    })

    # ── VBAP ──────────────────────────────────────────────────────────────────
    if vbap_df is not None and not vbap_df.empty:
        vbap = _ssa_load(vbap_df)

        doc_col = _ssa_pick(
            vbap,
            ["Sales Document", "VBELN"],
        )
        item_col = _ssa_pick(
            vbap,
            ["Sales Document Item", "Item (SD)", "POSNR"],
        )
        mat_col = _ssa_pick(
            vbap,
            ["Material", "MATNR"],
        )
        plt_col = _ssa_pick(
            vbap,
            ["Plant", "WERKS"],
        )
        rej_col = _ssa_pick(
            vbap,
            ["Reason for Rejection", "ABGRU"],
        )
        status_col = _ssa_pick(
            vbap,
            ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"],
        )

        # Items that are complete (Overall Status = 'C') or rejected
        # (VBAP-ABGRU filled) are excluded from all mandatory-field checks.
        def is_rejected(row):
            return _ssa_item_is_done(row, status_col, rej_col)

        # ── Route ─────────────────────────────────────────────────────────────
        route_col = _ssa_pick(
            vbap,
            ["Route", "ROUTE"],
        )

        if route_col:
            for _, row in vbap.iterrows():
                if is_rejected(row):
                    continue

                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_GROUP):
                    continue

                if _ssa_norm(row[route_col]) == "":
                    errors.append({
                        "Table": "VBAP",
                        "Sales Document": doc,
                        "Item": (
                            _ssa_norm(row[item_col])
                            if item_col
                            else ""
                        ),
                        "Material": (
                            _ssa_norm(row[mat_col])
                            if mat_col
                            else ""
                        ),
                        "Plant": (
                            _ssa_norm(row[plt_col])
                            if plt_col
                            else ""
                        ),
                        "Field": "Route",
                        "Value": "",
                        "Error": "Mandatory field blank: 'Route'",
                    })

        # ── IFS ID ────────────────────────────────────────────────────────────
        ifs_col = _ssa_pick(
            vbap,
            ["IFS ID"],
        )

        if ifs_col:
            for _, row in vbap.iterrows():
                if is_rejected(row):
                    continue

                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
                    continue

                if _ssa_norm(row[ifs_col]) == "":
                    errors.append({
                        "Table": "VBAP",
                        "Sales Document": doc,
                        "Item": (
                            _ssa_norm(row[item_col])
                            if item_col
                            else ""
                        ),
                        "Material": (
                            _ssa_norm(row[mat_col])
                            if mat_col
                            else ""
                        ),
                        "Plant": (
                            _ssa_norm(row[plt_col])
                            if plt_col
                            else ""
                        ),
                        "Field": "IFS ID",
                        "Value": "",
                        "Error": "Mandatory field blank: 'IFS ID'",
                    })

        # ── Supplier_code ─────────────────────────────────────────────────────
        sup_col = _ssa_pick(
            vbap,
            ["Supplier_code"],
        )

        if sup_col:
            # ZPRO is excluded from the Supplier_code check.
            supplier_allowed_types = {
                "ZOR",
                "ZTAM",
                "ZORA",
            }

            for _, row in vbap.iterrows():
                if is_rejected(row):
                    continue

                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
                    continue

                doc_type = doc_type_map.get(doc, "").upper()

                if doc_type not in supplier_allowed_types:
                    continue

                if _ssa_norm(row[sup_col]) == "":
                    errors.append({
                        "Table": "VBAP",
                        "Sales Document": doc,
                        "Item": (
                            _ssa_norm(row[item_col])
                            if item_col
                            else ""
                        ),
                        "Material": (
                            _ssa_norm(row[mat_col])
                            if mat_col
                            else ""
                        ),
                        "Plant": (
                            _ssa_norm(row[plt_col])
                            if plt_col
                            else ""
                        ),
                        "Field": "Supplier_code",
                        "Value": "",
                        "Error": (
                            "Mandatory field blank: 'Supplier_code'"
                        ),
                    })

        # ── Plant ─────────────────────────────────────────────────────────────
        if plt_col:
            for _, row in vbap.iterrows():
                if is_rejected(row):
                    continue

                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
                    continue

                if _ssa_norm(row[plt_col]) == "":
                    errors.append({
                        "Table": "VBAP",
                        "Sales Document": doc,
                        "Item": (
                            _ssa_norm(row[item_col])
                            if item_col
                            else ""
                        ),
                        "Material": (
                            _ssa_norm(row[mat_col])
                            if mat_col
                            else ""
                        ),
                        "Plant": "",
                        "Field": "Plant",
                        "Value": "",
                        "Error": "Mandatory field blank: 'Plant'",
                    })

        # ── Acct Assmt Grp Mat. (KTGRM) ──────────────────────────────────────
        aag_col = _ssa_pick(
            vbap,
            [
                "Acct Assmt Grp Mat.",
                "Account Assignment Group",
                "Acct Assignment Grp",
                "Acc.Assignment Grp.",
                "KTGRM",
            ],
        )

        if aag_col:
            # KTGRM must not be blank for active items belonging to any of
            # these Sales Document Types.
            mandatory_ktgrm_doc_types = {
                "ZLZ",
                "ZLZM",
                "ZRE",
                "ZPRO",
                "ZOR",
                "ZORM",
                "ZFD",
                "ZFD2",
                "ZRE2",
                "ZMD",
                "ZMD2",
                "ZMC",
                "ZMC2",
                "ZDID",
                "ZDIC",
            }

            # For these Sales Document Types, non-S materials require Z1.
            non_s_z1_doc_types = {
                "ZLZ",
                "ZLZM",
                "ZRE",
            }

            # For these Sales Document Types, the approved S-material mapping
            # must be validated.
            s_material_mapping_doc_types = {
                "ZPRO",
                "ZOR",
                "ZORM",
                "ZFD",
                "ZFD2",
                "ZRE2",
                "ZMD",
                "ZMD2",
                "ZMC",
                "ZMC2",
                "ZDID",
                "ZDIC",
            }

            # S-materials in scope for the MVKE-based KTGRM comparison.
            # The expected KTGRM per material is no longer hardcoded here —
            # it is looked up live from MVKE (field "Acct Assmt Grp Mat."),
            # see SSA_S_MATERIAL_LIST / _ssa_build_s_material_mvke_map.
            s_material_ktgrm_map = _ssa_build_s_material_mvke_map(
                mvke_df, SSA_S_MATERIAL_LIST
            )

            for _, row in vbap.iterrows():
                # The rule applies only when VBAP-ABGRU is blank.
                if is_rejected(row):
                    continue

                doc = (
                    _ssa_norm(row[doc_col])
                    if doc_col
                    else ""
                )
                item = (
                    _ssa_norm(row[item_col])
                    if item_col
                    else ""
                )
                material = (
                    _ssa_norm(row[mat_col])
                    if mat_col
                    else ""
                )
                plant = (
                    _ssa_norm(row[plt_col])
                    if plt_col
                    else ""
                )

                doc_type = doc_type_map.get(doc, "").upper()
                material_upper = material.upper()
                ktgrm = _ssa_norm(row[aag_col]).upper()

                # The special KTGRM rule applies only to the document types
                # explicitly listed in the business rule.
                if doc_type not in mandatory_ktgrm_doc_types:
                    continue

                # Rule 1:
                # KTGRM must not be blank for all listed document types.
                if ktgrm == "":
                    errors.append({
                        "Table": "VBAP",
                        "Sales Document": doc,
                        "Item": item,
                        "Material": material,
                        "Plant": plant,
                        "Field": "Acct Assmt Grp Mat.",
                        "Value": "",
                        "Error": (
                            "Mandatory field blank: "
                            "'Acct Assmt Grp Mat.' "
                            f"for active item with Sales Document Type "
                            f"'{doc_type}'"
                        ),
                    })
                    continue

                # Rule 2:
                # For ZLZ, ZLZM and ZRE, non-S materials require KTGRM Z1.
                if (
                    doc_type in non_s_z1_doc_types
                    and not material_upper.startswith("S-")
                ):
                    expected_ktgrm = "Z1"

                    if ktgrm != expected_ktgrm:
                        errors.append({
                            "Table": "VBAP",
                            "Sales Document": doc,
                            "Item": item,
                            "Material": material,
                            "Plant": plant,
                            "Field": "Acct Assmt Grp Mat.",
                            "Value": ktgrm,
                            "Error": (
                                f"Invalid KTGRM '{ktgrm}': "
                                f"Sales Document Type '{doc_type}' with "
                                f"non-S material '{material}' requires "
                                f"KTGRM '{expected_ktgrm}'"
                            ),
                        })

                    continue

                # Rule 3:
                # For the configured document types, validate S-materials
                # included in the approved mapping.
                if (
                    doc_type in s_material_mapping_doc_types
                    and material_upper.startswith("S-")
                ):
                    expected_ktgrm = s_material_ktgrm_map.get(
                        material_upper
                    )

                    # If the S-material is not included in the approved map,
                    # only the mandatory-field check applies. Since KTGRM is
                    # already known to be filled, no error is required.
                    if expected_ktgrm is None:
                        continue

                    if ktgrm != expected_ktgrm:
                        errors.append({
                            "Table": "VBAP",
                            "Sales Document": doc,
                            "Item": item,
                            "Material": material,
                            "Plant": plant,
                            "Field": "Acct Assmt Grp Mat.",
                            "Value": ktgrm,
                            "Error": (
                                f"Invalid KTGRM '{ktgrm}': "
                                f"Material '{material_upper}' requires "
                                f"KTGRM '{expected_ktgrm}' (per MVKE) for "
                                f"Sales Document Type '{doc_type}'"
                            ),
                        })

    return pd.DataFrame(errors)


def ssa_validar_terms_of_payment(vbak_df, vbkd_df=None):
    """
    Header-level validation:
      - Terms of Payment must not be empty.

    Source:
      - VBKD~Terms of Payment / ZTERM
      - Header-level VBKD record:
            POSNR blank or numerically equal to zero

    Note:
      VBAK-ZTERM is not used by this check; only VBKD is inspected.
      vbak_df is still required to know which Sales Documents to check.

    Performance:
      - Builds the VBKD lookup once.
      - Does not filter or copy the complete VBKD DataFrame for every
        Sales Document in VBAK.
    """
    errors = []

    empty_cols = [
        "Table",
        "Sales Document",
        "Item",
        "Material",
        "Plant",
        "Field",
        "Value",
        "Error",
    ]

    if vbak_df is None or vbak_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbak = _ssa_load(vbak_df)

    vbak_doc = _ssa_pick(
        vbak,
        [
            "Sales Document",
            "Sales Doc.",
            "VBELN",
        ],
    )

    if vbak_doc is None:
        return pd.DataFrame(columns=empty_cols)

    # Normalize the relevant VBAK values once.
    vbak_docs = vbak[vbak_doc].map(_ssa_norm)

    # -------------------------------------------------------------------------
    # Build the VBKD header-level Terms of Payment lookup once
    # -------------------------------------------------------------------------
    vbkd_terms_map = {}

    if vbkd_df is not None and not vbkd_df.empty:
        vbkd = _ssa_load(vbkd_df)

        vbkd_doc = _ssa_pick(
            vbkd,
            [
                "SD Document",
                "Sales Document",
                "Sales Doc.",
                "VBELN",
            ],
        )

        vbkd_item = _ssa_pick(
            vbkd,
            [
                "Item (SD)",
                "Sales Document Item",
                "POSNR",
                "Item",
            ],
        )

        vbkd_top = _ssa_pick(
            vbkd,
            [
                "Terms of Payment",
                "Terms of Payment Key",
                "ZTERM",
            ],
        )

        if vbkd_doc is not None and vbkd_top is not None:
            required_columns = [vbkd_doc, vbkd_top]

            if vbkd_item is not None:
                required_columns.append(vbkd_item)

            # Prevent duplicated column selection if aliases resolve
            # to the same physical column.
            required_columns = list(dict.fromkeys(required_columns))

            vbkd_work = vbkd.loc[:, required_columns].copy()

            vbkd_work["_ssa_doc"] = vbkd_work[vbkd_doc].map(_ssa_norm)
            vbkd_work["_ssa_terms"] = vbkd_work[vbkd_top].map(_ssa_norm)

            if vbkd_item is not None:
                vbkd_work["_ssa_item"] = vbkd_work[vbkd_item].map(_ssa_norm)

                # SAP item numbers can be imported from Excel in several forms:
                # '', 0, 0.0, 000000, etc.
                header_mask = vbkd_work["_ssa_item"].apply(
                    _ssa_is_zero_or_blank
                )

                vbkd_work = vbkd_work.loc[header_mask]

            # Remove rows that cannot contribute to the fallback lookup.
            vbkd_work = vbkd_work.loc[
                (vbkd_work["_ssa_doc"] != "")
                & (vbkd_work["_ssa_terms"] != "")
            ]

            if not vbkd_work.empty:
                # Preserve the original functional behavior:
                # use the first non-empty Terms of Payment found for
                # each Sales Document at header level.
                vbkd_terms_map = (
                    vbkd_work
                    .drop_duplicates(
                        subset=["_ssa_doc"],
                        keep="first",
                    )
                    .set_index("_ssa_doc")["_ssa_terms"]
                    .to_dict()
                )

    # -------------------------------------------------------------------------
    # Validate — one Sales Document per VBAK row, value sourced from VBKD only
    # -------------------------------------------------------------------------
    for row_index in vbak.index:
        doc = vbak_docs.at[row_index]
        value = vbkd_terms_map.get(doc, "")

        if value == "":
            errors.append({
                "Table": "VBKD",
                "Sales Document": doc,
                "Item": "",
                "Material": "",
                "Plant": "",
                "Field": "Terms of Payment",
                "Value": "",
                "Error": "Mandatory field blank: 'Terms of Payment'",
            })

    if not errors:
        return pd.DataFrame(columns=empty_cols)

    return pd.DataFrame(errors, columns=empty_cols)


def ssa_validar_shipping_point(vbap_df):
    """
    Item-level (only for Sales Document Types in SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
    - Shipping Point must not be empty
    - Shipping Point must be equal to Plant
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbap_df is None or vbap_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbap = _ssa_load(vbap_df)

    doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "POSNR"])
    mat_col  = _ssa_pick(vbap, ["Material", "MATNR"])
    plt_col  = _ssa_pick(vbap, ["Plant", "WERKS"])
    shp_col  = _ssa_pick(vbap, ["Shipping Point/Receiving Pt", "Shipping Point", "VSTEL"])
    status_col = _ssa_pick(vbap, ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"])
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

    if not all([doc_col, item_col, mat_col, plt_col, shp_col, status_col]):
        return pd.DataFrame(columns=empty_cols)

    for _, row in vbap.iterrows():
        if _ssa_item_is_done(row, status_col, rej_col):
            continue

        doc = _ssa_norm(row[doc_col])
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
            continue

        item  = _ssa_norm(row[item_col])
        mat   = _ssa_norm(row[mat_col])
        plant = _ssa_norm(row[plt_col])
        shp   = _ssa_norm(row[shp_col])

        if shp == '':
            errors.append({
                "Table": "VBAP",
                "Sales Document": doc,
                "Item": item,
                "Material": mat,
                "Plant": plant,
                "Field": "Shipping Point",
                "Value": "",
                "Error": "Mandatory field blank: 'Shipping Point'",
            })
        elif plant != '' and shp != plant:
            errors.append({
                "Table": "VBAP",
                "Sales Document": doc,
                "Item": item,
                "Material": mat,
                "Plant": plant,
                "Field": "Shipping Point",
                "Value": shp,
                "Error": f"Shipping Point must be equal to Plant (current Shipping Point='{shp}', Plant='{plant}')",
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


def ssa_validar_unloading_point_consistency(vbpa_df, vbak_df=None):
    """
    Header vs Item (only for Sales Document Types in
    SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP, and further restricted to Sales
    Document Type 'ZTAM' or 'ZORA'):
    - Unloading Point from header must equal Unloading Point from item.
    - Uses VBPA partner function 'SH'.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbpa_df is None or vbpa_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbpa = _ssa_load(vbpa_df)
    doc_type_map = _ssa_doc_type_map(vbak_df)
    allowed_types = {"ZTAM", "ZORA"}

    doc_col = _ssa_pick(vbpa, ["SD Document", "VBELN"])
    item_col = _ssa_pick(vbpa, ["Item (SD)", "POSNR"])
    pf_col   = _ssa_pick(vbpa, ["Partner Function", "PARVW"])
    up_col   = _ssa_pick(vbpa, ["Unloading Point"])

    if not all([doc_col, item_col, pf_col, up_col]):
        return pd.DataFrame(columns=empty_cols)

    sh = vbpa[vbpa[pf_col].astype(str).str.strip().str.upper() == "SH"].copy()
    if sh.empty:
        return pd.DataFrame(columns=empty_cols)

    for doc, grp in sh.groupby(sh[doc_col].astype(str).str.strip()):
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
            continue
        # If VBAK/doc-type info is available, also require ZTAM/ZORA
        if doc_type_map and doc_type_map.get(doc, '') not in allowed_types:
            continue

        hdr_rows = grp[grp[item_col].astype(str).str.strip().isin(['0', '', '000000', '0000000000'])]
        hdr_val = _ssa_first_nonempty(*hdr_rows[up_col].tolist()) if not hdr_rows.empty else ''

        item_rows = grp[~grp[item_col].astype(str).str.strip().isin(['0', '', '000000', '0000000000'])]
        if item_rows.empty:
            continue

        for _, row in item_rows.iterrows():
            item = _ssa_norm(row[item_col])
            item_val = _ssa_norm(row[up_col])

            if hdr_val != '' and item_val != '' and hdr_val != item_val:
                errors.append({
                    "Table": "VBPA",
                    "Sales Document": doc,
                    "Item": item,
                    "Material": "",
                    "Plant": "",
                    "Field": "Unloading Point",
                    "Value": item_val,
                    "Error": f"Unloading Point mismatch: header='{hdr_val}' vs item='{item_val}'",
                })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


def ssa_validar_customer_reference(vbap_df, vbkd_df, vbak_df=None):
    """
    Item-level Customer Reference validation using VBKD-BSTKD.

    Resolution order:
      1. Item-specific VBKD row for the same VBELN/POSNR.
      2. Header-level VBKD row where POSNR is blank or zero.

    Rejected VBAP items are skipped. The rule applies to Sales Document
    Types in SSA_ORDER_TYPES_ZLZ_GROUP (ZLZ, ZLZM), and to those in
    SSA_ORDER_TYPES_ZTAM_GROUP except ZPRO, ZOR, ZFD, ZMC, ZMD, ZDIC, ZDID
    and ZRE2.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbap_df is None or vbap_df.empty or vbkd_df is None or vbkd_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbap = _ssa_load(vbap_df)
    vbkd = _ssa_load(vbkd_df)
    doc_type_map = _ssa_doc_type_map(vbak_df)
    exclude_types = {"ZPRO", "ZOR", "ZFD", "ZMC", "ZMD", "ZDIC", "ZDID", "ZRE2"}

    ap_doc = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    ap_item = _ssa_pick(vbap, ["Sales Document Item", "Item (SD)", "POSNR"])
    ap_mat = _ssa_pick(vbap, ["Material", "MATNR"])
    ap_plant = _ssa_pick(vbap, ["Plant", "WERKS"])
    ap_status = _ssa_pick(vbap, ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"])
    ap_rej = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

    kd_doc = _ssa_pick(vbkd, ["SD Document", "Sales Document", "VBELN"])
    kd_item = _ssa_pick(vbkd, ["Item (SD)", "Sales Document Item", "POSNR"])
    kd_ref = _ssa_pick(vbkd, [
        "Customer Reference", "Customer Purchase Order Number",
        "Customer PO Number", "Cust. Reference", "BSTKD",
    ])

    if not all([ap_doc, ap_item, ap_status, kd_doc, kd_ref]):
        return pd.DataFrame(columns=empty_cols)

    def norm_item(value):
        raw = _ssa_norm(value)
        if _ssa_is_zero_or_blank(raw):
            return ""
        try:
            number = float(raw.replace(',', '.'))
            if number.is_integer():
                return str(int(number))
        except Exception:
            pass
        stripped = raw.lstrip('0')
        return stripped if stripped else ""

    item_map = {}
    header_map = {}
    for _, row in vbkd.iterrows():
        doc = _ssa_norm(row[kd_doc])
        if not doc:
            continue
        item = norm_item(row[kd_item]) if kd_item else ""
        reference = _ssa_norm(row[kd_ref])
        target = header_map if item == "" else item_map
        key = doc if item == "" else (doc, item)
        if key not in target or (target[key] == "" and reference != ""):
            target[key] = reference

    for _, row in vbap.iterrows():
        if _ssa_item_is_done(row, ap_status, ap_rej):
            continue

        doc = _ssa_norm(row[ap_doc])
        doc_type = doc_type_map.get(doc, "").upper()
        applies = (
            _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_GROUP)
            or (
                _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZTAM_GROUP)
                and doc_type not in exclude_types
            )
        )
        if not applies:
            continue

        item_display = _ssa_norm(row[ap_item])
        item_key = norm_item(row[ap_item])
        reference = _ssa_first_nonempty(
            item_map.get((doc, item_key), ""),
            header_map.get(doc, ""),
        )
        if reference == "":
            errors.append({
                "Table": "VBKD",
                "Sales Document": doc,
                "Item": item_display,
                "Material": _ssa_norm(row[ap_mat]) if ap_mat else "",
                "Plant": _ssa_norm(row[ap_plant]) if ap_plant else "",
                "Field": "Customer Reference",
                "Value": "",
                "Error": "Mandatory field blank: 'Customer Reference' (VBKD-BSTKD blank at item and header level)",
            })

    return pd.DataFrame(errors, columns=empty_cols)


def ssa_validar_weights(vbap_df, vbfa_df=None, vbep_df=None):
    """
    Item-level weight validation for document types in
    SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP.

    Weight checks:
      - Gross Weight and Net Weight must not be blank or zero.
      - Gross Weight must equal Net Weight when both values are present.

    Demand gating for ZLZ/ZLZM:
      - Match VBAP to VBEP by Sales Document (VBELN) and Item (POSNR).
      - Validate weights only when at least one matching VBEP row has
        Release Type/ABART = '1'.
      - If no matching VBEP row exists, or all matching rows carry another
        Release Type, active demand is absent and weight validation is skipped.
      - If VBEP is not supplied or its required columns are unavailable, the
        ZLZ/ZLZM weight check is skipped and the dependency is exposed in the
        Field Diagnostics output.

    OBD gating for the remaining applicable document types:
      - If usable VBFA data is supplied, validate weights only after an
        Outbound Delivery exists for the item.
      - VBFA does not add an additional gate to ZLZ/ZLZM because those types
        are gated by active demand in VBEP.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbap_df is None or vbap_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbap = _ssa_load(vbap_df)
    obd_map, gate_on_obd = _ssa_build_obd_set(vbfa_df)
    active_demand, vbep_ready = _ssa_build_active_demand_set(vbep_df)

    doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "Item (SD)", "POSNR"])
    mat_col = _ssa_pick(vbap, ["Material", "MATNR"])
    plt_col = _ssa_pick(vbap, ["Plant", "WERKS"])
    gwt_col = _ssa_pick(vbap, ["Gross Weight", "BRGEW"])
    nwt_col = _ssa_pick(vbap, ["Net Weight", "NTGEW"])
    status_col = _ssa_pick(
        vbap,
        ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"],
    )
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

    if not all([doc_col, item_col, mat_col, plt_col, gwt_col, nwt_col, status_col]):
        return pd.DataFrame(columns=empty_cols)

    for _, row in vbap.iterrows():
        if _ssa_item_is_done(row, status_col, rej_col):
            continue

        doc = _ssa_norm(row[doc_col])
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
            continue

        item = _ssa_norm(row[item_col])
        item_key = _ssa_item_key(row[item_col])
        doc_type = _SSA_ACTIVE_DOC_TYPE_MAP.get(doc, "").upper()

        if doc_type in SSA_ORDER_TYPES_ZLZ_GROUP:
            # ZLZ/ZLZM: weights are mandatory only when active demand exists.
            if not vbep_ready or (doc, item_key) not in active_demand:
                continue
        elif gate_on_obd and not _ssa_has_obd(obd_map, doc, item):
            # Other applicable types retain the existing VBFA/OBD gate.
            continue

        mat = _ssa_norm(row[mat_col])
        plant = _ssa_norm(row[plt_col])
        gwt = row[gwt_col]
        nwt = row[nwt_col]

        if _ssa_is_zero_or_blank(gwt):
            errors.append({
                "Table": "VBAP",
                "Sales Document": doc,
                "Item": item,
                "Material": mat,
                "Plant": plant,
                "Field": "Gross Weight",
                "Value": _ssa_norm(gwt),
                "Error": "Gross Weight must not be blank or zero: active demand exists in VBEP (Release Type='1')"
                    if doc_type in SSA_ORDER_TYPES_ZLZ_GROUP
                    else "Gross Weight must not be blank or zero",
            })

        if _ssa_is_zero_or_blank(nwt):
            errors.append({
                "Table": "VBAP",
                "Sales Document": doc,
                "Item": item,
                "Material": mat,
                "Plant": plant,
                "Field": "Net Weight",
                "Value": _ssa_norm(nwt),
                "Error": "Net Weight must not be blank or zero: active demand exists in VBEP (Release Type='1')"
                    if doc_type in SSA_ORDER_TYPES_ZLZ_GROUP
                    else "Net Weight must not be blank or zero",
            })

        if (
            not _ssa_is_zero_or_blank(gwt)
            and not _ssa_is_zero_or_blank(nwt)
            and not _ssa_float_equal(gwt, nwt)
        ):
            errors.append({
                "Table": "VBAP",
                "Sales Document": doc,
                "Item": item,
                "Material": mat,
                "Plant": plant,
                "Field": "Gross Weight / Net Weight",
                "Value": f"Gross={_ssa_norm(gwt)} | Net={_ssa_norm(nwt)}",
                "Error": "Gross Weight and Net Weight must have the same value",
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


def ssa_validar_incoterms_ddp_gb(vbak_df, vbkd_df, kna1_df, vbpa_df=None, vbap_df=None):
    """
    New check — Incoterms (VBKD) must not be blank, for Sales Document
    Types in SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP.

    Existing check — if Incoterms (VBKD) = 'DDP' AND Customer Country = 'GB':
      - VBAK Alt.Tax Classific. must be '1'
      - VBAK Tax Departure C/R  must be 'GB'

    Country resolved via:
      Primary:  VBAK~Sold-to Party → KNA1~Customer → KNA1~Country/Region Key
      Optional: VBPA (Partner Function = 'SP') → Country/Region Key

    Exclude Rejected Documents — if vbap_df is supplied, Sales Documents
    where EVERY VBAP item is 'done' (Overall Status = C, or Reason for
    Rejection/VBAP-ABGRU filled) are skipped entirely (neither the
    Incoterms-blank check nor the DDP+GB check apply).
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant",
                  "Field", "Value", "Error"]

    if any(df is None or df.empty for df in [vbak_df, vbkd_df, kna1_df]):
        return pd.DataFrame(columns=empty_cols)

    vbak = _ssa_load(vbak_df)
    vbkd = _ssa_load(vbkd_df)
    kna1 = _ssa_load(kna1_df)
    fully_rejected_docs = _ssa_fully_rejected_docs(vbap_df)

    # Column picks
    vbak_doc  = _ssa_pick(vbak, ["Sales Document", "VBELN"])
    vbak_stp  = _ssa_pick(vbak, ["Sold-to Party", "KUNNR"])
    vbak_atc  = _ssa_pick(vbak, ["Alt.Tax Classific.", "TAXK1"])
    vbak_tdc  = _ssa_pick(vbak, ["Tax Departure C/R", "STCEG_L"])
    vbkd_doc  = _ssa_pick(vbkd, ["SD Document", "VBELN"])
    vbkd_inc  = _ssa_pick(vbkd, ["Incoterms", "INCO1"])
    kna1_cust = _ssa_pick(kna1, ["Customer", "KUNNR"])
    kna1_ctry = _ssa_pick(kna1, ["Country/Region Key", "LAND1"])

    if not all([vbak_doc, vbak_stp, vbak_atc, vbak_tdc,
                vbkd_doc, vbkd_inc, kna1_cust, kna1_ctry]):
        return pd.DataFrame(columns=empty_cols)

    # Build KNA1 map: customer -> country
    kna1_map = dict(zip(
        kna1[kna1_cust].astype(str).str.strip(),
        kna1[kna1_ctry].astype(str).str.strip().str.upper()
    ))

    # Optional: VBPA country override (Partner Function = SP)
    vbpa_map = {}
    if vbpa_df is not None and not vbpa_df.empty:
        vbpa = _ssa_load(vbpa_df)
        vp_doc = _ssa_pick(vbpa, ["SD Document", "VBELN"])
        vp_pf  = _ssa_pick(vbpa, ["Partner Function", "PARVW"])
        vp_cty = _ssa_pick(vbpa, ["Country/Region Key", "LAND1"])
        if all([vp_doc, vp_pf, vp_cty]):
            sp_rows = vbpa[vbpa[vp_pf].astype(str).str.strip().str.upper() == 'SP']
            for _, r in sp_rows.iterrows():
                vbpa_map[_ssa_norm(r[vp_doc])] = _ssa_norm(r[vp_cty]).upper()

    # Build VBKD map: sales doc -> incoterms (raw, for the blank check)
    vbkd_inc_map = {}
    for _, r in vbkd.iterrows():
        doc = _ssa_norm(r[vbkd_doc])
        inc = _ssa_norm(r[vbkd_inc])
        vbkd_inc_map[doc] = inc

    # Incoterms must not be blank, for Sales Document Types in
    # SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP.
    for _, row in vbak.iterrows():
        doc = _ssa_norm(row[vbak_doc])
        if doc in fully_rejected_docs:
            continue
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
            continue
        inc_val = vbkd_inc_map.get(doc, '')
        if inc_val == '':
            errors.append({
                "Table": "VBKD", "Sales Document": doc, "Item": '', "Material": '', "Plant": '',
                "Field": "Incoterms",
                "Value": '',
                "Error": "Mandatory field blank: 'Incoterms'",
            })

    # Evaluate each VBAK row — DDP + GB
    for _, row in vbak.iterrows():
        doc      = _ssa_norm(row[vbak_doc])
        if doc in fully_rejected_docs:
            continue
        stp      = _ssa_norm(row[vbak_stp])
        incoterm = vbkd_inc_map.get(doc, '').upper()
        if incoterm != 'DDP':
            continue

        # Resolve country — VBPA override first, then KNA1
        country = vbpa_map.get(doc) or kna1_map.get(stp, '')
        if country != 'GB':
            continue

        # DDP + GB → check Alt.Tax Classific.
        atc_val = _ssa_norm(row[vbak_atc])
        if atc_val != '1':
            errors.append({
                "Table": "VBAK", "Sales Document": doc, "Item": '', "Material": '', "Plant": '',
                "Field": "Alt.Tax Classific.",
                "Value": atc_val,
                "Error": (f"Incoterms=DDP and Country=GB: "
                          f"Alt.Tax Classific. must be '1' (current: '{atc_val}')"),
            })

        # DDP + GB → check Tax Departure C/R
        tdc_val = _ssa_norm(row[vbak_tdc])
        if tdc_val.upper() != 'GB':
            errors.append({
                "Table": "VBAK", "Sales Document": doc, "Item": '', "Material": '', "Plant": '',
                "Field": "Tax Departure C/R",
                "Value": tdc_val,
                "Error": (f"Incoterms=DDP and Country=GB: "
                          f"Tax Departure C/R must be 'GB' (current: '{tdc_val}')"),
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


def ssa_validar_incompletion_log(vbak_df, vbap_df):
    """
    Incompletion Log checks (rejected items are skipped):

    A) VBAK Overall Header (FA) must not be 'A'
    B) VBAP Item (JC) must not be 'A'

    Exception:
    - VBAP items belonging to order type ZTAM are ignored when
      Order Quantity (KWMENG) is zero.
    """
    errors = []
    completed_docs = _ssa_fully_rejected_docs(vbap_df)

    # Store Sales Documents whose order type is ZTAM.
    ztam_docs = set()

    # ── VBAK header check ─────────────────────────────────────────────────────
    if vbak_df is not None and not vbak_df.empty:
        vbak = _ssa_load(vbak_df)

        doc_col = _ssa_pick(
            vbak,
            [
                "Sales Document",
                "VBELN",
            ],
        )

        order_type_col = _ssa_pick(
            vbak,
            [
                "Sales Document Type",
                "Order Type",
                "Document Type",
                "AUART",
            ],
        )

        oh_col = _ssa_pick(
            vbak,
            [
                "Overall Header",
                "UVALL",
            ],
        )

        # Identify all ZTAM Sales Documents.
        if doc_col and order_type_col:
            for _, row in vbak.iterrows():
                doc = _ssa_norm(row[doc_col])
                order_type = _ssa_norm(row[order_type_col]).upper()

                if order_type == "ZTAM":
                    ztam_docs.add(doc)

        if oh_col:
            for _, row in vbak.iterrows():
                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                if doc in completed_docs:
                    continue

                val = _ssa_norm(row[oh_col]).upper()

                if val == "A":
                    errors.append({
                        "Table": "VBAK",
                        "Sales Document": doc,
                        "Item": "",
                        "Material": "",
                        "Plant": "",
                        "Field": "Overall Header",
                        "Value": val,
                        "Error": (
                            "Incompletion Log: Overall Header = 'A' "
                            "(document is incomplete)"
                        ),
                    })

    # ── VBAP item check (rejected items are skipped) ─────────────────────────
    if vbap_df is not None and not vbap_df.empty:
        vbap = _ssa_load(vbap_df)

        doc_col = _ssa_pick(
            vbap,
            [
                "Sales Document",
                "VBELN",
            ],
        )

        item_col = _ssa_pick(
            vbap,
            [
                "Sales Document Item",
                "POSNR",
            ],
        )

        mat_col = _ssa_pick(
            vbap,
            [
                "Material",
                "MATNR",
            ],
        )

        plt_col = _ssa_pick(
            vbap,
            [
                "Plant",
                "WERKS",
            ],
        )

        qty_col = _ssa_pick(
            vbap,
            [
                "Order Quantity",
                "KWMENG",
            ],
        )

        status_col = _ssa_pick(
            vbap,
            [
                "Overall Status",
                "Overall SD Status",
                "Overall Processing Status",
                "GBSTA",
                "GBSTK",
            ],
        )

        rej_col = _ssa_pick(
            vbap,
            [
                "Reason for Rejection",
                "ABGRU",
            ],
        )

        itm_col = _ssa_pick(
            vbap,
            [
                "UVPOS",
                "Item Incompletion Status",
                "Item Status",
            ],
        )

        if itm_col and status_col:
            for _, row in vbap.iterrows():
                doc = _ssa_norm(row[doc_col]) if doc_col else ""

                # Completed or rejected items are skipped.
                if _ssa_item_is_done(row, status_col, rej_col):
                    continue

                # Exception:
                # Ignore ZTAM items when Order Quantity (KWMENG) is zero.
                if doc in ztam_docs and qty_col:
                    raw_qty = row[qty_col]

                    try:
                        if pd.isna(raw_qty):
                            qty = None
                        elif isinstance(raw_qty, str):
                            qty_text = raw_qty.strip().replace(" ", "")

                            # Supports values such as:
                            # 0
                            # 0.00
                            # 0,00
                            # 1,000.00
                            # 1.000,00
                            if "," in qty_text and "." in qty_text:
                                if qty_text.rfind(",") > qty_text.rfind("."):
                                    qty_text = (
                                        qty_text
                                        .replace(".", "")
                                        .replace(",", ".")
                                    )
                                else:
                                    qty_text = qty_text.replace(",", "")
                            elif "," in qty_text:
                                qty_text = qty_text.replace(",", ".")

                            qty = float(qty_text)
                        else:
                            qty = float(raw_qty)

                    except (TypeError, ValueError):
                        qty = None

                    if qty == 0:
                        continue

                val = _ssa_norm(row[itm_col]).upper()

                if val != "A":
                    continue

                errors.append({
                    "Table": "VBAP",
                    "Sales Document": doc,
                    "Item": (
                        _ssa_norm(row[item_col])
                        if item_col else ""
                    ),
                    "Material": (
                        _ssa_norm(row[mat_col])
                        if mat_col else ""
                    ),
                    "Plant": (
                        _ssa_norm(row[plt_col])
                        if plt_col else ""
                    ),
                    "Field": "Item",
                    "Value": val,
                    "Error": (
                        "Incompletion Log: Item status = 'A' "
                        "(item is incomplete)"
                    ),
                })

    return pd.DataFrame(errors)


def ssa_validar_rejected_material_status(vbak_df, vbap_df, marc_df):
    """
    Report status 70/75/99 items whose VBAP Overall Status is not C.
    Items that are complete (Overall Status = 'C') or rejected
    (VBAP-ABGRU filled) are skipped.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]
    if vbap_df is None or vbap_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=empty_cols)
    vbap = _ssa_load(vbap_df)
    marc = _ssa_load(marc_df)
    doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "Item (SD)", "POSNR"])
    mat_col = _ssa_pick(vbap, ["Material", "MATNR"])
    plt_col = _ssa_pick(vbap, ["Plant", "WERKS"])
    overall_status_col = _ssa_pick(vbap, ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"])
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])
    marc_mat = _ssa_pick(marc, ["Material", "MATNR"])
    marc_plt = _ssa_pick(marc, ["Plant", "WERKS"])
    marc_psms = _ssa_pick(marc, ["Plant-Specific Material Status", "Plant Specific Material Status", "MTSTB"])
    if not all([doc_col, mat_col, plt_col, overall_status_col, marc_mat, marc_plt, marc_psms]):
        return pd.DataFrame(columns=empty_cols)
    def norm_psms(value):
        value = _ssa_norm(value)
        if not value:
            return ""
        try:
            return str(int(float(value.replace(",", "."))))
        except (TypeError, ValueError):
            return value.upper()
    marc_map = {}
    for _, r in marc.iterrows():
        marc_map[(_ssa_norm(r[marc_mat]), _ssa_norm(r[marc_plt]).upper())] = norm_psms(r[marc_psms])
    for _, row in vbap.iterrows():
        doc = _ssa_norm(row[doc_col])
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP):
            continue
        if _ssa_item_is_done(row, overall_status_col, rej_col):
            continue
        overall_status = _ssa_norm(row[overall_status_col]).upper()
        material = _ssa_norm(row[mat_col])
        plant = _ssa_norm(row[plt_col]).upper()
        material_status = marc_map.get((material, plant), "")
        if material_status not in {"70", "75", "99"}:
            continue
        errors.append({
            "Table": "VBAP", "Sales Document": doc,
            "Item": _ssa_norm(row[item_col]) if item_col else "",
            "Material": material, "Plant": plant, "Field": "Overall Status",
            "Value": overall_status,
            "Error": (f"Order is active: Material Status {material_status} found for "
                      f"Material={material}, Plant={plant}, but Overall Status is "
                      f"'{overall_status or '[blank]'}' instead of 'C'"),
        })
    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


def ssa_validar_order_without_items(vbak_df, vbap_df):
    """
    Flags any Sales Document (VBAK) that has no corresponding line items in VBAP.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbak_df is None or vbak_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbak = _ssa_load(vbak_df)
    vbak_doc_col = _ssa_pick(vbak, ["Sales Document", "VBELN"])
    if vbak_doc_col is None:
        return pd.DataFrame(columns=empty_cols)

    vbap_docs = set()
    if vbap_df is not None and not vbap_df.empty:
        vbap = _ssa_load(vbap_df)
        vbap_doc_col = _ssa_pick(vbap, ["Sales Document", "VBELN"])
        if vbap_doc_col:
            vbap_docs = set(vbap[vbap_doc_col].astype(str).str.strip())

    for _, row in vbak.iterrows():
        doc = _ssa_norm(row[vbak_doc_col])
        if doc == '':
            continue
        if doc not in vbap_docs:
            errors.append({
                "Table": "VBAK", "Sales Document": doc, "Item": '', "Material": '', "Plant": '',
                "Field": "Items",
                "Value": '',
                "Error": "Sales Document has no line items in VBAP",
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


def ssa_validar_customer_material(vbap_df):
    """
    Item-level:
    - Customer Material Number (KDMAT) must not be empty (rejected items are skipped).
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbap_df is None or vbap_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbap = _ssa_load(vbap_df)

    doc_col  = _ssa_pick(vbap, ["Sales Document", "VBELN"])
    item_col = _ssa_pick(vbap, ["Sales Document Item", "POSNR"])
    mat_col  = _ssa_pick(vbap, ["Material", "MATNR"])
    plt_col  = _ssa_pick(vbap, ["Plant", "WERKS"])
    cmat_col = _ssa_pick(vbap, ["Customer Material Number", "Material by Customer", "KDMAT"])
    status_col = _ssa_pick(vbap, ["Overall Status", "Overall SD Status", "Overall Processing Status", "GBSTA", "GBSTK"])
    rej_col = _ssa_pick(vbap, ["Reason for Rejection", "ABGRU"])

    if not all([doc_col, item_col, mat_col, plt_col, cmat_col, status_col]):
        return pd.DataFrame(columns=empty_cols)

    for _, row in vbap.iterrows():
        if _ssa_item_is_done(row, status_col, rej_col):
            continue

        if _ssa_norm(row[cmat_col]) == '':
            errors.append({
                "Table": "VBAP",
                "Sales Document": _ssa_norm(row[doc_col]),
                "Item": _ssa_norm(row[item_col]),
                "Material": _ssa_norm(row[mat_col]),
                "Plant": _ssa_norm(row[plt_col]),
                "Field": "Customer Material Number",
                "Value": "",
                "Error": "Mandatory field blank: 'Customer Material Number'",
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)


def ssa_validar_mrp_for_ds(vbak_df):
    """
    Header-level (only for Sales Document Types in SSA_ORDER_TYPES_ZLZ_GROUP):
    - VBAK~ABDIS ("MRP for DlvSchType") must not be blank.
    """
    errors = []
    empty_cols = ["Table", "Sales Document", "Item", "Material", "Plant", "Field", "Value", "Error"]

    if vbak_df is None or vbak_df.empty:
        return pd.DataFrame(columns=empty_cols)

    vbak = _ssa_load(vbak_df)
    doc_col   = _ssa_pick(vbak, ["Sales Document", "VBELN"])
    abdis_col = _ssa_pick(vbak, ["MRP for DlvSchType", "MRPforDS", "ABDIS"])

    if not all([doc_col, abdis_col]):
        return pd.DataFrame(columns=empty_cols)

    for _, row in vbak.iterrows():
        doc = _ssa_norm(row[doc_col])
        if not _ssa_doc_type_in(doc, SSA_ORDER_TYPES_ZLZ_GROUP):
            continue
        if _ssa_norm(row[abdis_col]) == '':
            errors.append({
                "Table": "VBAK", "Sales Document": doc, "Item": '', "Material": '', "Plant": '',
                "Field": "MRP for DlvSchType",
                "Value": "",
                "Error": "Mandatory field blank: 'MRP for DlvSchType' (ABDIS)",
            })

    return pd.DataFrame(errors, columns=empty_cols) if not errors else pd.DataFrame(errors)
