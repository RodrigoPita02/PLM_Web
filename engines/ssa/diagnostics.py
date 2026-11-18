"""Diagnósticos de input do SSA."""

import pandas as pd

from engines.ssa.constants import SSA_DIAGNOSTIC_COLUMNS
from engines.ssa.helpers import _ssa_load, _ssa_pick

def _ssa_diag_check(df, rule, check, table, fields, source_required=True):
    """
    Inspect one table dependency and return one diagnostic row per logical field.

    Status values:
      - OK: field exists and the corresponding check can use it.
      - MISSING: required field does not exist; the check is skipped/partial.
      - SOURCE MISSING: required source table was not supplied or is empty.
      - OPTIONAL SOURCE MISSING: optional source was not supplied; fallback applies.
    """
    rows = []
    source_available = df is not None and not df.empty
    work = _ssa_load(df) if source_available else None

    if not source_available:
        status = "SOURCE MISSING" if source_required else "OPTIONAL SOURCE MISSING"
        for logical_field, definition in fields.items():
            candidates = definition.get("candidates", [logical_field])
            required = definition.get("required", True)
            impact = definition.get("impact", "Validation cannot use this field.")
            if not required and source_required:
                status_for_field = "OPTIONAL SOURCE MISSING"
            else:
                status_for_field = status
            rows.append({
                "Rule": rule,
                "Check": check,
                "Table": table,
                "Logical Field": logical_field,
                "Accepted Column Names": ", ".join(candidates),
                "Status": status_for_field,
                "Impact": impact,
            })
        return rows

    for logical_field, definition in fields.items():
        candidates = definition.get("candidates", [logical_field])
        required = definition.get("required", True)
        column = _ssa_pick(work, candidates)
        if column is not None:
            status = "OK"
            impact = f"Using source column '{column}'."
        elif required:
            status = "MISSING"
            impact = definition.get("impact", "This validation check is skipped.")
        else:
            status = "OPTIONAL MISSING"
            impact = definition.get("impact", "Optional behaviour is unavailable.")
        rows.append({
            "Rule": rule,
            "Check": check,
            "Table": table,
            "Logical Field": logical_field,
            "Accepted Column Names": ", ".join(candidates),
            "Status": status,
            "Impact": impact,
        })
    return rows


def _ssa_run_input_diagnostics(vbak, vbap, vbkd, kna1, marc=None, vbpa=None, vbfa=None, vbep=None, mvke=None):
    """Validate all table/field dependencies used by the current SSA engine."""
    diagnostics = []

    def add(df, rule, check, table, fields, source_required=True):
        diagnostics.extend(_ssa_diag_check(
            df, rule, check, table, fields, source_required=source_required
        ))

    # Shared applicability and active-document logic.
    add(vbak, "Engine Setup", "Document type mapping", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Type": {"candidates": ["Sales Document Type", "Order Type", "AUART"],
                                "impact": "Rules based on AUART cannot determine applicability."},
    })
    add(vbap, "Engine Setup", "Closed-document and item scope", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"],
                                 "impact": "Closed/rejected documents and items cannot be excluded correctly."},
        "Material": {"candidates": ["Material", "MATNR"],
                     "impact": "S-* material scoping cannot be applied."},
        "Plant": {"candidates": ["Plant", "WERKS"], "required": False,
                  "impact": "Plant filter cannot be applied."},
    })

    # Mandatory Fields, split by independently executable sub-check.
    for field, groups in [
        ("Customer Group 1", "ZLZ/ZLZM, ZTAM/ZOR/..., ZDIC/ZMC/..., ZMD/ZDSD/... (all order types)"),
        ("Customer Group 2", "ZLZ/ZLZM and ZTAM/ZOR/..."),
        ("Customer Group 3", "ZLZ/ZLZM"),
    ]:
        fields = {
            "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
            "Overall Header": {
                "candidates": ["Overall Header", "UVALL"],
                "required": False,
                "impact": "Completed VBAK documents with Overall Header='C' cannot be skipped.",
            },
            field: {"candidates": [field], "impact": f"Mandatory check for {field} ({groups}) is skipped."},
        }
        if field == "Customer Group 3":
            fields["MRP for DlvSchType"] = {"candidates": ["MRP for DlvSchType", "MRPforDS", "ABDIS"], "required": False,
                                    "impact": "ZLZ exception ABDIS=E cannot be identified; blank ABDIS is treated as not E."}
        add(vbak, "Mandatory Fields", field, "VBAK", fields)

    mandatory_vbap_checks = [
        ("Route", ["Route", "ROUTE"]),
        ("IFS ID", ["IFS ID"]),
        ("Supplier_code", ["Supplier_code"]),
        ("Plant", ["Plant", "WERKS"]),
        ("Acct Assmt Grp Mat.", ["Acct Assmt Grp Mat.", "Account Assignment Group", "Acct Assignment Grp", "Acc.Assignment Grp.", "KTGRM"]),
    ]
    for logical, candidates in mandatory_vbap_checks:
        add(vbap, "Mandatory Fields", logical, "VBAP", {
            "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
            "Sales Document Item": {"candidates": ["Sales Document Item", "Item (SD)", "POSNR"], "required": False,
                                    "impact": "Item identifier will be blank in reported errors."},
            "Material": {"candidates": ["Material", "MATNR"], "required": logical == "Acct Assmt Grp Mat.",
                         "impact": "Material-dependent KTGRM logic cannot run." if logical == "Acct Assmt Grp Mat." else "Material will be blank in reported errors."},
            "Plant": {"candidates": ["Plant", "WERKS"], "required": logical == "Plant",
                      "impact": "Plant mandatory check is skipped." if logical == "Plant" else "Plant will be blank in reported errors."},
            "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                     "impact": "Rejected items cannot be excluded from this check."},
            logical: {"candidates": candidates, "impact": f"Mandatory check for {logical} is skipped."},
        })

    add(vbak, "Terms of Payment", "Applicable Sales Documents", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "Sales Doc.", "VBELN"]},
    })
    add(vbkd, "Terms of Payment", "Header value", "VBKD", {
        "Sales Document": {"candidates": ["SD Document", "Sales Document", "Sales Doc.", "VBELN"]},
        "Item": {"candidates": ["Item (SD)", "Sales Document Item", "POSNR", "Item"], "required": False,
                 "impact": "Cannot restrict the lookup to header POSNR; first document value may be used."},
        "Terms of Payment": {"candidates": ["Terms of Payment", "Terms of Payment Key", "ZTERM"],
                             "impact": "Terms of Payment check is skipped."},
    })

    add(vbap, "Shipping Point", "Mandatory and equal to Plant", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Item": {"candidates": ["Sales Document Item", "POSNR"]},
        "Material": {"candidates": ["Material", "MATNR"]},
        "Plant": {"candidates": ["Plant", "WERKS"]},
        "Shipping Point": {"candidates": ["Shipping Point/Receiving Pt", "Shipping Point", "VSTEL"]},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                 "impact": "Rejected items cannot be excluded."},
    })

    add(vbpa, "Unloading Point Consistency", "Header versus item", "VBPA", {
        "Sales Document": {"candidates": ["SD Document", "VBELN"]},
        "Item": {"candidates": ["Item (SD)", "POSNR"]},
        "Partner Function": {"candidates": ["Partner Function", "PARVW"]},
        "Unloading Point": {"candidates": ["Unloading Point"]},
    }, source_required=True)

    add(vbkd, "Customer Reference", "VBKD-BSTKD item/header fallback", "VBKD", {
        "Sales Document": {"candidates": ["SD Document", "Sales Document", "VBELN"]},
        "Item": {"candidates": ["Item (SD)", "Sales Document Item", "POSNR"], "required": False,
                 "impact": "Item-specific lookup is unavailable; document-level fallback only."},
        "Customer Reference": {"candidates": ["Customer Reference", "Customer Purchase Order Number", "Customer PO Number", "Cust. Reference", "BSTKD"],
                               "impact": "Customer Reference validation is skipped."},
    })
    add(vbap, "Customer Reference", "Items to validate", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Item": {"candidates": ["Sales Document Item", "Item (SD)", "POSNR"]},
        "Material": {"candidates": ["Material", "MATNR"], "required": False, "impact": "Material is blank in errors."},
        "Plant": {"candidates": ["Plant", "WERKS"], "required": False, "impact": "Plant is blank in errors."},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                 "impact": "Rejected items cannot be excluded."},
    })

    add(vbap, "Weights", "Gross/Net values", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Item": {"candidates": ["Sales Document Item", "POSNR"]},
        "Material": {"candidates": ["Material", "MATNR"]},
        "Plant": {"candidates": ["Plant", "WERKS"]},
        "Gross Weight": {"candidates": ["Gross Weight"]},
        "Net Weight": {"candidates": ["Net Weight"]},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                 "impact": "Rejected items cannot be excluded."},
    })
    add(vbfa, "Weights", "OBD gating", "VBFA", {
        "Preceding Document": {"candidates": ["Preceding Document", "VBELV"]},
        "Preceding Item": {"candidates": ["Item (SD) Preceding Doc.", "Preceding Item", "POSNV"], "required": False,
                           "impact": "OBD gating can only be applied at document level."},
        "Subsequent Document": {"candidates": ["Subsequent Document", "VBELN"]},
        "Preceding Document Category": {
            "candidates": ["Subsequent Doc. Cat.", "Preceding Document Category", "VBTYP_N"],
            "impact": "OBD gating is skipped because the Subsequent Document cannot be confirmed as a delivery (category J).",
        },
    }, source_required=False)

    add(vbep, "Weights", "Active demand gating", "VBEP", {
        "Sales Document": {"candidates": ["Sales Document", "SD Document", "VBELN"],
                           "impact": "ZLZ/ZLZM active demand cannot be matched to VBAP."},
        "Sales Document Item": {"candidates": ["Sales Document Item", "Item (SD)", "POSNR"],
                                "impact": "ZLZ/ZLZM active demand cannot be matched at item level."},
        "Release Type": {"candidates": ["Release Type", "ABART"],
                         "impact": "ZLZ/ZLZM active demand cannot be identified; weight checks are skipped for those document types."},
    }, source_required=False)
    add(vbak, "Incoterms", "DDP + GB header checks", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sold-to Party": {"candidates": ["Sold-to Party", "KUNNR"]},
        "Alt.Tax Classific.": {"candidates": ["Alt.Tax Classific.", "TAXK1"]},
        "Tax Departure C/R": {"candidates": ["Tax Departure C/R", "STCEG_L"]},
    })
    add(vbkd, "Incoterms", "Mandatory Incoterms", "VBKD", {
        "Sales Document": {"candidates": ["SD Document", "VBELN"]},
        "Incoterms": {"candidates": ["Incoterms", "INCO1"]},
    })
    add(kna1, "Incoterms", "Customer country", "KNA1", {
        "Customer": {"candidates": ["Customer", "KUNNR"]},
        "Country/Region Key": {"candidates": ["Country/Region Key", "LAND1"]},
    })
    add(vbpa, "Incoterms", "Optional SP country override", "VBPA", {
        "Sales Document": {"candidates": ["SD Document", "VBELN"]},
        "Partner Function": {"candidates": ["Partner Function", "PARVW"]},
        "Country/Region Key": {"candidates": ["Country/Region Key", "LAND1"]},
    }, source_required=False)

    add(vbak, "Incompletion Log", "Overall Header", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"], "required": False,
                           "impact": "Document identifier is blank in errors."},
        "Overall Header": {"candidates": ["Overall Header", "UVALL"], "impact": "Header incompletion check is skipped."},
    })
    add(vbap, "Incompletion Log", "Item status", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"], "required": False},
        "Sales Document Item": {"candidates": ["Sales Document Item", "POSNR"], "required": False},
        "Material": {"candidates": ["Material", "MATNR"], "required": False},
        "Plant": {"candidates": ["Plant", "WERKS"], "required": False},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                 "impact": "Rejected items cannot be excluded."},
        "Item Incompletion Status": {"candidates": ["UVPOS", "Item Incompletion Status", "Item Status"],
                                     "impact": "Item incompletion check is skipped."},
    })

    add(vbap, "Rejected + Status 70-75-99", "Open item identification", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Item": {"candidates": ["Sales Document Item", "POSNR"], "required": False},
        "Material": {"candidates": ["Material", "MATNR"]},
        "Plant": {"candidates": ["Plant", "WERKS"]},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"]},
    })
    add(marc, "Rejected + Status 70-75-99", "Plant material status", "MARC", {
        "Material": {"candidates": ["Material", "MATNR"]},
        "Plant": {"candidates": ["Plant", "WERKS"]},
        "Plant-Specific Material Status": {"candidates": ["Plant-Specific Material Status", "Plant Specific Material Status", "MTSTB"]},
    }, source_required=True)

    add(mvke, "Mandatory Fields", "S-material Acct Assmt Grp Mat. lookup", "MVKE", {
        "Material": {"candidates": ["Material", "MATNR"]},
        "Sales Organization": {"candidates": ["Sales Organization", "Sales Org", "VKORG"]},
        "Distribution Channel": {"candidates": ["Distribution Channel", "Distr. Channel", "VTWEG"]},
        "Acct Assmt Grp Mat.": {"candidates": ["Acct Assmt Grp Mat.", "Account Assignment Group", "KTGRM"],
                                "impact": "S-materials in SSA_S_MATERIAL_LIST fall back to the plain non-blank check instead of the MVKE comparison."},
    }, source_required=False)

    add(vbak, "Order Without Items", "Header documents", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
    })
    add(vbap, "Order Without Items", "Item documents", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
    })

    add(vbap, "Customer Material", "Mandatory KDMAT", "VBAP", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "Sales Document Item": {"candidates": ["Sales Document Item", "POSNR"]},
        "Material": {"candidates": ["Material", "MATNR"]},
        "Plant": {"candidates": ["Plant", "WERKS"]},
        "Customer Material Number": {"candidates": ["Customer Material Number", "Material by Customer", "KDMAT"]},
        "Reason for Rejection": {"candidates": ["Reason for Rejection", "ABGRU"], "required": False,
                                 "impact": "Rejected items cannot be excluded."},
    })

    add(vbak, "MRP for DS Type", "Mandatory ABDIS", "VBAK", {
        "Sales Document": {"candidates": ["Sales Document", "VBELN"]},
        "MRP for DlvSchType": {"candidates": ["MRP for DlvSchType", "MRPforDS", "ABDIS"]},
    })

    return pd.DataFrame(diagnostics, columns=SSA_DIAGNOSTIC_COLUMNS)


def _ssa_print_input_diagnostics(diagnostics_df):
    """Print a concise execution/readiness report for every SSA check."""
    line = "=" * 110
    print("\n" + line)
    print("SSA INPUT FIELD DIAGNOSTICS")
    print(line)

    if diagnostics_df is None or diagnostics_df.empty:
        print("[SSA][WARNING] No diagnostics were produced.")
        print(line + "\n")
        return

    problem_statuses = {"MISSING", "SOURCE MISSING", "OPTIONAL MISSING", "OPTIONAL SOURCE MISSING"}
    problems = diagnostics_df[diagnostics_df["Status"].isin(problem_statuses)]

    grouped = diagnostics_df.groupby(["Rule", "Check"], sort=False)
    for (rule, check), group in grouped:
        required_problem = group["Status"].isin({"MISSING", "SOURCE MISSING"}).any()
        optional_problem = group["Status"].isin({"OPTIONAL MISSING", "OPTIONAL SOURCE MISSING"}).any()
        state = "SKIPPED/PARTIAL" if required_problem else ("RUN WITH FALLBACK" if optional_problem else "READY")
        print(f"[SSA][{state}] {rule} | {check}")

    if problems.empty:
        print("\n[SSA][OK] All configured source fields were found.")
    else:
        print("\nMissing or unavailable dependencies:")
        for _, row in problems.iterrows():
            print(
                f" - [{row['Status']}] {row['Rule']} | {row['Check']} | "
                f"{row['Table']}.{row['Logical Field']}"
            )
            print(f"   Accepted: {row['Accepted Column Names']}")
            print(f"   Impact: {row['Impact']}")

    print(line + "\n")
