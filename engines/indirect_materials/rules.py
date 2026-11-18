"""Regras de validação do Indirect Materials."""

import re

from engines.common.helpers import normalize_code
from engines.indirect_materials.constants import IM_MATERIAL_TYPES
from engines.indirect_materials.helpers import (
    _append_error,
    _is_blank,
    _is_x,
    _material_key,
    _norm,
    _pick,
    _to_float,
)

def im_validar_basic(sap_df, errors):
    """Basic / General Data fields (MARA + MAKT)."""
    mat_col      = _pick(sap_df, ["Material", "MATNR"])
    plant_col    = _pick(sap_df, ["Plant", "WERKS"])
    mt_col       = _pick(sap_df, ["Material Type", "MTART"])
    desc_col     = _pick(sap_df, ["Description", "Material Description", "MAKTX"])
    base_uom_col = _pick(sap_df, ["Base UOM", "Base Unit of Measure", "MEINS"])
    div_col      = _pick(sap_df, ["Division", "SPART"])
    mg_col       = _pick(sap_df, ["Material Group", "MATKL"])
    gicg_col     = _pick(sap_df, ["General Item Category Group", "Item Category Group", "MTPOS"])
    gross_col    = _pick(sap_df, ["Gross Weight", "BRGEW"])
    net_col      = _pick(sap_df, ["Net Weight", "NTGEW"])
    uow_col      = _pick(sap_df, ["Unit of Weight (ISO Format)", "Unit of Weight", "GEWEI"])
    pack_grp_col = _pick(sap_df, ["Matl Grp Pack.Matls", "Material Group: Packaging Materials", "MAGRV"])

    for idx, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = row[plant_col] if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # MATNR length — MARA
        if _is_blank(material):
            _append_error(
                errors, material=material, plant=plant, field="Material",
                value="", error="Mandatory field blank", table="MARA",
            )
        elif len(_norm(material)) > 40:
            _append_error(
                errors, material=material, plant=plant, field="Material",
                value=material, error="Material number exceeds 40 characters", table="MARA",
            )

        # MAKTX — MAKT
        if desc_col:
            desc = _norm(row[desc_col])
            if _is_blank(desc):
                _append_error(
                    errors, material=material, plant=plant, field="Description",
                    value="", error="Mandatory field blank", table="MAKT",
                )
            elif len(desc) > 40:
                _append_error(
                    errors, material=material, plant=plant, field="Description",
                    value=desc, error=f"Description exceeds 40 characters (len={len(desc)})",
                    table="MAKT",
                )

        # MEINS — MARA
        if base_uom_col and _is_blank(row[base_uom_col]):
            _append_error(
                errors, material=material, plant=plant, field="Base UOM",
                value="", error="Mandatory field blank", table="MARA",
            )

        # SPART — MARA
        if div_col:
            val = normalize_code(row[div_col], as_int=True, zero_pad=2)
            if val == "":
                _append_error(
                    errors, material=material, plant=plant, field="Division",
                    value="", error="Mandatory field blank", table="MARA",
                )
            elif val != "00":
                _append_error(
                    errors, material=material, plant=plant, field="Division",
                    value=row[div_col], error="Should be '00'", table="MARA",
                )

        # MATKL — MARA
        if mg_col:
            mg = _norm(row[mg_col]).upper()
            if _is_blank(mg):
                _append_error(
                    errors, material=material, plant=plant, field="Material Group",
                    value="", error="Mandatory field blank", table="MARA",
                )
            elif mt in {"HIBE", "ERSA"} and not re.fullmatch(r"MT\d{7}", mg):
                _append_error(
                    errors, material=material, plant=plant, field="Material Group",
                    value=mg, error="HIBE/ERSA must match pattern MTXXXXXXX (MT + 7 digits)",
                    table="MARA",
                )
            elif mt == "VERP" and not re.fullmatch(r"LF\d{7}", mg):
                _append_error(
                    errors, material=material, plant=plant, field="Material Group",
                    value=mg, error="VERP must match pattern LFXXXXXXX (LF + 7 digits)",
                    table="MARA",
                )

        # MTPOS — MARA  (only ERSA/HIBE)
        if gicg_col and mt in {"ERSA", "HIBE"}:
            val = _norm(row[gicg_col]).upper()
            if _is_blank(val):
                _append_error(
                    errors, material=material, plant=plant,
                    field="General Item Category Group",
                    value="", error="Mandatory field blank", table="MARA",
                )
            elif val != "NORM":
                _append_error(
                    errors, material=material, plant=plant,
                    field="General Item Category Group",
                    value=val, error="ERSA/HIBE must be 'NORM'", table="MARA",
                )

        # BRGEW — MARA (0 / 0.000 counts as blank)
        gross_val = _to_float(row[gross_col]) if gross_col else None
        if gross_col and (gross_val is None or gross_val == 0):
            _append_error(
                errors, material=material, plant=plant, field="Gross Weight",
                value=row[gross_col], error="Mandatory field blank or zero", table="MARA",
            )
        # NTGEW — MARA (0 / 0.000 counts as blank)
        net_val = _to_float(row[net_col]) if net_col else None
        if net_col and (net_val is None or net_val == 0):
            _append_error(
                errors, material=material, plant=plant, field="Net Weight",
                value=row[net_col], error="Mandatory field blank or zero", table="MARA",
            )
        # GEWEI — MARA (only mandatory if the material actually has a weight > 0)
        has_weight = (gross_val is not None and gross_val > 0) or (net_val is not None and net_val > 0)
        if uow_col and has_weight and _is_blank(row[uow_col]):
            _append_error(
                errors, material=material, plant=plant,
                field="Unit of Weight (ISO Format)",
                value="", error="Mandatory field blank when a weight is set", table="MARA",
            )

        # MAGRV — MARA  (only VERP)
        if pack_grp_col and mt == "VERP":
            val = _norm(row[pack_grp_col]).upper()
            if val != "ZSHP":
                _append_error(
                    errors, material=material, plant=plant,
                    field="Matl Grp Pack.Matls",
                    value=val, error="VERP must be 'ZSHP'", table="MARA",
                )


def im_validar_mrp_view1(sap_df, errors):
    """MRP View 1 — all fields from MARC."""
    mat_col      = _pick(sap_df, ["Material", "MATNR"])
    plant_col    = _pick(sap_df, ["Plant", "WERKS"])
    mt_col       = _pick(sap_df, ["Material Type", "MTART"])
    mrp_type_col = _pick(sap_df, ["MRP Type", "DISMM"])
    mrp_ctrl_col = _pick(sap_df, ["MRP Controller", "DISPO"])
    mrp_grp_col  = _pick(sap_df, ["MRP Group", "DISGR"])
    psms_col     = _pick(sap_df, ["Plant Specific Material Status", "MMSTA"])
    lot_col      = _pick(sap_df, ["Lot Size", "Lot Sizing Procedure", "DISLS"])
    min_lot_col  = _pick(sap_df, ["Min Lot Size", "Minimum Lot Size", "BSTMI"])
    max_lot_col  = _pick(sap_df, ["Maximum Lot Size", "Max Lot Size", "BSTMA"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # DISMM — MARC
        if mrp_type_col:
            val      = _norm(row[mrp_type_col]).upper()
            expected = "ND" if plant == "TN02" and mt == "VERP" else "PD"
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Type", value="",
                              error="Mandatory field blank", table="MARC")
            elif val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="MRP Type", value=val,
                              error=f"Should be '{expected}'", table="MARC")

        # DISPO — MARC
        if mrp_ctrl_col:
            val = _norm(row[mrp_ctrl_col]).upper()
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Controller", value="",
                              error="Mandatory field blank", table="MARC")
            elif mt in {"HIBE", "ERSA"} and not re.fullmatch(r"NP\d", val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Controller", value=val,
                              error="HIBE/ERSA must match NPX (NP + digit)", table="MARC")
            elif mt == "VERP" and not re.fullmatch(r"P\d{2}", val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Controller", value=val,
                              error="VERP must match PXX (P + 2 digits)", table="MARC")

        # DISGR — MARC
        if mrp_grp_col:
            val      = _norm(row[mrp_grp_col]).upper()
            expected = "ZSTO" if plant == "TN02" and mt == "VERP" else "ZIM"
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="MRP Group", value="",
                              error="Mandatory field blank", table="MARC")
            elif val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="MRP Group", value=val,
                              error=f"Should be '{expected}'", table="MARC")

        # MMSTA — MARC
        if psms_col:
            val = normalize_code(row[psms_col], as_int=True)
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="Plant Specific Material Status", value="",
                              error="Mandatory field blank", table="MARC")
            elif mt == "VERP" and val not in {"45", "70", "99"}:
                _append_error(errors, material=material, plant=plant,
                              field="Plant Specific Material Status", value=val,
                              error="VERP must be 45, 70 or 99", table="MARC")
            elif mt in {"ERSA", "HIBE"} and val not in {"80", "70", "99"}:
                _append_error(errors, material=material, plant=plant,
                              field="Plant Specific Material Status", value=val,
                              error="ERSA/HIBE must be 80, 70 or 99", table="MARC")

        # DISLS — MARC
        if lot_col:
            val = _norm(row[lot_col]).upper()
            if val != "H1":
                _append_error(errors, material=material, plant=plant,
                              field="Lot Sizing Procedure ", value=val,
                              error="Should be 'H1'", table="MARC")

        # BSTMI — MARC (mandatory for ERSA / HIBE / VERP)
        if min_lot_col and mt in {"ERSA", "HIBE", "VERP"} and _is_blank(row[min_lot_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Minimum Lot Size", value="",
                          error="Mandatory field blank for ERSA/HIBE/VERP", table="MARC")

        # BSTMA — MARC (mandatory for ERSA / HIBE / VERP)
        if max_lot_col and mt in {"ERSA", "HIBE", "VERP"} and _is_blank(row[max_lot_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Maximum Lot Size", value="",
                          error="Mandatory field blank for ERSA/HIBE/VERP", table="MARC")


def im_validar_mrp_view2(sap_df, errors):
    """MRP View 2 — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    mt_col    = _pick(sap_df, ["Material Type", "MTART"])
    mg_col    = _pick(sap_df, ["Material Group", "MATKL"])
    proc_col  = _pick(sap_df, ["Procurement Type", "BESKZ"])
    spec_col  = _pick(sap_df, ["Special Procurement", "SOBSL"])
    stor_col  = _pick(sap_df, ["Storage Location Ext Proc", "LGPRO"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""
        mg       = _norm(row[mg_col]).upper() if mg_col else ""

        # BESKZ — MARC
        if proc_col:
            val = _norm(row[proc_col]).upper()
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="Procurement Type", value="",
                              error="Mandatory field blank", table="MARC")
            elif val != "F":
                _append_error(errors, material=material, plant=plant,
                              field="Procurement Type", value=val,
                              error="Must be 'F' for ERSA/HIBE/VERP", table="MARC")

        # SOBSL — MARC  (only TN02)
        if spec_col and plant == "TN02":
            val = _norm(row[spec_col]).upper()
            if val != "Z3":
                _append_error(errors, material=material, plant=plant,
                              field="Special Procurement Type", value=val,
                              error="TN02 must be 'Z3'", table="MARC")

        # LGPRO — MARC  (Stor. Loc. for Ext. Proc.)
        # Expected values per Plant × Material Type (from mapping table):
        #   PT01: HIBE → P160 | VERP → P155 | ERSA → P160
        #   PT02: HIBE → P260 | VERP → P255 | ERSA → P260
        #   TN01: HIBE → T160 | VERP → T155 | ERSA → T160
        #   TN02: HIBE → blank | VERP → T255 | ERSA → blank
        # Special case: ERSA + Material Group MT0020012 may also use:
        #   PT01 → P100 | PT02 → P200 | TN01 → T100 | TN02 → blank
        _LGPRO_EXPECTED = {
            ("PT01", "HIBE"): "P160",
            ("PT01", "VERP"): "P155",
            ("PT01", "ERSA"): "P160",
            ("PT02", "HIBE"): "P260",
            ("PT02", "VERP"): "P255",
            ("PT02", "ERSA"): "P260",
            ("TN01", "HIBE"): "T160",
            ("TN01", "VERP"): "T155",
            ("TN01", "ERSA"): "T160",
            ("TN02", "HIBE"): "",
            ("TN02", "VERP"): "T255",
            ("TN02", "ERSA"): "",
        }
        _LGPRO_ALLOWED_FOR_MT0020012 = {
            "PT01": "P100",
            "PT02": "P200",
            "TN01": "T100",
            "TN02": "",
        }
        if stor_col and mt in {"HIBE", "ERSA", "VERP"}:
            expected_lgpro = _LGPRO_EXPECTED.get((plant, mt))
            val_lgpro = _norm(row[stor_col]).upper()
            if expected_lgpro is None:
                # Plant/MT combination not in mapping — skip
                pass
            else:
                allowed_values = {expected_lgpro} if expected_lgpro != "" else {""}
                if mt == "ERSA" and mg == "MT0020012":
                    special_value = _LGPRO_ALLOWED_FOR_MT0020012.get(plant)
                    if special_value is not None:
                        allowed_values.add(special_value)
                if "" in allowed_values:
                    if not _is_blank(val_lgpro):
                        if val_lgpro not in allowed_values:
                            allowed_values_str = ", ".join(sorted(v for v in allowed_values if v)) or "blank"
                            _append_error(errors, material=material, plant=plant,
                                          field="Storage Location for Ext. Procurement", value=val_lgpro,
                                          error=f"Must be blank or one of: {allowed_values_str}", table="MARC")
                else:
                    if val_lgpro not in allowed_values:
                        allowed_values_str = ", ".join(sorted(allowed_values))
                        _append_error(errors, material=material, plant=plant,
                                      field="Storage Location for Ext. Procurement", value=val_lgpro,
                                      error=f"Should be one of: {allowed_values_str}", table="MARC")


def im_validar_mrp_view3(sap_df, errors):
    """MRP View 3 — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    per_col   = _pick(sap_df, ["Period Indicator", "PERKZ"])
    avail_col = _pick(sap_df, ["Availability Check", "MTVFP"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""

        # PERKZ — MARC
        if per_col:
            val = _norm(row[per_col]).upper()
            if val != "M":
                _append_error(errors, material=material, plant=plant,
                              field="Period Indicator", value=val,
                              error="Should be 'M'", table="MARC")

        # MTVFP — MARC
        if avail_col:
            val = normalize_code(row[avail_col], as_int=True, zero_pad=2)
            if val != "01":
                _append_error(errors, material=material, plant=plant,
                              field="Availability Check", value=row[avail_col],
                              error="Should be '01'", table="MARC")


def im_validar_purchasing(sap_df, errors):
    """Purchasing View — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    mt_col    = _pick(sap_df, ["Material Type", "MTART"])
    pg_col    = _pick(sap_df, ["Purchasing Group", "EKGRP"])
    auto_col  = _pick(sap_df, ["Auto PO", "Automatic Purchase Order Allowed", "AUTPU"])
    src_col   = _pick(sap_df, ["Source List Required", "KORDB"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # EKGRP — MARC
        if pg_col:
            val     = _norm(row[pg_col]).upper()
            allowed = {"IMR", "IMT", "IPS"}
            if plant == "TN02" and mt == "VERP":
                if val != "STO":
                    _append_error(errors, material=material, plant=plant,
                                  field="Purchasing Group", value=val,
                                  error="TN02 VERP must be 'STO'", table="MARC")
            elif mt == "VERP":
                if val != "IFL":
                    _append_error(errors, material=material, plant=plant,
                                  field="Purchasing Group", value=val,
                                  error="VERP must be 'IFL'", table="MARC")
            elif mt in {"HIBE", "ERSA"}:
                if val not in allowed:
                    _append_error(errors, material=material, plant=plant,
                                  field="Purchasing Group", value=val,
                                  error="HIBE/ERSA must be IMR, IMT or IPS", table="MARC")

        # AUTPU — MARC
        if auto_col and not _is_x(row[auto_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Automatic Purchase Order Allowed", value=row[auto_col],
                          error="Must be 'X'", table="MARC")

        # KORDB — MARC
        if src_col and not _is_x(row[src_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Source List Required", value=row[src_col],
                          error="Must be 'X'", table="MARC")


def im_validar_plant_storage(sap_df, errors):
    """Plant Data / Storage 1 — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    mt_col    = _pick(sap_df, ["Material Type", "MTART"])
    per_col   = _pick(sap_df, ["Period Ind.", "IPRKZ"])
    exp_col   = _pick(sap_df, ["Expiration Date", "MHDRZ"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # IPRKZ — MARC (Period Ind. for SLED — mandatory, must be "D" for ERSA/HIBE/VERP)
        if per_col and mt in {"ERSA", "HIBE", "VERP"}:
            val = _norm(row[per_col]).upper() 
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="Period Ind. for SLED", value="",
                              error="Mandatory field blank for ERSA/HIBE/VERP", table="MARC")
            elif val != "D":
                _append_error(errors, material=material, plant=plant,
                              field="Period Ind. for SLED", value=val,
                              error="Should be 'D'", table="MARC")

        # MHDRZ — MARC (Expiration Date — must be blank for ERSA/HIBE/VERP)
        if exp_col and mt in {"ERSA", "HIBE", "VERP"} and not _is_blank(row[exp_col]):
            _append_error(errors, material=material, plant=plant,
                          field="Expiration Date", value=row[exp_col],
                          error="Must be blank for ERSA/HIBE/VERP", table="MARC")


def im_validar_general_plant(sap_df, errors):
    """General Plant Data — all fields from MARC."""
    mat_col   = _pick(sap_df, ["Material", "MATNR"])
    plant_col = _pick(sap_df, ["Plant", "WERKS"])
    mt_col    = _pick(sap_df, ["Material Type", "MTART"])
    pc_col    = _pick(sap_df, ["Profit Center", "PRCTR"])
    neg_col   = _pick(sap_df, ["Negative Stocks", "XMCNG"])
    log_col   = _pick(sap_df, ["Log. Handling Group", "LOGGR"])

    for _, row in sap_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[plant_col]).upper() if plant_col else ""
        mt       = _norm(row[mt_col]).upper() if mt_col else ""

        # PRCTR — MARC
        if pc_col:
            val      = _norm(row[pc_col]).upper()
            expected = {
                "PT01": "PT00351",
                "PT02": "PT00361",
                "TN01": "TN00413",
                "TN02": "TN00413",
            }.get(plant)
            if _is_blank(val):
                _append_error(errors, material=material, plant=plant,
                              field="Profit Center", value="",
                              error="Mandatory field blank", table="MARC")
            elif expected and val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="Profit Center", value=val,
                              error=f"Should be '{expected}'", table="MARC")

        # XMCNG — MARC
        if neg_col:
            is_x = _is_x(row[neg_col])
            if mt == "VERP" and not is_x:
                _append_error(errors, material=material, plant=plant,
                              field="Neg. stocks in plant", value=row[neg_col],
                              error="VERP must be 'X'", table="MARC")
            elif mt != "VERP" and is_x:
                _append_error(errors, material=material, plant=plant,
                              field="Neg. stocks in plant", value=row[neg_col],
                              error="Only VERP may be 'X'", table="MARC")

        # LOGGR — MARC
        if log_col:
            val = _norm(row[log_col]).upper()
            if val != "NORM":
                _append_error(errors, material=material, plant=plant,
                              field="Log. Handling Group", value=val,
                              error="Should be 'NORM'", table="MARC")


def im_validar_mbew(mbew_df, mt_map, errors, plant_filter=None):
    """Accounting / Valuation — fields from MBEW."""
    if mbew_df is None or mbew_df.empty:
        return
    mat_col  = _pick(mbew_df, ["Material", "MATNR"])
    area_col = _pick(mbew_df, ["Plant", "Valuation Area", "BWKEY", "WERKS"])
    vc_col   = _pick(mbew_df, ["Valuation Class", "BKLAS"])
    pd_col   = _pick(mbew_df, ["Price Determination", "Material Price Determination: Control", "MLAST"])
    pc_col   = _pick(mbew_df, ["Price Control", "VPRSV"])
    pu_col   = _pick(mbew_df, ["Price Unit", "PEINH"])
    lifo_col = _pick(mbew_df, ["LIFO/FIFO-relevant", "XLIFO"])

    for _, row in mbew_df.iterrows():
        material = row[mat_col] if mat_col else ""
        plant    = _norm(row[area_col]).upper() if area_col else ""
        mt       = mt_map.get(_material_key(material), "")
        if mt not in IM_MATERIAL_TYPES:
            continue
        if plant_filter and plant != _norm(plant_filter).upper():
            continue

        # BKLAS — MBEW
        if vc_col:
            val      = normalize_code(row[vc_col], as_int=True)
            expected = {"HIBE": "3030", "ERSA": "3040"}.get(mt)
            if expected and val != expected:
                _append_error(errors, material=material, plant=plant,
                              field="Valuation Class", value=val,
                              error=f"{mt} must be '{expected}'", table="MBEW")

        # MLAST — MBEW  (ERSA/HIBE only)
        if pd_col and mt in {"ERSA", "HIBE"}:
            val = normalize_code(row[pd_col], as_int=True)
            if val != "2":
                _append_error(errors, material=material, plant=plant,
                              field="Price Control Determination", value=row[pd_col],
                              error="ERSA/HIBE must be '2'", table="MBEW")

        # VPRSV — MBEW  (ERSA/HIBE only)
        if pc_col and mt in {"ERSA", "HIBE"}:
            val = _norm(row[pc_col]).upper()
            if val != "V":
                _append_error(errors, material=material, plant=plant,
                              field="Price Control", value=val,
                              error="ERSA/HIBE must be 'V'", table="MBEW")

        # PEINH — MBEW
        if pu_col:
            val = normalize_code(row[pu_col], as_int=True)
            if val != "10000":
                _append_error(errors, material=material, plant=plant,
                              field="Price Unit", value=row[pu_col],
                              error="Must be '10000'", table="MBEW")

        # XLIFO — MBEW
        if lifo_col and not _is_x(row[lifo_col]):
            _append_error(errors, material=material, plant=plant,
                          field="LIFO/FIFO-relevant", value=row[lifo_col],
                          error="Must be 'X'", table="MBEW")


def im_validar_mara_wh_packaging(mara_df, mt_map, errors):
    """WH & Packaging fields — all from MARA."""
    if mara_df is None or mara_df.empty:
        return
    mat_col = _pick(mara_df, ["Material", "MATNR"])
    mg_col  = _pick(mara_df, ["Material Group", "MATKL"])
    wh_col  = _pick(mara_df, ["WH Material Group", "Warehouse Material Group"])
    hut_col = _pick(mara_df, ["Handling Unit Type", "HUTYP"])
    pkg_col = _pick(mara_df, ["Packaging Mat. Type", "Packaging material type"])
    ms_col  = _pick(mara_df, ["Maintenance status", "Maintenance Status", "PSTAT"])

    for _, row in mara_df.iterrows():
        material = row[mat_col] if mat_col else ""
        mt       = mt_map.get(_material_key(material), "")
        if mt not in IM_MATERIAL_TYPES:
            continue
        mg = _norm(row[mg_col]).upper() if mg_col else ""

        # WH Material Group — MARA  (ERSA + specific MG only) — must be blank
        if wh_col and mt == "ERSA" and mg == "MT0020012":
            val = _norm(row[wh_col]).upper()
            if not _is_blank(val):
                _append_error(errors, material=material, plant="",
                              field="WH Material Group", value=val,
                              error="ERSA MT0020012 must be blank", table="MARA")

        # Maintenance status — MARA  (ERSA + MT0020012 must contain letter 'C')
        if ms_col and mt == "ERSA" and mg == "MT0020012":
            ms_val = _norm(row[ms_col]).upper()
            if "C" not in ms_val:
                _append_error(errors, material=material, plant="",
                              field="Maintenance Status (MARA)", value=ms_val,
                              error="ERSA MT0020012 must have letter 'C' in Maintenance status",
                              table="MARA")

        # HUTYP — MARA  (mandatory for VERP)
        if hut_col and mt == "VERP" and _is_blank(row[hut_col]):
            _append_error(errors, material=material, plant="",
                          field="Handling Unit Type", value="",
                          error="VERP cannot be blank", table="MARA")

        # Packaging Mat. Type — MARA  (mandatory for VERP)
        if pkg_col and mt == "VERP" and _is_blank(row[pkg_col]):
            _append_error(errors, material=material, plant="",
                          field="Packaging Mat. Type", value="",
                          error="Mandatory for VERP", table="MARA")


def im_validar_mvke(mvke_df, mt_map, errors):
    """Sales / Distribution fields — MVKE."""
    if mvke_df is None or mvke_df.empty:
        return
    mat_col = _pick(mvke_df, ["Material", "MATNR"])
    icg_col = _pick(mvke_df, ["Item Category Group", "MTPOS"])
    tax_cols = [
        c for c in mvke_df.columns
        if "tax" in str(c).lower() or "TAX" in str(c)
    ]

    for _, row in mvke_df.iterrows():
        material = row[mat_col] if mat_col else ""
        mt       = mt_map.get(_material_key(material), "")
        if mt != "VERP":
            continue

        # MTPOS — MVKE
        if icg_col:
            val = _norm(row[icg_col]).upper()
            if val != "NORM":
                _append_error(errors, material=material, plant="",
                              field="Item Category Group (MVKE)", value=val,
                              error="VERP must be 'NORM'", table="MVKE")

        # Tax data — MVKE
        if tax_cols:
            has_tax = any(not _is_blank(row[c]) for c in tax_cols)
            if not has_tax:
                _append_error(errors, material=material, plant="",
                              field="Tax Data (MVKE)", value="",
                              error="Tax data required for VERP", table="MVKE")
