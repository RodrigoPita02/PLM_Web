"""Constantes do Indirect Materials (ERSA/HIBE/VERP)."""



IM_MATERIAL_TYPES = frozenset({"ERSA", "HIBE", "VERP"})


IM_PLANTS = frozenset({"PT01", "PT02", "TN01", "TN02"})


ERROR_COLS = [
    "Material", "Plant", "Field", "Value", "Error", "Error_Type", "Table", "View",
]


# IM-specific Field → View mapping (aligned with RolesViews.xlsx)
# Fields not in the Roles sheet are mapped to "Not applicable"
_IM_FIELD_TO_VIEW = {
    # Basic Data (K)
    "Material":                             "K",
    "Description":                          "K",
    "Material Type":                        "K",
    "Base UOM":                             "K",
    "Division":                             "K",
    "Material Group":                       "K",
    "General Item Category Group":          "K",
    "Gross Weight":                         "K",
    "Net Weight":                           "K",
    "Unit of Weight (ISO Format)":          "K",
    "Matl Grp Pack.Matls":                  "K",
    # MRP / Purchasing (D / E)
    "MRP Type":                             "D",
    "MRP Controller":                       "D",
    "MRP Group":                            "D",
    "Plant Specific Material Status":       "D",
    "Lot Sizing Procedure ":                "D",
    "Minimum Lot Size":                     "D",
    "Maximum Lot Size":                     "D",
    "Procurement Type":                     "D",
    "Special Procurement Type":             "D",
    "Storage Location for Ext. Procurement":"D",
    "Period Indicator":                     "D",
    "Availability Check":                   "D",
    "Purchasing Group":                     "E",
    "Automatic Purchase Order Allowed":     "E",
    "Source List Required":                 "E",
    # Storage / General Plant (L)
    "Period Ind. for SLED":                 "L",
    "Expiration Date":                      "L",
    "Profit Center":                        "L",
    "Neg. stocks in plant":                 "L",
    "Log. Handling Group":                  "L",
    # Accounting (B)
    "Valuation Class":                      "B",
    "Price Control Determination":          "B",
    "Price Control":                        "B",
    "Price Unit":                           "B",
    "LIFO/FIFO-relevant":                   "B",
    # Sales (V)
    "Packaging Mat. Type":                  "V",
    "Item Category Group (MVKE)":           "V",
    "Tax Data (MVKE)":                      "V",
    # Warehouse Management (S) — only conditionally required (see _im_postprocess)
    "WH Material Group":                    "S",
    "Maintenance Status (MARA)":            "C",
    # No view
    "Handling Unit Type":                   "Not applicable",
}


_IM_VIEW_TO_DESC = {
    "K": "Basic Data",
    "D": "MRP",
    "E": "Purchasing",
    "L": "Storage",
    "B": "Accounting",
    "V": "Sales",
    "S": "Warehouse management",
    "C": "Classification",
}


# Required views per IM material type
# Note: "B" (Accounting) is NOT required for VERP — packaging materials
# are not accounting-relevant the same way, so it's excluded from scope.
_IM_REQUIRED_VIEWS = {
    "ERSA": {"K", "D", "E", "L", "B"},
    "HIBE": {"K", "D", "E", "L", "B"},
    "VERP": {"K", "D", "E", "L", "V"},
}


# All fields expected per view (for fields_expected column in view_tabs)
_IM_VIEW_TO_FIELDS: dict[str, list[str]] = {}


for _f, _v in _IM_FIELD_TO_VIEW.items():
    if _v != "Not applicable":
        _IM_VIEW_TO_FIELDS.setdefault(_v, []).append(_f)
