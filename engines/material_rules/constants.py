"""Constantes das regras de validação de material master (SAP)."""


RULE_MATERIAL_TYPE_OVERRIDES = {
    "VPRMKF-10849-LB": {
        "__default__": "FERT",
        "validar_enterproj_number": "HALB",
        "validar_material_group": "HALB",
        "validar_valuation_class": "HALB",
    }
}


# Field name used by the MARA "Maintenance status must contain 'C'" validation.
# Errors of this field are treated as a MISSING VIEW (view 'C' - Classification),
# so they go to the Missing View tables instead of "Field With Error".
MARA_MAINT_STATUS_C_FIELD = 'Maintenance Status (MARA)'


# Materials that do NOT need view 'C' (Classification) -> never reported by the
# MARA "Maintenance status must contain 'C'" rule, hence never shown as
# Missing View 'C' nor in the Field With Error / summary / coverage tables.
# Independent from BATCH_MANAGEMENT_HALB_LIKE_MATERIALS (that one also changes
# the Batch Management rules; this one only affects view 'C').
MARA_MAINT_STATUS_C_EXEMPT_MATERIALS = {
    'VPTXDF-10A855-DA',
}


# Field names as DISPLAYED in the generated file (internal names, Roles lookups
# and field_rules keep the original names).
FIELD_DISPLAY_RENAMES = {
    'Export/Import Group': 'Intrastat Group',
    'Batch Management Required Ind.': 'Batch Management (Plant)',
    'Batch Management(Plant)': 'Batch Management (Plant)',
    'Batch Management (MARA)': 'Batch Management (Global)',
}


BATCH_GLOBAL_FIELD = 'Batch Management (Global)'


_BATCH_PLANT_FIELD_KEYS = {
    'batch management required ind.',
    'batch management(plant)',
    'batch management (plant)',
}


BATCH_MANAGEMENT_HALB_LIKE_MATERIALS = {
    "VPTVQF-15K258-AC",
    "VPVFEF-10841-BA",
    "VPVFEF-10841-BC",
    "VPVFEF-19G526-AA",
    "VPVFEF-19G526-CA",
    "VPWXRF-10841-AA",
}
