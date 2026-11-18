"""Constantes do SSA (tipos de ordem, colunas de diagnóstico)."""


SSA_DIAGNOSTIC_COLUMNS = [
    "Rule",
    "Check",
    "Table",
    "Logical Field",
    "Accepted Column Names",
    "Status",
    "Impact",
]


SSA_ORDER_TYPES_ZLZ_GROUP = {"ZLZ", "ZLZM"}


SSA_ORDER_TYPES_ZTAM_GROUP = {"ZTAM", "ZOR", "ZPRO", "ZFD", "ZINI", "ZORA", "ZFD2", "ZORM"}


SSA_ORDER_TYPES_ZDIC_GROUP = {"ZDIC", "ZMC", "ZRE", "ZDSC", "ZRE2", "ZMC2"}


SSA_ORDER_TYPES_ZMD_GROUP = {"ZMD", "ZDSD", "ZDID", "ZMD2"}


# Convenience union used by several rules that apply to both the ZLZ/ZLZM
# and ZTAM/ZOR/... Sales Document Types.
SSA_ORDER_TYPES_ZLZ_OR_ZTAM_GROUP = SSA_ORDER_TYPES_ZLZ_GROUP | SSA_ORDER_TYPES_ZTAM_GROUP


# Sales Document Types for which Customer Group 2 is mandatory.
SSA_ORDER_TYPES_CUSTOMER_GROUP_2 = {"ZLZ", "ZLZM", "ZOR", "ZTAM"}


# Every Sales Document Type known to the SSA engine.
SSA_ORDER_TYPES_ALL = (
    SSA_ORDER_TYPES_ZLZ_GROUP
    | SSA_ORDER_TYPES_ZTAM_GROUP
    | SSA_ORDER_TYPES_ZDIC_GROUP
    | SSA_ORDER_TYPES_ZMD_GROUP
)


# S-materials validated, for the applicable Sales Document Types, against
# the "Acct Assmt Grp Mat." value carried for them in MVKE (see
# _ssa_build_s_material_mvke_map and Rule 3 inside ssa_validar_mandatory_fields).
SSA_S_MATERIAL_LIST = [
    "S-306", "S-503", "S-304", "S-300", "S-200", "S-202", "S-301", "S-309",
    "S-100", "S-302", "S-507", "S-506", "S-305", "S-303", "S-502", "S-504",
    "S-505", "S-101", "S-102", "S-500", "S-501", "S-308", "S-509", "S-310",
    "S-311",
]


# Order types whose items must NEVER be reported by any SSA rule when the
# Reason for Rejection (VBAP-ABGRU) is filled. Add "ZLZM" here to extend it.
SSA_REJECTION_EXCLUDED_ORDER_TYPES = {"ZLZ"}
