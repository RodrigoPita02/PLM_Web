"""Estado partilhado do SSA (mapa documento→tipo da execução corrente)."""


# Populated by run_ssa_validation. The compatibility helper below deliberately
# uses AUART and no longer reads the first digit of the document number.
_SSA_ACTIVE_DOC_TYPE_MAP = {}
