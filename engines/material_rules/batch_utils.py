"""Utilitários de Batch Management e renomeação de campos para display."""

import pandas as pd

from engines.common.helpers import _norm_str
from engines.material_rules.constants import (
    BATCH_GLOBAL_FIELD,
    BATCH_MANAGEMENT_HALB_LIKE_MATERIALS,
    FIELD_DISPLAY_RENAMES,
    _BATCH_PLANT_FIELD_KEYS,
)

def _field_key(x):
    """Case / whitespace-insensitive key used to match field names."""
    import re as _re
    return _re.sub(r'\s+', ' ', str(x)).strip().lower()


def apply_display_field_renames(df_):
    """
    Rename the 'Field' column VALUES to the names shown in the report
    (FIELD_DISPLAY_RENAMES). Matching ignores case and extra spaces.
    Safe on empty frames / frames without a 'Field' column.
    """
    if df_ is None or not hasattr(df_, 'columns') or 'Field' not in df_.columns or df_.empty:
        return df_
    lookup = {_field_key(k): v for k, v in FIELD_DISPLAY_RENAMES.items()}
    df_ = df_.copy()
    df_['Field'] = df_['Field'].astype(object).map(
        lambda f: lookup.get(_field_key(f), f) if isinstance(f, str) else f
    )
    return df_


def consolidar_batch_management_errors(resultados, sap_df=None):
    """
    If a material has the plant-level Batch Management error in ALL of its plants
    (at least 2) AND the global Batch Management error, keep ONLY the global one.
    Works on the raw `resultados` list (in place). Returns the number of removed errors.
    """
    if not resultados:
        return 0

    global_key = _field_key(BATCH_GLOBAL_FIELD)
    global_mats = set()
    err_plants = {}
    for r in resultados:
        fk = _field_key(r.get('Field', ''))
        mat = _norm_str(r.get('Material'))
        if fk == global_key:
            global_mats.add(mat)
        elif fk in _BATCH_PLANT_FIELD_KEYS:
            err_plants.setdefault(mat, set()).add(_norm_str(r.get('Plant')).upper())

    if not global_mats:
        return 0

    all_plants = {}
    if sap_df is not None and not sap_df.empty and {'Material', 'Plant'}.issubset(sap_df.columns):
        for m, pl in zip(sap_df['Material'], sap_df['Plant']):
            all_plants.setdefault(_norm_str(m), set()).add(_norm_str(pl).upper())

    drop = set()
    for mat in global_mats:
        ep = err_plants.get(mat, set())
        plants = all_plants.get(mat, ep)
        if len(ep) >= 2 and plants <= ep:
            drop.add(mat)

    if not drop:
        return 0

    before = len(resultados)
    resultados[:] = [
        r for r in resultados
        if not (
            _field_key(r.get('Field', '')) in _BATCH_PLANT_FIELD_KEYS
            and _norm_str(r.get('Material')) in drop
        )
    ]
    return before - len(resultados)


def get_batch_management_effective_material_type(material, real_mt, plant_type):
    """
    Effective Material Type for the Batch Management rules only.

    ROH + Plant Type PP/COI on one of BATCH_MANAGEMENT_HALB_LIKE_MATERIALS
    is treated as 'HALB'. Everything else keeps its real Material Type.
    """
    material = '' if material is None or (isinstance(material, float) and pd.isna(material)) else str(material).strip()
    real_mt = '' if real_mt is None or (isinstance(real_mt, float) and pd.isna(real_mt)) else str(real_mt).strip().upper()
    plant_type = '' if plant_type is None or (isinstance(plant_type, float) and pd.isna(plant_type)) else str(plant_type).strip().upper()

    if (
        material.upper() in BATCH_MANAGEMENT_HALB_LIKE_MATERIALS
        and real_mt == 'ROH'
        and plant_type in {'PP', 'COI'}
    ):
        return 'HALB'
    return real_mt
