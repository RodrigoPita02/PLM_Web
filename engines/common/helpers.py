"""Helpers genéricos de normalização e DataFrames, partilhados por todos os engines."""

import pandas as pd


def normalize_code(value, *, as_int=False, zero_pad=None):
    """
    Normalizes a code value from Excel:
    - Strips whitespace.
    - Converts numeric strings (e.g. '2.0', '40.0') to integer strings.
    - Applies zero-padding if specified (e.g. zero_pad=4 -> '1' -> '0001').
    Returns a normalized string.
    """
    if pd.isna(value):
        return ''
    s = str(value).strip()

    try:
        num = float(s)
        if as_int or s.endswith('.0') or s.isdigit():
            s = str(int(num))
    except ValueError:
        pass

    if zero_pad:
        try:
            s = str(int(s)).zfill(zero_pad)
        except ValueError:
            if len(s) < zero_pad and s.isdigit():
                s = s.zfill(zero_pad)

    return s


def _normalize_columns(df, rename_map):
    """Case-insensitive column rename (keys in rename_map are matched lowercased)."""
    df = df.copy()
    df.columns = [
        rename_map.get(str(col).strip().lower(), str(col).strip())
        for col in df.columns
    ]
    if df.columns.duplicated().any():
        df = df.T.groupby(level=0).first().T
    return df


def _normalize_df_columns(df, rename_map):
    """Same as _normalize_columns but tolerates df=None (used on optional tables)."""
    if df is None:
        return df
    df = df.copy()
    df.columns = [
        rename_map.get(str(col).strip().lower(), str(col).strip())
        for col in df.columns
    ]
    if df.columns.duplicated().any():
        df = df.T.groupby(level=0).first().T
    return df


def _find_col_ci(df, candidates):
    """Case-insensitive column lookup. Returns the real column name or None."""
    if df is None:
        return None
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        hit = lookup.get(str(cand).strip().lower())
        if hit is not None:
            return hit
    return None


def _normalize_code_as_str_basic(x):
    """Converts numeric-like values (e.g. '55.0') to plain integers ('55')."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    try:
        return str(int(float(s)))
    except Exception:
        return s


def _pick_col_upper(df, candidates):
    """Returns the first matching column name from a list of candidates."""
    return next((c for c in candidates if c in df.columns), None)


def _norm_str(x) -> str:
    """Normalises value to stripped string; NaN -> ''."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    return str(x).strip()


def _safe_name(uploaded, default_name: str) -> str:
    """
    Returns the filename of an uploaded file object (Streamlit UploadedFile
    or any object with a .name attribute).  Falls back to default_name.
    """
    return getattr(uploaded, "name", default_name) if uploaded is not None else default_name


def _coerce_df_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalises DataFrame column dtypes so PyArrow / Streamlit can serialise
    without errors.

    Problems addressed:
      - 'Line' column: mixed int/str object dtype  -> Int64 (nullable int)
      - Categorical columns -> converted to string to avoid Arrow edge-cases
      - Mixed-type object columns -> converted to string as last resort
    """
    if df is None or df.empty:
        return df
    df = df.copy()

    for col in df.columns:
        if col == 'Line':
            try:
                df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')
            except Exception:
                df[col] = df[col].astype(str)
            continue

        if hasattr(df[col], 'cat'):
            df[col] = df[col].astype(str)
            continue

        if df[col].dtype == object:
            types_present = {type(v) for v in df[col].dropna()}
            if len(types_present) > 1:
                df[col] = df[col].astype(str).replace('nan', '')
            elif types_present == {int} or types_present == {float}:
                try:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                except Exception:
                    pass

    return df
