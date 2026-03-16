# utils/fast_loader.py
from __future__ import annotations
import io
import polars as pl
import pandas as pd

def _to_bytesio(uploaded_or_path):
    # Streamlit UploadedFile
    if hasattr(uploaded_or_path, "getvalue"):
        return io.BytesIO(uploaded_or_path.getvalue())
    # BytesIO/file-like
    if hasattr(uploaded_or_path, "read"):
        try:
            uploaded_or_path.seek(0)
        except Exception:
            pass
        return io.BytesIO(uploaded_or_path.read())
    # str path
    return uploaded_or_path  # polars/pandas aceitam str path diretamente


def _read_with_polars(src, sheet_name):
    # Polars não aceita sheet_name como inteiro — converter para None e
    # seleccionar pelo índice manualmente quando necessário
    sheet_by_index = isinstance(sheet_name, int)

    # 1) tentativa direta com override global para texto
    if not sheet_by_index:
        try:
            return pl.read_excel(src, sheet_name=sheet_name, dtype=str)
        except Exception:
            pass

    # 2) ler todas as folhas e escolher pela posição ou pela 1ª não vazia
    try:
        sheets = pl.read_excel(src, sheet_name=None, dtype=str)
        if isinstance(sheets, dict) and sheets:
            if sheet_by_index:
                keys = list(sheets.keys())
                idx  = sheet_name if sheet_name < len(keys) else 0
                df   = sheets[keys[idx]]
                if df is not None and df.shape[0] > 0:
                    return df
            for name, df in sheets.items():
                try:
                    if df is not None and df.shape[0] > 0 and df.shape[1] > 0:
                        return df
                except Exception:
                    continue
            first_key = next(iter(sheets))
            return sheets[first_key]
    except Exception:
        pass

    raise ValueError("Polars could not match any sheet for this workbook")

def _read_with_pandas(src, sheet_name):
    """
    Lê Excel com pandas (openpyxl) – compatibilidade máxima.
    Retorna pandas.DataFrame.
    """
    try:
        return pd.read_excel(src, engine="openpyxl", sheet_name=sheet_name)
    except Exception:
        # enumerar folhas e escolher 1ª não vazia
        try:
            xls = pd.ExcelFile(src, engine="openpyxl")
            for name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=name)
                if df is not None and not df.empty:
                    return df
            # se todas vazias, devolve a primeira
            return pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        except Exception as e:
            raise e


def load_xlsx_fast(uploaded_file, sheet_name=0):
    """
    Comportamento:
    - sheet_name=None → devolve dict[str, pandas.DataFrame] (igual ao pandas)
    - sheet_name=0/"SAPUI5 Export" → devolve pandas.DataFrame
    """

    if uploaded_file is None:
        return None

    src = _to_bytesio(uploaded_file)

    # ---------------------------------
    # CASO 1 — sheet_name=None → DICT
    # ---------------------------------
    if sheet_name is None:
        # Tenta Polars primeiro
        try:
            sheets = pl.read_excel(src, sheet_name=None, dtype=str)
            out = {name: df.to_pandas() for name, df in sheets.items()}
            return out
        except Exception:
            pass

        # Fallback pandas
        try:
            xls = pd.ExcelFile(src, engine="openpyxl")
            out = {}
            for name in xls.sheet_names:
                out[name] = pd.read_excel(xls, sheet_name=name)
            return out
        except Exception as e:
            raise e

    # ---------------------------------
    # CASO 2 — sheet_name específico → DF
    # ---------------------------------
    # 1) Tenta Polars
    try:
        pl_df = _read_with_polars(src, sheet_name=sheet_name)
        return pl_df.to_pandas()
    except Exception:
        pass

    # 2) Fallback pandas
    return _read_with_pandas(src, sheet_name=sheet_name)