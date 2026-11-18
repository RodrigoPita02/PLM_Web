"""Leitura de ficheiros Excel carregados pelo utilizador."""

import io
import pandas as pd


def load_xlsx(file, *, sheet_name=None):
    """
    Loads an Excel file from multiple source types:
    - str (file path on disk)
    - Streamlit UploadedFile (uses .getvalue())
    - Generic file-like with .read()
    - io.BytesIO
    """
    if file is None:
        return None

    if isinstance(file, str):
        return pd.read_excel(file, engine="openpyxl", sheet_name=sheet_name)

    if hasattr(file, "getvalue"):
        bio = io.BytesIO(file.getvalue())
        return pd.read_excel(bio, engine="openpyxl", sheet_name=sheet_name)

    if hasattr(file, "read"):
        try:
            file.seek(0)
        except Exception:
            pass
        bio = io.BytesIO(file.read())
        return pd.read_excel(bio, engine="openpyxl", sheet_name=sheet_name)

    if isinstance(file, io.BytesIO):
        file.seek(0)
        return pd.read_excel(file, engine="openpyxl", sheet_name=sheet_name)

    raise ValueError(f"Unsupported file type: {type(file)}")
