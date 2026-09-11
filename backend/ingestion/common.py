import os
import pandas as pd
from backend.config import SUPPORTED_EXCEL_EXTENSIONS


def excel_files(folder):
    if not os.path.isdir(folder):
        return []
    return sorted(
        os.path.join(folder, name)
        for name in os.listdir(folder)
        if os.path.isfile(os.path.join(folder, name))
        and os.path.splitext(name)[1].lower() in SUPPORTED_EXCEL_EXTENSIONS
    )


def read_excel_table(filepath):
    ext=os.path.splitext(filepath)[1].lower()
    return pd.read_excel(filepath, engine="xlrd" if ext == ".xls" else None)


def clean_columns(df):
    df=df.copy()
    df.columns=[str(c).strip() for c in df.columns]
    return df


def missing_columns(df, required):
    return sorted(required - set(df.columns))