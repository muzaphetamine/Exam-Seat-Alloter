import os
import pandas as pd
from .common import read_excel_table, clean_columns, missing_columns

ROOM_COLUMNS = {"RoomName", "Rows", "Cols"}


def read_rooms_from_files(filepaths):
    frames=[]
    errors=[]
    for filepath in filepaths:
        try:
            df =_clean_columns(_read_excel_table(filepath))
            missing =_missing_columns(df, ROOM_COLUMNS)
            if missing:
                raise ValueError(f"missing required columns: {', '.join(missing)}")
            df=df[["RoomName", "Rows", "Cols"]].copy()
            df=df.dropna(how="all")
            df["RoomName"] =df["RoomName"].astype(str).str.strip()
            if (df["RoomName"] == "").any() or df["RoomName"].eq("nan").any():
                raise ValueError("contains a blank RoomName")
            df["Rows"] =pd.to_numeric(df["Rows"], errors="coerce")
            df["Cols"] =pd.to_numeric(df["Cols"], errors="coerce")
            if df[["Rows", "Cols"]].isna().any().any():
                raise ValueError("Rows and Cols must be numeric")
            if (df[["Rows", "Cols"]] <= 0).any().any():
                raise ValueError("Rows and Cols must be greater than zero")
            df["Rows"] =df["Rows"].astype(int)
            df["Cols"] =df["Cols"].astype(int)
            frames.append(df)
        except Exception as exc:
            errors.append(f"{os.path.basename(filepath)}: {exc}")
    if errors:
        raise ValueError("Room input error(s):\n- " + "\n- ".join(errors))
    if not frames:
        raise ValueError("No room Excel files were supplied.")
    merged= pd.concat(frames, ignore_index=True)
    duplicate_names= merged[merged.duplicated("RoomName", keep=False)]["RoomName"].unique().tolist()
    if duplicate_names:
        raise ValueError("Duplicate room names found across room files: " + ", ".join(map(str, duplicate_names)))
    return merged
