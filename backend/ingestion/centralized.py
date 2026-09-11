import os
import pandas as pd
from .common import read_excel_table, clean_columns, missing_columns

CENTRALIZED_STUDENT_COLUMNS = {"USN","Name","Course","SubjectCode"}
CENTRALIZED_SCHEDULE_COLUMNS = {"Date","Session","SubjectCode"}

def read_centralized_students_from_files(filepaths):
    frames=[]
    errors=[]
    for filepath in filepaths:
        try:
            df =clean_columns(read_excel_table(filepath))
            missing=missing_columns(df, CENTRALIZED_STUDENT_COLUMNS)
            if missing:
                raise ValueError(f"missing required columns: {', '.join(missing)}")
            df =df[["USN", "Name", "Course", "SubjectCode"]].copy()
            df =df.dropna(how="all")
            for col in ["USN", "Name", "Course", "SubjectCode"]:
                df[col] =df[col].fillna("").astype(str).str.strip()
            if (df["USN"] == "").any() or (df["SubjectCode"] == "").any():
                raise ValueError("USN and SubjectCode cannot be blank")
            frames.append(df)
        except Exception as exc:
            errors.append(f"{os.path.basename(filepath)}: {exc}")
    if errors:
        raise ValueError("Student input error(s):\n- " + "\n- ".join(errors))
    if not frames:
        raise ValueError("No student Excel files were supplied.")
    merged= pd.concat(frames, ignore_index=True)
    duplicate_rows=merged.duplicated(subset=["USN", "SubjectCode"], keep=False)
    if duplicate_rows.any():
        duplicates =merged.loc[duplicate_rows, ["USN", "SubjectCode"]].drop_duplicates()
        details= [f"{r.USN} / {r.SubjectCode}" for r in duplicates.itertuples()]
        raise ValueError("Duplicate student-subject records found: " + ", ".join(details))
    return merged


def read_centralized_schedules_from_files(filepaths):
    frames=[]
    errors=[]
    for filepath in filepaths:
        try:
            df =clean_columns(read_excel_table(filepath))
            missing =missing_columns(df, CENTRALIZED_SCHEDULE_COLUMNS)
            if missing:
                raise ValueError(f"missing required columns: {', '.join(missing)}")
            df = df[["Date", "Session", "SubjectCode"]].copy()
            df = df.dropna(how="all")
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            if df["Date"].isna().any():
                raise ValueError("contains an invalid or blank Date")
            df["Session"] =df["Session"].fillna("").astype(str).str.strip()
            df["SubjectCode"] =df["SubjectCode"].fillna("").astype(str).str.strip()
            if (df["Session"] == "").any() or (df["SubjectCode"] == "").any():
                raise ValueError("Session and SubjectCode cannot be blank")
            frames.append(df)
        except Exception as exc:
            errors.append(f"{os.path.basename(filepath)}: {exc}")

    if errors:
        raise ValueError("Schedule input error(s):\n- " + "\n- ".join(errors))
    if not frames:
        raise ValueError("No schedule Excel files were supplied.")

    merged = pd.concat(frames, ignore_index=True)
    duplicate_rows=merged.duplicated(
        subset=["Date", "Session", "SubjectCode"], keep=False
    )
    if duplicate_rows.any():
        duplicates=merged.loc[duplicate_rows, ["Date", "Session", "SubjectCode"]].drop_duplicates()
        details =[
            f"{r.Date.strftime('%d-%m-%Y')} / {r.Session} / {r.SubjectCode}"
            for r in duplicates.itertuples()
        ]
        raise ValueError("Duplicate schedule records found: " + ", ".join(details))
    return merged

