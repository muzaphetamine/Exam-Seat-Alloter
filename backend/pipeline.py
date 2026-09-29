import os
import re
from backend.ingestion.session import (
    extract_metadata,
    normalize_datetime,
    read_students_from_file
)
from backend.core.conflicts import check_conflicts
from backend.core.allocator import allocate_session
from backend.output.conflict_sheet import save_conflicts
from backend.output.workbook import create_workbook


def process_session_data(
    session_students,
    rooms_df,
    output_folder,
    input_filename,
    date_obj,
    start_dt=None,
    end_dt=None,
    session_label=None
):
    date_str = date_obj.strftime("%Y-%m-%d")

    if start_dt is not None and end_dt is not None:
        session_str=(
            f"{start_dt.strftime('%I:%M%p')} TO {end_dt.strftime('%I:%M%p')}"
        )
        filename_session_str=(
            f"{start_dt.strftime('%I-%M%p')} TO {end_dt.strftime('%I-%M%p')}"
        )
    else:
        session_str =str(session_label)
        filename_session_str = re.sub(r'[<>:"/\\|?*]', "_", session_str)
    
    conflicts =check_conflicts(session_students)
    if conflicts:
        conflict_file =save_conflicts(
            conflicts,
            input_filename,
            date_str,
            session_str,
            output_folder
        )
        return {
            "file": input_filename,
            "status": "conflict",
            "message": f"{len(conflicts)} conflicts found",
            "conflict_file": conflict_file
        }

    allocation_result = allocate_session(session_students, rooms_df)
    wb = create_workbook(
        allocation_result,
        session_students,
        rooms_df,
        date_obj,
        start_dt,
        end_dt,
        session_label
    )
    output_filename = (
        f"Roomallotment on "
        f"{date_obj.strftime('%d-%b-%Y')} "
        f"{filename_session_str}.xlsx"
    )
    output_path = os.path.join(output_folder, output_filename)
    wb.save(output_path)

    return {
        "file": input_filename,
        "status": "success",
        "output": output_filename,
        "allocated": len(allocation_result["allocated_students"]),
        "unallocated": len(allocation_result["unallocated_students"]),
        "total": len(session_students)
    }


def process_session_files(
    session_files,
    rooms_df,
    output_folder
):
    results=[]
    for session_file in session_files:
        try:
            date_str, time_str = extract_metadata(session_file)
            date_obj, start_dt, end_dt = normalize_datetime(date_str, time_str)

            if not date_obj:
                results.append({
                    "file": os.path.basename(session_file),
                    "status": "error",
                    "message": "Could not extract date/time from file"
                })
                continue

            students_df = read_students_from_file(session_file)
            result = process_session_data(
                students_df,
                rooms_df,
                output_folder,
                os.path.basename(session_file),
                date_obj,
                start_dt,
                end_dt
            )
            results.append(result)

        except Exception as exc:
            results.append({
                "file": os.path.basename(session_file),
                "status": "error",
                "message": str(exc)
            })
    return results


def process_centralized_data(
    students_df,
    schedule_df,
    rooms_df,
    output_folder
):
    results=[]
    for _, schedule in schedule_df.iterrows():
        raw_codes = str(schedule["SubjectCode"])
        subject_codes = [
            code.strip().upper()
            for code in re.split(r"[;,/]", raw_codes)
            if code.strip()
        ]
        session_students =students_df[
            students_df["SubjectCode"].isin(subject_codes)
        ].copy()
        if session_students.empty:
            continue

        date_obj = schedule["Date"]
        session_label = str(schedule["Session"]).strip()
        input_filename = (f"{date_obj.strftime('%Y-%m-%d')}_{session_label}.xlsx")

        try:
            result = process_session_data(
                session_students,
                rooms_df,
                output_folder,
                input_filename,
                date_obj,
                session_label=session_label
            )
            results.append(result)
        except Exception as exc:
            results.append({
                "file": input_filename,
                "status": "error",
                "message": str(exc)
            })
    return results