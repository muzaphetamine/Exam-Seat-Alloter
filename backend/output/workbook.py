from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
from openpyxl import Workbook
from backend.output.room_sheet import create_room_sheet
from backend.output.summary_sheet import create_summary_sheet

THIN_BORDER=Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
HEADER_FILL=PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
TITLE_FONT=Font(bold=True, size=14, color="FFFFFF")
HEADER_FONT=Font(bold=True, color="FFFFFF")
BODY_FONT=Font(size=10)


def build_master_allocations(room_allocations):
    allocations=[]
    for room in room_allocations:
        room_name = room["Room"]
        for student in room["Students"]:
            branch = student["Course"]
            subcode = student["SubjectCode"]
            found=False
            for alloc in allocations:
                if (alloc["Room"]==room_name and alloc["Branch"]==branch and alloc["SubCode"]==subcode):
                    alloc["USNs"].append(student["USN"])
                    alloc["Total"]+=1
                    found=True
                    break
            if not found:
                allocations.append({
                    "Room": room_name,
                    "Branch": branch,
                    "SubCode": subcode,
                    "SubName": student.get("SubjectName", student["SubjectCode"]),
                    "Sem": student.get("Semester", ""),
                    "USNs": [student["USN"]],
                    "Total": 1
                })
    return allocations


def create_room_allotment_sheet(
    wb,
    allocations,
    date_obj,
    start_dt=None,
    end_dt=None,
    session_label=None
):
    allocations =sorted(allocations, key=lambda x: (x["Branch"], x["SubCode"]))
    ws =wb.create_sheet("Room Allocation", 0)
    date_str =date_obj.strftime("%d-%b-%Y")
    #session_str =f"{start_dt.strftime('%I:%M%p')} TO {end_dt.strftime('%I:%M%p')}"
    if start_dt is not None and end_dt is not None:
        session_str=(
            f"{start_dt.strftime('%I:%M%p')} TO "
            f"{end_dt.strftime('%I:%M%p')}"
        )
    elif session_label:
        session_str = str(session_label)
    else:
        session_str = ""

    ws.append(["EXAM SEAT ALLOCATION"])
    ws.merge_cells("A1:G1")
    ws["A1"].font=TITLE_FONT
    ws["A1"].fill=HEADER_FILL
    ws["A1"].alignment=Alignment(horizontal="center", vertical="center")
    ws["A1"].border=THIN_BORDER
    ws.row_dimensions[1].height=30

    ws.append([f"DATE : {date_str}    SESSION : {session_str}"])
    ws.merge_cells("A2:G2")
    ws["A2"].font =Font(name='Calibri', bold=True, size=11)
    ws["A2"].alignment=Alignment(horizontal="center")
    headers=["Room No", "BRANCH", "Sub code", "Sub Name", "Sem", "ALLOTTED USN'S", "TOTAL"]
    ws.append(headers)
    
    for cell in ws[3]:
        cell.font=HEADER_FONT
        cell.fill=HEADER_FILL
        cell.border=THIN_BORDER
        cell.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height=25
    
    for alloc in allocations:
        ws.append([
            alloc["Room"],
            alloc["Branch"],
            alloc["SubCode"],
            alloc["SubName"],
            alloc["Sem"],
            ", ".join(alloc["USNs"]),
            alloc["Total"]
        ])
        
        r=ws.max_row
        for c in range(1,8):
            cell=ws.cell(row=r, column=c)
            cell.font=BODY_FONT
            cell.border=THIN_BORDER
            cell.alignment=Alignment(
                horizontal="center" if c in [1, 2, 3, 5, 7] else "left",
                vertical="top",
                wrap_text=True
            )
            if (r-4)%2==0:
                cell.fill =PatternFill(start_color='F8F9FA', end_color='F8F9FA', fill_type='solid')
    
    ws.column_dimensions["A"].width=12
    ws.column_dimensions["B"].width=15
    ws.column_dimensions["C"].width=12
    ws.column_dimensions["D"].width=35
    ws.column_dimensions["E"].width=8
    ws.column_dimensions["F"].width=80
    ws.column_dimensions["G"].width=10


def create_unallocated_sheet(wb, unallocated_students):
    if unallocated_students:
        ws_unalloc = wb.create_sheet(title="Unallocated Students")
        ws_unalloc.append(["Unallocated Students"])
        ws_unalloc['A1'].font = Font(name='Calibri', bold=True, size=14, color='CC0000')
        ws_unalloc.merge_cells('A1:D1')
        ws_unalloc['A1'].alignment = Alignment(horizontal='center')
        ws_unalloc.append([])
        
        ws_unalloc.append(["USN", "Name", "Course", "SubjectCode"])
        for col_idx in range(1,5):
            cell=ws_unalloc.cell(row=3, column=col_idx)
            cell.font=Font(name='Calibri', bold=True)
            cell.fill=PatternFill(start_color='FFE6E6', end_color='FFE6E6', fill_type='solid')
            cell.alignment=Alignment(horizontal='center')
            cell.border=THIN_BORDER
        
        for idx, student in enumerate(unallocated_students):
            ws_unalloc.append([student["USN"], student["Name"], student["Course"], student["SubjectCode"]])   
            row_num =4+idx
            for col_idx in range(1,5):
                cell=ws_unalloc.cell(row=row_num, column=col_idx)
                cell.font=Font(name='Calibri', size=9)
                cell.border=THIN_BORDER
                if col_idx==1:
                    cell.alignment = Alignment(horizontal='center')


def create_workbook(
    allocation_result,
    session_students,
    rooms_df,
    date_obj,
    start_dt=None,
    end_dt=None,
    session_label=None
):
    wb=Workbook()
    if wb.active: wb.remove(wb.active)

    room_allocations = allocation_result["rooms"]
    allocated_students = allocation_result["allocated_students"]
    unallocated_students = allocation_result["unallocated_students"]

    master_allocations = build_master_allocations(room_allocations)
    create_room_allotment_sheet(
        wb,
        master_allocations,
        date_obj,
        start_dt,
        end_dt,
        session_label
    )

    for room in room_allocations:
        create_room_sheet(
            wb,
            room["Room"],
            room["Layout"],
            room["CourseHeader"],
            room["Students"]
        )
    create_unallocated_sheet(wb, unallocated_students)
    create_summary_sheet(
        wb,
        session_students,
        allocated_students,
        unallocated_students,
        rooms_df,
        date_obj,
        start_dt,
        end_dt,
        session_label
    )
    return wb