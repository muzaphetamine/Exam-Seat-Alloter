def check_conflicts(students_df):
    conflicts=[]
    for usn, group in students_df.groupby("USN"):
        if len(group)>1:
            student_data =group.iloc[0]
            branch = student_data.get("Course", "UNKNOWN")
            conflicts.append({
                "USN": usn,
                "Name": student_data["Name"],
                "Branch": branch,
                "Subjects": ", ".join(
                    f"{row['SubjectCode']} ({row.get('SubjectName', row['SubjectCode'])})"
                    for _, row in group.iterrows()
                ),
                "SubjectCount": len(group)
            })
    return conflicts