from .allocator import extract_course_from_usn


def check_conflicts(students_df):
    conflicts=[] 
    for usn, group in students_df.groupby('USN'):
        if len(group)>1:
            student_data=group.iloc[0]
            conflicts.append({
                'USN': usn,
                'Name': student_data['Name'],
                'Branch': extract_course_from_usn(usn),
                'Subjects': ', '.join(f"{row['SubjectCode']} ({row.get('SubjectName', row['SubjectCode'])})" for _, row in group.iterrows()),
                'SubjectCount': len(group)
            })
    return conflicts