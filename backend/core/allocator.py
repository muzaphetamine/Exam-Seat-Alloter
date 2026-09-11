import re
from collections import defaultdict


def extract_course_from_usn(usn):
    match=re.search(r'\d{2}([A-Z]{2,3})\d{3}$', str(usn))
    return match.group(1) if match else 'UNKNOWN'


def group_students_by_course_subject(session_students):
    session_students['ExtractedCourse']=session_students['USN'].apply(extract_course_from_usn)
    course_subject_groups =defaultdict(list)
    for _, student in session_students.iterrows():
        key =(student['ExtractedCourse'], student['SubjectCode'])
        course_subject_groups[key].append(student)  
    return dict(course_subject_groups)


def create_room_pairs(course_subject_groups):
    groups=list(course_subject_groups.keys())
    room_pairs=[]
    used_groups=set()
    for i, group1 in enumerate(groups):
        if group1 in used_groups:
            continue  
        course1, subject1 =group1
        best_pair=None
        for j, group2 in enumerate(groups[i+1:], i+1):
            if group2 in used_groups:
                continue
            course2, subject2 = group2
            if course1!=course2 and subject1 != subject2:
                best_pair=group2
                break
        if not best_pair:
            for j, group2 in enumerate(groups[i+1:], i+1):
                if group2 in used_groups:
                    continue     
                course2, subject2=group2
                if course1!=course2:
                    best_pair=group2
                    break
        if not best_pair:
            for j, group2 in enumerate(groups[i+1:], i+1):
                if group2 in used_groups:
                    continue
                best_pair=group2
                break
        if best_pair:
            room_pairs.append((group1, best_pair))
            used_groups.add(group1)
            used_groups.add(best_pair)
        else:
            room_pairs.append((group1, None))
            used_groups.add(group1)
    return room_pairs


def allocate_students_to_room(group1_students, group2_students, rows, cols):
    room_layout=[['' for _ in range(cols)] for _ in range(rows)]
    course_header=['']*cols
    room_students=[]
    group1_idx=0
    group2_idx=0
    
    group1_course =group1_students[0]['ExtractedCourse'] if group1_students else ''
    group1_subject =group1_students[0]['SubjectCode'] if group1_students else ''
    group2_course =group2_students[0]['ExtractedCourse'] if group2_students else ''
    group2_subject =group2_students[0]['SubjectCode'] if group2_students else ''
    
    for col in range(cols):
        if col%2==0:
            current_group =group1_students
            current_idx =group1_idx
            course_header[col]=f"{group1_course}({group1_subject})" if group1_course else ''
        else:
            current_group =group2_students if group2_students else group1_students
            current_idx= group2_idx if group2_students else group1_idx
            course_header[col]=f"{group2_course}({group2_subject})" if group2_students and group2_course else f"{group1_course}({group1_subject})" if group1_course else ''
        
        for row in range(rows):
            if current_idx < len(current_group):
                student=current_group[current_idx]
                room_layout[row][col]=student['USN']
                room_students.append(student)
                current_idx+=1
            else:
                break
        
        if col%2==0 or not group2_students:
            group1_idx=current_idx
        else:
            group2_idx=current_idx
    return room_layout, room_students, course_header


