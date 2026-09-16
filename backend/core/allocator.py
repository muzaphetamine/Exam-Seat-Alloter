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


def allocate_students_to_room(group1_key,group2_key,remaining_groups,rows,cols):
    room_layout=[['' for _ in range(cols)] for _ in range(rows)]
    course_header=['']*cols
    room_students=[]
    allocated_counts=defaultdict(int)

    def get_group_students(group_key):
        if group_key is None:
            return []
        return remaining_groups.get(group_key, [])

    def choose_replacement_group(other_group_key, active_keys):
        other_subject =(other_group_key[1] if other_group_key is not None else None)
        for key, students in remaining_groups.items():
            if not students: continue
            if key in active_keys: continue
            if other_subject is None or key[1]!=other_subject:
                return key
        for key, students in remaining_groups.items():
            if not students: continue
            if key in active_keys: continue
            return key
        return None

    active_group1 =group1_key
    active_group2 =group2_key
    for col in range(cols):
        if col%2==0:
            stream=1
            current_key=active_group1
            other_key=active_group2
        else:
            stream=2
            current_key=active_group2
            other_key=active_group1
        current_group =get_group_students(current_key)

        if not current_group:
            active_keys ={key for key in (active_group1, active_group2) if key is not None}
            replacement =choose_replacement_group(other_key, active_keys)
            if stream==1:
                active_group1=replacement
            else:
                active_group2=replacement
            current_key =replacement
            current_group =get_group_students(replacement)

        if not current_group:
            fallback_key =other_key
            fallback_group =get_group_students(fallback_key)
            if not fallback_group:
                break
            current_key =fallback_key
            current_group =fallback_group

        current_index =allocated_counts[current_key]
        course_header[col] =(f"{current_key[0]}({current_key[1]})")
        for row in range(rows):
            if current_index>=len(current_group):
                break
            student =current_group[current_index]
            room_layout[row][col] =student['USN']
            room_students.append(student)
            current_index += 1
        allocated_counts[current_key] = current_index

    return (room_layout, room_students, course_header, dict(allocated_counts))


def allocate_session(session_students, rooms_df):
    course_subject_groups =group_students_by_course_subject(session_students)
    room_pairs =create_room_pairs(course_subject_groups)
    remaining_groups ={
        key: list(students) for key, students in course_subject_groups.items()
    }
    room_allocations=[]
    all_allocated_students=[]

    room_pair_idx=0
    for _, room_row in rooms_df.iterrows():
        if room_pair_idx >= len(room_pairs):
            break
        room_name = room_row["RoomName"]
        rows=int(room_row["Rows"])
        cols=int(room_row["Cols"])
        group1_key, group2_key =room_pairs[room_pair_idx]
        group1_students = remaining_groups.get(group1_key, [])
        group2_students =(
            remaining_groups.get(group2_key, [])
            if group2_key
            else []
        )
        if not group1_students and not group2_students:
            room_pair_idx+=1
            continue

        (   room_layout,
            room_students,
            course_header,
            allocated_counts
        )= allocate_students_to_room(
            group1_key,
            group2_key,
            remaining_groups,
            rows,
            cols
        )

        if not room_students:
            continue
        room_allocations.append({
            "Room": room_name,
            "Layout": room_layout,
            "CourseHeader": course_header,
            "Students": room_students,
        })
        all_allocated_students.extend(room_students)

        for group_key, count in allocated_counts.items():
            if count<=0: continue
            remaining_groups[group_key] = (remaining_groups.get(group_key,[])[count:])

        group1_done = not remaining_groups.get(group1_key)
        group2_done = (group2_key is None or not remaining_groups.get(group2_key))
        if group1_done and group2_done:
            room_pair_idx +=1

    unallocated_students=[]
    for students in remaining_groups.values():
        unallocated_students.extend(students)

    return {
        "rooms": room_allocations,
        "allocated_students": all_allocated_students,
        "unallocated_students": unallocated_students,
    }