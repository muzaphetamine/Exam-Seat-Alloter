A local, offline exam seat allocation system that generates Excel-based seating arrangements from student, room, and schedule data.
The system is designed to reduce opportunities for students to copy by grouping students by course and subject and alternating different subject groups across room columns.

Requirements  
-Python 3.10+  
-Flask  
-Pandas  
-OpenPyXL  
-xlrd  
-Werkzeug  
Install dependencies with:  
pip install -r requirements.txt  


Running the Application  
-Run this from the project root: python -m backend.app  
-Open the localhost link  
-And use the application


Two modes available:  
1) SESSION MODE
Specifically built for college format
Uses 2 varieties of input files
-Room file(s): consists of room capacity data
<img width="755" height="134" alt="e1" src="https://github.com/user-attachments/assets/dcc967f2-dc24-4daf-925a-33be266d01c8" />
-Session file(s): each file consists student data for one exam session only
<img width="335" height="143" alt="e2" src="https://github.com/user-attachments/assets/e2c7c756-7e9e-4abf-ab70-8d8685160cb0" />

2) CENTRALIZED MODE
Built for general use case
Uses 3 varieties of input files
-Room file(s): consists of room capacity data
<img width="755" height="134" alt="e1" src="https://github.com/user-attachments/assets/dcc967f2-dc24-4daf-925a-33be266d01c8" />
-Student file(s): consists general data of students
<img width="535" height="184" alt="e3" src="https://github.com/user-attachments/assets/0a12a566-dc5c-479a-8867-c91911807cc8" />
-Schedule file(s): consists exam schedule data
<img width="554" height="161" alt="e4" src="https://github.com/user-attachments/assets/608cafee-9979-426a-8244-055e72cedb64" />


ALLOCATION LOGIC:
Students are grouped using (Course, SubjectCode)
Groups are initially paired to encourage different courses and different subjects in the same room.
For a pair of groups:
A B A B
A B A B
Students are placed column by column.
If one group finishes before the other, another remaining group can replace that stream.
For example:
A B A B
C B C B
C B C B
Replacement groups preferably use a different subject from the other active stream.
If only one group remains, the remaining seats are filled normally.
Students who cannot be placed because the available rooms are insufficient are reported as unallocated.


CONFLICT DETECTION:
Before allocation, the system checks whether the same USN appears multiple times within a session.
Detected conflicts are written to a separate Excel file under:
Conflicts/
Conflicting sessions are not allocated.


OUTPUT:
For each successfully processed session, the application generates an Excel workbook containing:

Room Allocation
A master sheet containing:
- Room
- Branch
- Subject code
- Subject name
- Semester
- Allocated USNs
- Total students
  
Room Sheets
Each allocated room gets its own sheet containing:
- Room name
- Column/course headers
- Seating layout
- Students assigned to the room
  
Session Summary:
- Total students
- Allocated students
- Unallocated students
- Average room capacity
- Total rooms
- Extra rooms needed
- Course distribution
  
Unallocated Students
Generated only when students remain unallocated.


Configuration:
Application settings are stored in: backend/config.py
