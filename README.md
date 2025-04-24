# Timetable Automation System

## Overview
This Timetable Automation System generates academic timetables for educational institution IIITDWD. The system intelligently schedules courses, allocates appropriate rooms, manages faculty assignments, and handles various constraints like breaks and section divisions, all while avoiding scheduling conflicts.

## Features
- **Automated Scheduling**: Generates complete timetables for all departments and semesters
- **Room Allocation**: Assigns rooms based on capacity and type (lecture rooms vs. labs)
- **Faculty Assignment**: Prevents faculty scheduling conflicts
- **Break Management**: Incorporates morning, lunch, and afternoon breaks
- **Section Division**: Handles sections and batches for labs and practical sessions
- **Constraint Handling**: Respects time slots, room availability, and faculty schedules
- **Excel Output**: Generates formatted Excel timetables with color-coding for different session types

## Requirements
- Python 3.8+
- Pandas
- OpenPyXL
- Random (standard library)
- Datetime (standard library)

## Input Files
The system requires the following CSV files:

1. **courses.csv**
   - Contains course details with attributes:
   - `COURSE_ID, DEPARTMENT, SEMESTER, COURSE_CODE, COURSE_NAME, L, T, P, S, C, SEMESTER_TYPE, FACULTY_ID, COMBINED, CAPACITY`
   - L, T, P represent Lecture, Tutorial, Practical hours respectively

2. **rooms.csv**
   - Contains room information with attributes:
   - `id, room no, capacity, room type`
   - Room types include LECTURE_ROOM, COMPUTER_LAB, HARDWARE_LAB, SEATER_120, SEATER_240

3. **faculty.csv**
   - Contains faculty information with attributes:
   - `faculty_id, faculty_name`

4. **electives.csv**
   - Contains elective course details with attributes:
   - `elective, elective_name, faculty_id, faculty_name, semester`

## Usage

1. Ensure all input CSV files are in the same directory as the script
2. Run the script:
   ```
   python timetable_generator.py
   ```
3. The output will be saved as `timetables.xlsx` in the current directory

## Scheduling Rules

- **Lectures**: 1.5 hours (3 slots of 30 minutes each)
- **Tutorials**: 1 hour (2 slots of 30 minutes each)
- **Labs**: 2 hours (4 slots of 30 minutes each)
- **Breaks**:
  - Morning: 10:30 - 11:00
  - Lunch: 13:30 - 14:30
  - Snacks: 16:30 - 17:00
- **Room Assignment**:
  - Lectures and tutorials are scheduled in lecture rooms
  - Labs are scheduled in computer or hardware labs
  - Room capacity must be sufficient for the course

## Output Format

The system generates an Excel file with multiple sheets, one for each department-semester-section combination. Each timetable shows:

- Days of the week (Monday to Friday)
- Time slots from 9:00 to 18:30
- Color-coded sessions:
  - Lectures: Light purple
  - Tutorials: Light pink
  - Labs: Light green
  - Breaks: Light gray

## Constraints Handled

- Faculty cannot teach multiple courses simultaneously
- A room cannot be allocated to multiple courses at the same time
- Courses are scheduled across different days to avoid overburdening
- Breaks are respected and no classes are scheduled during break times
- Labs are divided into batches if course capacity exceeds room capacity

## Example
A scheduled course will display:
- Course code
- Session type (LEC/TUT/LAB)
- Classroom number
- Each session spans multiple slots according to its duration

## Contributors
Team Time Table Traume:
- Ankit K - 23bcs015
- K Likith - 23bcs061
- K Avaneesh - 23bcs065
- KV Modak - 23bcs067