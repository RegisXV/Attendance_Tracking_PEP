import openpyxl
import os
import sys
import json
from collections import defaultdict

class Student:
    def __init__(self, firstname, lastname, student_id, studentemail, eventname, tally=0):
        self.firstname = firstname
        self.lastname = lastname
        self.student_id = str(student_id)
        self.studentemail = studentemail
        self.eventname = eventname
        self.tally = tally

    def to_dict(self):
        return {
            'firstname': self.firstname,
            'lastname': self.lastname,
            'student_id': self.student_id,
            'studentemail': self.studentemail,
            'eventname': self.eventname,
            'tally': self.tally
        }

    @staticmethod
    def from_dict(data):
        return Student(data['firstname'], data['lastname'], data['student_id'], data['studentemail'], data['eventname'], data['tally'])


def list_files(directory):
    return [f for f in os.listdir(directory) if f.endswith('.xlsx')]

def list_semesters(semester_directory):
    return [f for f in os.listdir(semester_directory) if f.endswith('.json')]

def load_students(filename):
    if os.path.exists(filename) and os.path.getsize(filename) > 0:  # Check if file is not empty
        with open(filename, 'r') as f:
            data = json.load(f)
            return {tuple(k.split(',')): Student.from_dict(v) for k, v in data.items()}
    return {}

def save_students(filename, students):
    with open(filename, 'w') as f:
        data = {','.join(map(str, k)): v.to_dict() for k, v in students.items()}  # Convert keys to strings
        json.dump(data, f)

def main():
    base_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
    directory = os.path.join(base_dir, 'Attendance_Sheets')
    semester_directory = os.path.join(base_dir, 'Semesters')

    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f'Attendance_Sheets folder created at: {directory}')

    if not os.path.exists(semester_directory): 
        os.makedirs(semester_directory)
        print(f'Semesters folder created at: {semester_directory}')

    
    semester_selected = None
    files = list_files(directory)
    semester_list = list_semesters(semester_directory)
    while (semester_selected) != '1':

        Semester_Choice = input('Press 1 to select an existing semester, press 2 to create a new semester (1/2): ')

        if Semester_Choice == '2':
            semester= input('Insert a semester (e.g. Fall 2021): ')
            students_file = os.path.join(semester_directory, f'{semester}_Students.json')
            if not os.path.exists(students_file):
                with open(students_file, 'w') as f:
                    json.dump({}, f)
                print(f'File created at: {students_file} initializaing program.')
                semester_selected = '1'
            else:
                semester_selected = '1'
                print(f'{semester} has been selected initializing program.')

        elif Semester_Choice == '1':
            if not semester_list:
                print('No semesters found please create a new semester')
                sys.exit()
            print('Existing semesters:')
            for idx, semester in enumerate(semester_list):
                print(f"{idx + 1}: {semester}")
            semester_choice = int(input('Select a semester by number: ')) - 1

            if semester_choice < 0 or semester_choice >= len(semester_list):
                print('Invalid choice')
                continue

            semester = semester_list[semester_choice]
            students_file = os.path.join(semester_directory, f'{semester}')
            semester_selected = '1'
            print(f'{semester} has been selected initializing program.')


    while True:
        files = list_files(directory)
        if not files:
            print(f'No files found in {directory}. Please add Excel files to this folder and run the program again.')
            sys.exit()

        choice2 = input('Press 1 to insert data into an event, press 2 to get final tallies for event, press 3 to list event names, press 4 to reset Attendance sheet folder, press 5 to exit the program (1/2/3/4/5): ')
        
        if choice2 == '1':
            choice3 = input('Press 1 to add a new event, press 2 to add to an existing event (1/2): ')
            students = load_students(students_file)
            
            if choice3 == '2':
                eventnames = list(set(student.eventname for student in students.values()))
                if not eventnames:
                    print('No Events found please add an event')
                    continue

                print('Existing event names:')
                for idx, event in enumerate(eventnames):
                    print(f"{idx + 1}: {event}")
                event_choice = int(input('Select an event by number: ')) - 1

                if event_choice < 0 or event_choice >= len(eventnames):
                    print('Invalid choice')
                    sys.exit()

                eventname = eventnames[event_choice]
            elif choice3 == '1':
                eventname = input('Enter the name of the new event: ')

            files = list_files(directory)
            if not files:
                print(f'No files found in {directory}. Please add Excel files to this folder and run the program again.')
                sys.exit() 

            print('Available files:')
            for idx, file in enumerate(files):
                print(f"{idx + 1}: {file}")

            choice = int(input('Select a file by number: ')) - 1

            if choice < 0 or choice >= len(files):
                print('Invalid choice')
                sys.exit()

            file_path = os.path.join(directory, files[choice])

            # Load the workbook
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active

            # Print the values of the first row
            for row in sheet.iter_rows(min_row=7, min_col=1, max_col=9):
                if all(cell.value is None for cell in row):
                    break
                student_id = row[8].value
                firstname = row[0].value
                lastname = row[1].value
                studentemail = row[2].value

                if student_id and eventname:
                    key = (str(student_id), eventname)
                    if key not in students:
                        students[key] = Student(firstname, lastname, student_id, studentemail, eventname, 1)
                    else:
                        students[key].tally += 1
            
            # Save updated students data
            save_students(students_file, students)
            print('Data saved')

        elif choice2 == '2':
            students = load_students(students_file)
            eventnames = list(set(student.eventname for student in students.values()))
            if not eventnames:
                    print('No Events found please add an event')
                    continue
            print('Existing event names:')
            for idx, event in enumerate(eventnames):
                print(f"{idx + 1}: {event}")
            event_choice = int(input('Select an event by number: ')) - 1

            if event_choice < 0 or event_choice >= len(eventnames):
                print('Invalid choice please select an existing event or create a new one')
                continue
            eventname = eventnames[event_choice]
            tally1 = int(input('Enter the desired number of events attended: '))
            
            for student in students.values():
                if student.eventname == eventname and student.tally == tally1:
                    print(f"Name-{student.firstname} {student.lastname} ID-({student.student_id}) Email-({student.studentemail}): {student.tally} Events Attended")

        elif choice2 == '3':
            students = load_students(students_file)
            
            eventnames = list(set(student.eventname for student in students.values()))
            if not eventnames:
                    print('No Events found please add an event')
                    continue
            print('Existing event names:')
            for event in eventnames:
                print(event)

        elif choice2 == '4':
            Reset=input('Are you sure you want to reset the Attendance_Sheets folder? (Y/N): ')
            if Reset == 'Y':
                Reset1=input('Please Type Reset with capital R to confirm: ')
                if Reset1 == 'Reset':
                    for file in os.listdir(directory):
                        file_path = os.path.join(directory, file)
                        os.remove(file_path)
                    print('All files removed from Attendance_Sheets folder')
                elif Reset1 != 'Reset':
                    print('Reset not confirmed')
                    continue
            elif Reset == 'N':
                print('Reset not confirmed')
                continue

        elif choice2 == '5':
            sys.exit()
if __name__ == '__main__':
    main()