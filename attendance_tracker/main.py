import openpyxl
import os
import sys
import json
from collections import defaultdict

class Student:
    def __init__(self, firstname, lastname, student_id, eventname, tally=0):
        self.firstname = firstname
        self.lastname = lastname
        self.student_id = str(student_id)
        self.eventname = eventname
        self.tally = tally

    def to_dict(self):
        return {
            'firstname': self.firstname,
            'lastname': self.lastname,
            'student_id': self.student_id,
            'eventname': self.eventname,
            'tally': self.tally
        }

    @staticmethod
    def from_dict(data):
        return Student(data['firstname'], data['lastname'], data['student_id'], data['eventname'], data['tally'])


def list_files(directory):
    return [f for f in os.listdir(directory) if f.endswith('.xlsx')]

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
    students_file = os.path.join(base_dir, 'students.json')

    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f'Attendance_Sheets folder created at: {directory}')
    
    if not os.path.exists(students_file):
        with open(students_file, 'w') as f:
            json.dump({}, f)

    files = list_files(directory)
    choice2 = None
    while (choice2) != '4':
        choice2 = input('Press 1 to insert data into an event, press 2 to get final tallies for event, press 3 to list event names, 4 to exit the program (1/2/3/4): ')
        
        if choice2 == '1':
            choice3 = input('Press 1 to add a new event, press 2 to add to an existing event (1/2): ')
            students = load_students(students_file)
            
            if choice3 == '2':
                eventnames = list(set(student.eventname for student in students.values()))
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

            # Check if the file exists
            if not files:
                print('No files found please add excel files to the Attendance_Sheets folder')
                delay = input('Press The Enter Key to exit the program')
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
            for row in sheet.iter_rows(min_row=6, min_col=1, max_col=9):
                if all(cell.value is None for cell in row):
                    break
                student_id = row[0].value
                firstname = row[1].value
                lastname = row[2].value

                if student_id and eventname:
                    key = (str(student_id), eventname)
                    if key not in students:
                        students[key] = Student(firstname, lastname, student_id, eventname, 1)
                    else:
                        students[key].tally += 1
            
            # Save updated students data
            save_students(students_file, students)
            print('Data saved')

        elif choice2 == '2':
            students = load_students(students_file)
            eventnames = list(set(student.eventname for student in students.values()))
            print('Existing event names:')
            for idx, event in enumerate(eventnames):
                print(f"{idx + 1}: {event}")
            event_choice = int(input('Select an event by number: ')) - 1

            if event_choice < 0 or event_choice >= len(eventnames):
                print('Invalid choice')

            eventname = eventnames[event_choice]
            tally1 = int(input('Enter the desired number of events attended: '))
            
            for student in students.values():
                if student.eventname == eventname and student.tally == tally1:
                    print(f"{student.firstname} {student.lastname} ({student.student_id}): {student.tally} events")

        elif choice2 == '3':
            students = load_students(students_file)
            eventnames = list(set(student.eventname for student in students.values()))
            print('Existing event names:')
            for event in eventnames:
                print(event)

        elif choice2 == '4':
            sys.exit()
if __name__ == '__main__':
    main()