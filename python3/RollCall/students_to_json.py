import re
import json


class Student:
    def __init__(self, first_name, last_name, t_num, lvl, major, classification):
        self.first_name = first_name
        self.last_name = last_name
        self.t_num = t_num
        self.lvl = lvl
        self.major = major
        self.classification = classification

    def print_short(self):
        print('{}: {},{}'.format(self.t_num, self.first_name, self.last_name))

    def __str__(self):
        return ('{},{}: {}\n'
                'Level: {}, Major: {}\n'
                'Classification: {}'.format(self.first_name, self.last_name, self.t_num,
                                            self.lvl, self.major, self.classification))

    def __iter__(self):
        yield 'first_name', self.first_name
        yield 'last_name', self.last_name
        yield 't_num', self.t_num
        yield 'lvl', self.lvl
        yield 'major', self.major
        yield 'classification', self.classification


arr = open('C:\\Users\powel\Desktop\Roll Call\\fall2017 - Final.txt').read().split('\n')

# print(arr[16])
classification = ''
first_name = ''
last_name = ''
major = ''
t_num = ''
lvl = ''

students = {}
students_high_lvl = {}

index = 0
# for i, section in enumerate(arr[9296].split()):
#     print('{}: {}'.format(i, section))
# for i, section in enumerate(arr[16].split()):
#     print('{}: {} '.format(i, section))
# print(arr[9296].split()[-10])
# print(arr[16].split()[-10])

for index, line in enumerate(arr):
    if re.search('Student Name', line):
        start_of_students_line = (index + 2)
        while arr[start_of_students_line] != "":
            last_name = arr[start_of_students_line].split()[1]
            first_name = ' '.join(arr[start_of_students_line].split()[2:-10])
            t_num = arr[start_of_students_line].split()[-10]
            lvl = arr[start_of_students_line].split()[-9]
            major = arr[start_of_students_line].split()[-8]
            classification = arr[start_of_students_line].split()[-7]
            start_of_students_line += 1
            student = Student(first_name, last_name, t_num, lvl, major, classification)
            print(student)
            students[student.t_num] = student.__dict__
            students_high_lvl['students'] = students

with open('student_info.json', 'w') as outfile:
    json.dump(students_high_lvl, outfile)
