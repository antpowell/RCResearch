import os
import re
import json
from openpyxl import Workbook


class Course:
    def __init__(self, instructor, course_code, course_name, section, day,
                 time, credits, building, room, college, department,
                 term, crn):
        self.instructor = instructor
        self.course_code = course_code
        self.course_name = course_name
        self.section = section
        self.day = day
        self.time = time
        self.credits = credits
        self.building = building
        self.room = room
        self.college = college
        self.department = department
        self.term = term
        self.crn = crn

    def print_short_form(self):
        print("Course: {} Section: {}\nInstructor: {}\n".format(self.course_code, self.section,
                                                                self.instructor))

    def print_all_details(self):
        print(
            "CRN:{}\n"
            "Instructor:{}\nCourse & Section:{} {}"
            "\nDay\Time:{}\{}\nCredits:{}\nBuilding Room:{}{}"
            "\nCollege & Department{} {}\nTerm: {}".format(self.crn,
                                                           self.instructor, self.course_code,
                                                           self.section, self.day, self.time,
                                                           self.credits, self.building, self.room,
                                                           self.college, self.department,
                                                           self.term))

    def __str__(self):
        return ("CRN:{}\n"
                "Instructor:{}\nCourse & Section:{} {}"
                "\nDay\Time:{}\{}\nCredits:{}\nBuilding Room:{}{}"
                "\nCollege & Department{} {}\nTerm: {}\n\n".format(self.crn,
                                                                   self.credits, self.building,
                                                                   self.room, self.college,
                                                                   self.department,
                                                                   self.instructor,
                                                                   self.course_code, self.section,
                                                                   self.day, self.time,
                                                                   self.term))

    def __iter__(self):
        # if(self.instructor != ''):
            yield 'instructor', self.instructor
            yield 'course_name', self.course_name
            yield 'course_code', self.course_code
            yield 'section', self.section
            yield 'day', self.day
            yield 'time', self.time
            yield 'credits', self.credits
            yield 'building', self.building
            yield 'room', self.room
            yield 'college', self.college
            yield 'department', self.department
            yield 'term', self.term
            yield 'crn', self.crn

    def  resetAll(self):
        course_short = ''
        course_long = ''
        instructor = ''
        department = ''
        building = ''
        credits = ''
        section = ''
        college = ''
        time = ''
        room = ''
        term = ''
        crn = ''
        day = ''

arr = open('C:\\Users\powel\Desktop\Roll Call\\fall2017 - Final.txt').read().split('\n')

# print(arr[16])
course_short = ''
course_long = ''
instructor = ''
department = ''
building = ''
credits = ''
section = ''
college = ''
time = ''
room = ''
term = ''
crn = ''
day = ''

courses = {}
courses_high_lvl = {}

# index = 0
for index, line in enumerate(arr):
    if re.search('Student Name', line):
        heading_line_number = (index - 13)
        while heading_line_number <= index:
            if re.search('CRN', arr[heading_line_number]):
                crn = arr[heading_line_number + 1].split()[0]
                term = arr[heading_line_number + 1].split()[1]
                course_short = ' '.join(arr[heading_line_number + 1].split()[2:4])
                section = arr[heading_line_number + 1].split()[4]
                course_long = ' '.join(arr[heading_line_number + 1].split()[5:-2])
                credits = arr[heading_line_number + 1].split()[-2]
                # print('{} {} {}{} {} {}'.format(crn, term, course_short, section, course_long,
                #                                 credits))

            if re.search('Lecture', arr[heading_line_number]) or re.search('Lab', arr[heading_line_number]) :
                if len(arr[heading_line_number].split()) > 5:
                    last = arr[heading_line_number].split()[0]
                    first = ' '.join(arr[heading_line_number].split()[1:-5])
                    instructor = '{} {}'.format(first, last)
                    day = arr[heading_line_number].split()[-4]
                    time = arr[heading_line_number].split()[-3]
                    building = arr[heading_line_number].split()[-2]
                    room = arr[heading_line_number].split()[-1]
                    # print('{}:{}/{} {} {}'.format(instructor, day, time, building, room))
                elif len(arr[heading_line_number].split()) > 0:
                    instructor = "Not Found"
                    day = arr[heading_line_number].split()[1]
                    time = arr[heading_line_number].split()[2]
                    building = arr[heading_line_number].split()[3]
                    room = arr[heading_line_number].split()[4]
                else:
                    print('Error: line contains no instructor information')
            if re.search('COLLEGE:', arr[heading_line_number]):
                college = ' '.join(arr[heading_line_number].split()[1:-1])
                # print(arr[heading_line_number].split())
                # print(college)
            if re.search('DEPARTMENT:', arr[heading_line_number]):
                department = ' '.join(arr[heading_line_number].split()[1:])
                # print(department)
            if instructor != '' and heading_line_number == (index-1):
                course = Course(instructor, course_short, course_long, section, day, time,
                                credits, building,
                                room, college, department, term, crn)
                courses["{} {}".format(course.course_code, course.section)] = course.__dict__
                course.resetAll()
            heading_line_number += 1
# print(courses)
courses_high_lvl['courses'] = courses

with open('full_course_information.json', 'w') as outfile:
    json.dump(courses_high_lvl, outfile)
# print(arr)
