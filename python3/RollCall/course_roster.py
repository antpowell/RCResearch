import re
import json


class CourseRoster:
    def __init__(self, course, students=[]):
        self.course = course
        self.students = students

    def __str__(self):
        return '{} {}'.format(self.course, self.students)

    def __iter__(self):
        yield 'course', self.course
        yield 'students', self.students

    def put_course(self, course):
        self.course = course

    def put_students(self, students):
        self.students = students


document_lines = open('C:\\Users\powel\Desktop\Roll Call\\fall2017 - Final.txt').read().split('\n')
students = []
course_code = ''
course_roster = CourseRoster
all_course_rosters = {}
all_course_rosters_high_lvl = {}
for index, line in enumerate(document_lines):
    '''
    When line contains CRN move down one line and get the Subject Course and Section and
    place 
    '''

    if re.search('CRN', line):
        course_code = ' '.join((document_lines[index + 1].split()[2:5]))
    elif re.search('Student Name', line):
        start_of_students_line = (index + 2)
        while not re.search('TSU Production \(SUN\)', document_lines[start_of_students_line]):
            if document_lines[start_of_students_line] == '' or re.search("Instructor's Signature",
                                                                         document_lines[
                                                                             start_of_students_line]) or re.search(
                    '______________________________',
                    document_lines[start_of_students_line]) or re.search('CONTINUED ON NEXT PAGE',
                                                                         document_lines[
                                                                             start_of_students_line]):
                pass
            else:
                students.append(document_lines[start_of_students_line])

            if re.search("Instructor's Signature", document_lines[start_of_students_line]):
                if not re.search('CONTINUED ON NEXT PAGE',
                                 document_lines[start_of_students_line + 5]):
                    course_roster = CourseRoster(course_code, students)
                    students = []

            start_of_students_line += 1
            # elif re.search('TSU Production \(SUN\)', line) and index != 0:
            #     if not re.search('(CONTINUED ON NEXT PAGE)', document_lines[index - 1]):
            #         students = []
        # print(course_roster)
        all_course_rosters[course_code] = course_roster.__dict__
        all_course_rosters_high_lvl['course_rosters'] = all_course_rosters
# print(all_course_rosters)

with open('course_roster.json', 'w') as outfile:
    json.dump(all_course_rosters_high_lvl, outfile)
