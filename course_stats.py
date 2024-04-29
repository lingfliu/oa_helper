# -*- coding: utf-8 -*-
import re

import xlrd
import xlwt
import numpy as np
from str_tools import is_number, is_empty

dir = 'D:\\nutshare\\lingfliu\\ecjtu\\2024-本科教学审核评估\\'

file_mission = '17.1-23.2教学任务书.xlsx'

data_mission = xlrd.open_workbook(dir + file_mission)

tab_names = data_mission.sheet_names()

tab = data_mission.sheet_by_index(0)

print(tab.name)
ncols = tab.ncols
nrows = tab.nrows

col_names = tab.row(0)[:]

col_names = [col_names[i].value for i in range(ncols) ]

semester_idx = -1
class_name_idx = -1
lec_idx = -1
course_name_idx = -1
for idx, cname in enumerate(col_names):
    if re.match(r'^学期', cname):
        semester_idx = idx
    if re.match(r'^合班信息', cname):
        class_name_idx = idx
    if re.match(r'^主讲教师$', cname):
        lec_idx = idx
    if re.match(r'^课程名称$', cname):
        course_name_idx = idx
print(col_names)
print('学期信息在第%d列' % (semester_idx+1))
print('合班信息在第%d列' % (class_name_idx+1))
print('课程名称在第%d列' % (course_name_idx+1))
print('主讲老师在第%d列' % (lec_idx+1))

courses = []
row_idx = [i+1 for i in range(nrows-1)]
for row in row_idx:
    # if class name match any of '物联网工程'
    class_name = str(tab.row(row)[class_name_idx].value)
    if re.match(r'物联网工程', class_name):
        if not '2016' in class_name:
            courses.append({
                'semester': tab.row(row)[semester_idx].value,
                'class_name': tab.row(row)[class_name_idx].value,
                'course_name': tab.row(row)[course_name_idx].value,
                'lec': tab.row(row)[lec_idx].value})
    if is_empty(tab.row(row)[semester_idx].value):
        print('学期信息为空')
    if is_empty(tab.row(row)[class_name_idx].value):
        print('合班信息为空')
    if is_empty(tab.row(row)[lec_idx].value):
        print('主讲教师信息为空')

# print(courses)
workbook = xlwt.Workbook()
workbook.add_sheet('物联网工程')
sheet = workbook.get_sheet(0)
sheet.write(0,0,'学期')
sheet.write(0,1, '合班信息')
sheet.write(0,2, '课程名称')
sheet.write(0,3, '主讲教师')
for i, course in enumerate(courses):
    sheet.write(i+1, 0, course['semester'])
    sheet.write(i+1, 1, course['class_name'])
    sheet.write(i+1, 2, course['course_name'])
    sheet.write(i+1, 3, course['lec'])
workbook.save(dir + '物联网工程课程信息.xls')

