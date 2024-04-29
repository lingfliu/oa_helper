# -*- coding: utf-8 -*-

import xlrd
import xlwt
import numpy as np
from str_tools import is_number, is_empty
import os
import re

p = os.path.abspath(os.path.dirname(__file__))
dir = 'D:\\nutshare_lingfeng\\ecjtu\\2024-本科教学审核评估\\'

file_score = '17-20培养目标达成度-直接评价-v2.xlsx'
file_weight = '17-20培养目标达成度-直接评价权重-v2.xlsx'


data_weight = xlrd.open_workbook(dir + file_weight)

tab_names = data_weight.sheet_names()
weights = []

courses = []

pmajor = []
pmin = []

for tab in data_weight.sheets():

    print(tab.name)
    ncols = tab.ncols
    nrows = tab.nrows

    p1 = tab.row(0)[1:]
    p1 = [int(p.value) for p in p1]
    p2 = tab.row(1)[1:]
    p2 = [int(p.value) for p in p2]

    pmajor = p1
    pmin = p2

    course_list = []
    for ir in range(nrows-2):
        course = tab.row(ir+2)[0].value
        course_list.append(course)

    courses.append(course_list)

    weight = np.zeros((len(p2), nrows-2))
    for ic in range(len(p2)):
        for ir in range(nrows-2):
            w = tab.row(ir+2)[ic+1].value
            if is_empty(w):
                w = 0
            else:
                vs = w.split('/')
                if len(vs) > 1:
                    weight[ic,ir] = float(vs[1])
                    # print('weight=', weight[ic,ir])

    # check the sum of each column
    print('权重核算', tab.name,np.sum(weight,1))

    weights.append(weight)

data_score = xlrd.open_workbook(dir + file_score)
scores = []
for tab in data_score.sheets():
    # print(tab.name)
    ncols = tab.ncols
    nrows = tab.nrows

    p1 = tab.row(0)[1:]
    p1 = [int(p.value) for p in p1]
    p2 = tab.row(1)[1:]
    p2 = [int(p.value) for p in p2]

    score = np.zeros((len(p2), nrows-2))
    for ic in range(len(p2)):
        for ir in range(nrows-2):
            v = tab.row(ir+2)[ic+1].value
            # check w is number
            if is_number(v):
                score[ic,ir] = float(v)
            else:
                score[ic,ir] = 0
    # check the sum of each column
    # print(np.average(score,1))
    scores.append(score)

for idx in range(len(weights)):
    weight = weights[idx]
    score = scores[idx]
    score_avg_weighted = np.average(weight*score,1)

    weight_sum = np.sum(weight,1)
    print('权重和', tab_names[idx])
    for i in range(len(score_avg_weighted)):
        print(pmajor[i], pmin[i], weight_sum[i])
    #
    # print('达成度平均分', tab_names[idx])
    # for i in range(len(score_avg_weighted)):
    #     print(pmajor[i], pmin[i], score_avg_weighted[i])

dir_score = 'D:\\nutshare_lingfeng\\ecjtu\\2024-本科教学审核评估\\物联网2020-工程认证材料\\成绩'
# dir_score = 'D:\\nutshare_lingfeng\\ecjtu\\2024-本科教学审核评估\\物联网2020-工程认证材料\\成绩'
files_score = ['2017物联网学生最终成绩.xls', '2018.xls', '2019.xls', '2020.xls']

score_avg = []
for idx, course in enumerate(courses):
    data_score = xlrd.open_workbook(dir_score + '\\' + files_score[idx])
    tab = data_score.sheet_by_index(0)
    avg = []
    ncols = tab.ncols
    nrows = tab.nrows

    col_names = tab.row(0)[:]
    col_names = [col_names[i].value for i in range(ncols)]

    idx_cname = -1
    idx_score = -1

    for idx, cname in enumerate(col_names):
        if re.match(r'^课程名称', cname):
            idx_cname = idx
        if re.match(r'^成绩', cname):
            idx_score = idx

    for c in course:
        cname = c.split('(')[0]

        cnt = 0
        score_sum = 0
        for ir in range(nrows-2):
            txt = str(tab.row(ir+2)[idx_cname].value)
            if cname in txt:
                score = str(tab.row(ir+2)[idx_score].value)
                if score == '优秀':
                    score_sum += 95
                elif score == '良好':
                    score_sum += 85
                elif score == '中等':
                    score_sum += 75
                elif score == '及格':
                    score_sum += 65
                elif score == '不及格':
                    score_sum += 50
                elif score == '合格':
                    score_sum += 80
                elif score == '不合格':
                    score_sum += 50
                elif score == '缺考':
                    score_sum += 0
                elif score == '缓考':
                    score_sum += 0
                elif score == '取消资格':
                    score_sum += 0
                elif score == '免修':
                    score_sum += 90
                else:
                    score_sum += float(score)
                cnt += 1

        if cnt == 0: # 军训
            avg.append({
                'cname': cname,
                'score_avg': 80})
        else:
            avg.append({
                'cname': cname,
                'score_avg': score_sum/cnt})

    score_avg.append(avg)

workbook = xlwt.Workbook()
years = ['2017', '2018', '2019', '2020']
for i, avg in enumerate(score_avg):
    workbook.add_sheet(years[i])
    sheet = workbook.get_sheet(i)
    sheet.write(0, 0, '课程名称')
    sheet.write(0, 1, '平均成绩')
    for j, a in enumerate(avg):
        sheet.write(j + 1, 0, a['cname'])
        sheet.write(j + 1, 1, a['score_avg'])
workbook.save(dir + '17-20-物联网工程课程平均分.xls')

