# -*- coding: utf-8 -*-

import xlrd
import xlwt
import numpy as np
from str_tools import is_number, is_empty
import os
import re

p = os.path.abspath(os.path.dirname(__file__))
dir = 'D:\\nutshare_lingfeng\\ecjtu\\2024-本科教学审核评估\\'

file_score = '266306948_按文本_华东交大物联网工程毕业要求达成度毕业生调研_58_58.xlsx'


data_score = xlrd.open_workbook(dir + file_score)
scores = []
tab = data_score.sheets()[0]
ncols = tab.ncols
nrows = tab.nrows
years = ['2023', '2022', '2021', '2020']

score_list = []
for y in years:
    scores = np.zeros((ncols-7, 1))
    cnt = 0
    for ir in range(nrows-1):
        vy = tab.row(ir+1)[6].value
        if y == vy:
            cnt += 1
            for ic in range(7, ncols):
                v = tab.row(ir+1)[ic].value
                if v == '5':
                    v = 100
                elif v == '4':
                    v = 85
                elif v == '3':
                    v = 65
                elif v == '2':
                    v = 45
                elif v == '1':
                    v = 25
                scores[ic - 7] += v

    if cnt == 0:
        print('No data for year %s' % y)
    else:
        scores = scores / cnt
    score_list.append(scores)


workbook = xlwt.Workbook()
workbook.add_sheet('scores')
sheet = workbook.get_sheet(0)
for i, year in enumerate(years):
    scores = score_list[i]
    sheet.write(i, 0, year)
    for j in range(len(scores)):
        sheet.write(i, j + 1, scores[j,0])
workbook.save(dir + '毕业生间接评价.xls')
