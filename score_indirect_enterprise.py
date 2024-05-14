# -*- coding: utf-8 -*-

import xlrd
import xlwt
import numpy as np
from str_tools import is_number, is_empty
import os
import re

p = os.path.abspath(os.path.dirname(__file__))
dir = 'D:\\nutshare_lingfeng\\ecjtu\\2024-本科教学审核评估\\'

file_score = '266321610_按文本_华东交大物联网工程专业毕业要求达成度企业调研_9_9.xlsx'


data_score = xlrd.open_workbook(dir + file_score)
scores = []
tab = data_score.sheets()[0]
ncols = tab.ncols
nrows = tab.nrows
"""TODO 样本数太少，统一算，后续要补充"""
years = ['2023', '2022', '2021', '2020']

scores = np.zeros((ncols-8, 1))
cnt = 0
for ir in range(nrows-1):
    cnt += 1
    for ic in range(8, ncols):
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
        scores[ic - 8] += v

scores = scores / cnt


workbook = xlwt.Workbook()
workbook.add_sheet('scores')
sheet = workbook.get_sheet(0)
for j in range(len(scores)):
    sheet.write(0, j, scores[j,0])
workbook.save(dir + '企业间接评价.xls')
