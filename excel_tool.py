# -*- coding: utf-8 -*-

import xlrd
import xlwt
import numpy as np
from str_tools import is_number

dir = './'

file_score = '17-20培养目标达成度-直接评价.xlsx'
file_weight = '17-20培养目标达成度-直接评价权重.xlsx'

data_weight = xlrd.open_workbook(dir + file_weight)

tab_names = data_weight.sheet_names()
weights = []

pmajor = []
pmin = []
for tab in data_weight.sheets():
    # print(tab.name)
    ncols = tab.ncols
    nrows = tab.nrows

    p1 = tab.row(0)[1:]
    p1 = [int(p.value) for p in p1]
    p2 = tab.row(1)[1:]
    p2 = [int(p.value) for p in p2]

    pmajor = p1
    pmin = p2

    weight = np.zeros((len(p2), nrows-2))
    for ic in range(len(p2)):
        for ir in range(nrows-2):
            w = tab.row(ir+2)[ic+1].value
            if w == 'H':
                weight[ic,ir] = 0.2
            elif w == 'M':
                weight[ic,ir] = 0.1
            else:
                weight[ic,ir] = 0

    # check the sum of each column
    # print(np.sum(weight,1))

    weights.append(weight)

data_score = xlrd.open_workbook(dir + file_score)
scores = []
for tab in data_score.sheets():
    print(tab.name)
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

    print('达成度平均分', tab_names[idx])
    for i in range(len(score_avg_weighted)):
        print(pmajor[i], pmin[i], score_avg_weighted[i])