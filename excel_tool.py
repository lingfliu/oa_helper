# -*- coding: utf-8 -*-

import xlrd
import xlwt

file = '17-20培养目标达成度-直接评价.xlsx'

data = xlrd.open_workbook(file)

for tab in data.sheets():
    print(tab.name)
    nrows = tab.nrows

    for r in range(nrows):
        entries = tab.row(r)




