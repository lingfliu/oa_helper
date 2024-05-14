# -*- coding: utf-8 -*-

import xlrd
import xlwt
import numpy as np
from str_tools import is_number, is_empty
import os

# get root path of D:
# dir = 'D:\\nutshare\\lingfliu\\ecjtu\\2024-本科教学审核评估\\'
dir = 'D:\\nutshare_lingfeng\ecjtu\\2024-本科教学审核评估\\'


'''统计达成度评价'''

file_score_avg = '物联网工程课程平均分.xls' # failsafe average score

file_score = '17-20培养目标达成度-直接评价-v3.xlsx'

file_weight = '17-20培养目标达成度-直接评价权重-v2.xlsx'

data_weight = xlrd.open_workbook(dir + file_weight)

tab_names = data_weight.sheet_names()
weights = []

pmajor = [] # 指标点
pmin = [] # 指标分界点
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
    # print('权重核算', tab.name,np.sum(weight,1))

    weights.append(weight)

data_score_avg = xlrd.open_workbook(dir + file_score_avg)
score_avgs = []
for tab in data_score_avg.sheets():
    # print(tab.name)
    nrows = tab.nrows
    score_avg = {}

    for ir in range(nrows-1):
        course_name = tab.row(ir + 1)[0].value
        score = tab.row(ir + 1)[1].value
        score_avg[course_name] = float(score)
    score_avgs.append(score_avg)

data_score = xlrd.open_workbook(dir + file_score)
course_lists = []
scores = []
idx = 0
for tab in data_score.sheets():
    course_list = []
    score_avg = score_avgs[idx]
    ncols = tab.ncols
    nrows = tab.nrows

    p1 = tab.row(0)[1:]
    p1 = [int(p.value) for p in p1]
    p2 = tab.row(1)[1:]
    p2 = [int(p.value) for p in p2]

    score = np.zeros((len(p2), nrows-2))
    for ir in range(nrows - 2):
        course_name = str(tab.row(ir + 1)[0].value)
        course_list.append(course_name)
        for ic in range(len(p2)):
            w = tab.row(ir + 2)[ic + 1].value
            if is_empty(w):
                w = 0
            else:
                vs = w.split('/')
                if len(vs) > 1:
                    if is_number(vs[0]):
                        score[ic, ir] = float(vs[0])
                    else:
                        for k in score_avg.keys():
                            if k in course_name or course_name in k:
                                score[ic, ir] = score_avg[k]
                                break
    course_lists.append(course_list)
    scores.append(score)

    idx += 1

score_avg_weighteds = []
for idx in range(len(weights)):
    weight = weights[idx]
    score = scores[idx]
    score_avg_weighted = np.sum(weight*score,1)

    weight_sum = np.sum(weight,1)
    # print('权重和', tab_names[idx])
    # for i in range(len(score_avg_weighted)):
    #     print(pmajor[i], pmin[i], weight_sum[i])

    # print('达成度平均分', tab_names[idx])
    # for i in range(len(score_avg_weighted)):
    #     print(pmajor[i], pmin[i], score_avg_weighted[i])

    score_avg_weighteds.append(score_avg_weighted)

# load indirect score
file_score_indirect_student = '毕业生间接评价.xls'
data_score_indirect_student = xlrd.open_workbook(dir + file_score_indirect_student)
sheet = data_score_indirect_student.sheets()[0]
scores_indirect_student = np.zeros((4, len(pmajor))) # 2020, 2021, 2022, 2023
for ir in range(sheet.nrows):
    for ic in range(1,sheet.ncols):
        scores_indirect_student[ir, ic-1] = sheet.row(ir)[ic].value

file_score_indirect_enterprise = '企业间接评价.xls'
data_score_indirect_enterprise = xlrd.open_workbook(dir + file_score_indirect_enterprise)
sheet = data_score_indirect_enterprise.sheets()[0]
scores_indirect_enterprise = np.zeros((1,len(pmajor)))
for ic in range(sheet.ncols):
    scores_indirect_enterprise[0, ic] = sheet.row(0)[ic].value

scores_indirect = np.zeros((4, len(pmajor)))
for i in range(4):
    scores_indirect[i] = scores_indirect_student[i] * 0.5 + scores_indirect_enterprise[0] * 0.5

# write to excel
years = ['2017', '2018', '2019', '2020']
result_output = xlwt.Workbook(encoding='utf-8')
for i, y in enumerate(years):
    result_output.add_sheet(y)
    sheet = result_output.get_sheet(i)
    sheet.write(0, 0, '毕业要求')
    sheet.write(0, 1, '指标点')
    sheet.write(0, 2, '评价方式')
    sheet.write(0, 3, '对应教学环节')
    sheet.write(0, 4, '权重')
    sheet.write(0, 5, '指标达成情况')
    sheet.write(0, 6, '评价结果')
    sheet.write(0, 7, '最终评价结果')
    sheet.write(0, 8, '毕业要求达成情况')

    course_list = course_lists[i]
    score_avg_weighted = score_avg_weighteds[i]
    score = scores[i]
    weight = weights[i]

    p1rev = pmajor[0]
    cnt = 0
    p1start = cnt
    avgs = []
    for idx, p1 in enumerate(pmajor):

        p2 = pmin[idx] # 指标 = 'p1-p2'
        istart_min = cnt
        avg = 0
        for idx_c, w in enumerate(weight[idx,:]):
            if w > 0:
                sheet.write(cnt+1, 3, course_list[idx_c])
                sheet.write(cnt+1, 4, w)
                sheet.write(cnt+1, 5, '{:.4f}'.format(score[idx, idx_c]))
                cnt += 1
                avg += w*score[idx, idx_c]

        sheet.write_merge(istart_min+1, cnt+1, 1, 1, str(p1) + '.' + str(p2))

        sheet.write_merge(istart_min+1, cnt, 2, 2, '直接评价')

        sheet.write_merge(istart_min+1, cnt, 6, 6, '{:.4f}'.format(avg))

        # 间接评价
        sheet.write(cnt+1, 2, '间接评价')
        sheet.write(cnt+1, 3, '问卷调查')
        sheet.write(cnt+1, 4, '1')

        sheet.write(cnt+1, 5, scores_indirect[3-i, idx])
        sheet.write(cnt+1, 6, scores_indirect[3-i, idx])
        # 最终评价
        sheet.write_merge(istart_min+1, cnt+1, 7, 7, '{:.4f}'.format((avg+scores_indirect[3-i, idx])/2))

        avg = (avg + scores_indirect[3-i, idx]) / 2
        avgs.append(avg)

        cnt += 1

        if p1 != p1rev:
            p1rev = p1
            sheet.write_merge(p1start + 1, istart_min, 0, 0, p1 - 1)


            aavg = 0
            for avg in avgs:
                aavg += avg
            aavg = aavg / len(avgs)

            sheet.write_merge(p1start + 1, istart_min, 8, 8, '{:.4f}'.format(aavg))


            # update p1start
            p1start = istart_min

            avgs = []
            avgs.append(avg)

        elif idx == len(pmajor)-1:
            sheet.write_merge(p1start + 1, cnt, 0, 0, p1 )

            aavg = 0
            for avg in avgs:
                aavg += avg
            aavg = aavg / len(avgs)

            sheet.write_merge(p1start + 1, cnt, 8, 8, '{:.4f}'.format(aavg))


        istart_min = cnt

result_output.save(dir + '总评价表.xls')
print('done')


