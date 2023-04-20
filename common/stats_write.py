# -*- coding: utf-8 -*-
# @Time    : 2023/4/19 10:49
# @Author  : Mandy
# -*- coding: utf-8 -*-
import xlsxwriter as xw

import conftest
from common.excel import Excel
from common.source_write import evaluate_count
from common.stats_read import StatsData


read_filename = conftest.data_dir + '/stats_data.xls'
write_filename = conftest.data_dir + '/results_data.xls'
data = StatsData()
ex = Excel(read_filename)


def xw_to_results(score_data, evaluate_data):  # xlsxwriter库储存数据到excel
    workbook = xw.Workbook(write_filename)  # 创建工作簿
    worksheet1 = workbook.add_worksheet("环评平均分")  # 创建子表
    worksheet1.activate()  # 激活表
    title = ['id', '问卷id', '姓名', '平均分']  # 设置表头
    worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头
    i = 2  # 从第二行开始写入数据
    for j in range(len(score_data)):
        insertData = [score_data[j]["id"], score_data[j]["qn_id"], score_data[j]["name"], score_data[j]["avg"]]
        row = 'A' + str(i)
        worksheet1.write_row(row, insertData)
        i += 1

    worksheet2 = workbook.add_worksheet("评价次数")  # 创建子表
    worksheet2.activate()  # 激活表
    title = ['id', '姓名', '评价次数']  # 设置表头
    worksheet2.write_row('A1', title)  # 从A1单元格开始写入表头
    i = 2  # 从第二行开始写入数据
    for j in range(len(evaluate_data)):
        insertData = [evaluate_data[j]["id"], evaluate_data[j]["name"], evaluate_data[j]["count"]]
        row = 'A' + str(i)
        worksheet2.write_row(row, insertData)
        i += 1

    workbook.close()


def avg_score(name_list):
    score_list = []
    qn_id = data.qustion_id(1)
    for name in name_list:
        sum_score = 0
        for i in range(1, ex.get_nrows()):
            qn_id = data.qustion_id(i)
            evaluated_name = data.evaluated_name(i)
            score = str(data.avg_score(i))
            if name == evaluated_name:
                sum_score = float(score) + sum_score
        avg = sum_score / 2
        dic_row = {"id": str(list.index(name_list, name) + 1), "qn_id": qn_id, "name": name, "avg": avg}
        score_list.append(dic_row)
    return score_list


def write_results(qn_name, name_list):
    score_list = avg_score(name_list)
    evaluate_list = evaluate_count(qn_name, name_list)
    xw_to_results(score_list, evaluate_list)
