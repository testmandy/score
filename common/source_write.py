# -*- coding: utf-8 -*-
# @Time    : 2023/4/19 10:49
# @Author  : Mandy
# -*- coding: utf-8 -*-
import xlsxwriter as xw

import conftest
from common.excel import Excel
from common.source_read import SourceData

read_filename = conftest.data_dir + '/source_data.xls'
write_filename = conftest.data_dir + '/stats_data.xls'
data = SourceData()
ex = Excel(read_filename)


def get_question_list(name):
    """根据问题名称获取问卷列表"""
    qn_list = []
    for i in range(1, ex.get_nrows()):
        questionnaire_name = data.questionnaire_name(i)
        questionnaire_id = data.questionnaire_id(i)

        if name in questionnaire_name:
            qn_list.append(questionnaire_id)
    qn_list = list(set(qn_list))
    return sorted(qn_list)


def average_score(qn_id, q_id, e_name):
    """计算单个员工平均分"""
    score_list = []
    evaluator_list = []
    if q_id == 18:
        q_name = "内容"
    else:
        q_name = "现场表现"
    for i in range(1, ex.get_nrows()):
        evaluator_name = data.evaluator_name(i)
        evaluated_name = data.evaluated_name(i)
        grade = data.grade(i)
        score = data.score(i)
        question_id = data.question_id(i)
        questionnaire_id = data.questionnaire_id(i)
        if questionnaire_id == qn_id and question_id == q_id and evaluated_name == e_name and evaluator_name != e_name and grade != 'N':
            score_list.append(score)
            evaluator_list.append(evaluator_name)
    if len(score_list) > 0:
        average = sum(score_list) / len(score_list)
        dic_row = {"qn_id": qn_id, "q_name": q_name, "name": e_name, "avg": round(float(average), 2)}
    else:
        dic_row = {"qn_id": qn_id, "q_name": q_name, "name": e_name, "avg": 0}
    return dic_row


def evaluate_count(qn_name, name_list):
    """统计评价次数（去除 N和自己）"""
    sheet_list = []
    for name in name_list:
        count = 0
        for i in range(1, ex.get_nrows()):
            evaluator_name = data.evaluator_name(i)
            evaluated_name = data.evaluated_name(i)
            questionnaire_name = data.questionnaire_name(i)
            grade = data.grade(i)
            if evaluator_name == name and evaluated_name != evaluator_name and qn_name in questionnaire_name and grade != 'N':
                count += 1
        dic_row = {"id": str(list.index(name_list, name) + 1), "name": name, "count": count / 2}
        sheet_list.append(dic_row)
        # print("%s 评价次数： %d" % (name, count / 2))
        # print("评价人： %s" % name)
        # print("评价次数： %d" % int(count / 2))
    print(sheet_list)
    return sheet_list


def xw_to_stats(data_list):  # xlsxwriter库储存数据到excel
    """将结果写入统计表"""
    workbook = xw.Workbook(write_filename)  # 创建工作簿
    for data in data_list:
        worksheet1 = workbook.add_worksheet(str(list.index(data_list, data) + 1))  # 创建子表
        worksheet1.activate()  # 激活表
        title = ['问卷id', '评价类型', '被评价人', '平均分']  # 设置表头
        worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头
        i = 2  # 从第二行开始写入数据
        for j in range(len(data)):
            insertData = [data[j]["qn_id"], data[j]["q_name"], data[j]["name"], data[j]["avg"]]
            row = 'A' + str(i)
            worksheet1.write_row(row, insertData)
            i += 1
    workbook.close()


def write_stats(name, name_list):
    """分别写入内容和临场表现"""
    sheet_list = []
    qn_list = get_question_list(name)
    for qn in qn_list:
        data_list = []
        for name in name_list:
            testData = average_score(qn, 18, e_name=name)
            data_list.append(testData)
        for name in name_list:
            testData = average_score(qn, 19, e_name=name)
            data_list.append(testData)
        print(data_list)
        sheet_list.append(data_list)
    xw_to_stats(sheet_list)
