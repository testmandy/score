# -*- coding: utf-8 -*-
# @Time    : 2022/5/10 16:05
# @Author  : Mandy

from common.read_excel import Excel
from common.score_data import Data

name_list = ["cherry", "Gary", "陈培挺", "崔莹新", "傅萌", "金鼎强", "黎焕", "李晨阳", "刘道熠", "毛伟伟",
             "沈韦婷", "沈振鹏", "谭敦钊", "万露", "汪建明", "王晓丹", "吴定康", "吴章华", "熊艺", "余帆"]

data = Data()
ex = Excel()


def average_score(qn_id, q_id, e_name):
    score_list = []
    evaluator_list = []
    if q_id == 18:
        q_name = "内容"
    else:
        q_name = "现场表现"
    for i in range(1, ex.get_nrows()):
        evaluator_name = data.evaluator_name(i)
        evaluated_name = data.evaluated_name(i)
        score = data.score(i)
        question_id = data.question_id(i)
        questionnaire_id = data.questionnaire_id(i)
        if questionnaire_id == qn_id and question_id == q_id and evaluated_name == e_name and evaluator_name != e_name:
            score_list.append(score)
            evaluator_list.append(evaluator_name)
    if len(score_list) > 0:
        average = sum(score_list) / len(score_list)
        dic_row = {"qn_id": qn_id, "q_name": q_name, "name": e_name, "avg": round(float(average), 2)}
    else:
        dic_row = {"qn_id": qn_id, "q_name": q_name, "name": e_name, "avg": 0}
        # print("问卷 %s, 问题 %d, 被评价人：%s" % (qn_id, q_id, e_name))
        # print("分数列表： %s" % list_score)
        # print("评价人列表： %s" % list_evaluator)
        # print("总分： %s" % sum(list_score))
        # print("平均分： %s" % round(float(average), 2))
    return dic_row


def evaluate_count():
    for name in name_list:
        count = 0
        for i in range(1, ex.get_nrows()):
            evaluator_name = data.evaluator_name(i)
            evaluated_name = data.evaluated_name(i)
            if evaluator_name == name and evaluated_name != evaluator_name:
                count += 1
        # print("%s 评价次数： %d" % (name, count/2))
        # print(name)
        print(int(count/2))


def write(qn_list):
    sheet_list = []
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

    ex.xw_toExcel(sheet_list)


if __name__ == '__main__':
    ex = Excel()
    my_qn_list = [29, 32, 34, 35, 36]
    write(my_qn_list)
    evaluate_count()
