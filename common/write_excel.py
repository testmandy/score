# -*- coding: utf-8 -*-
# @Time    : 2023/4/19 10:49
# @Author  : Mandy
# -*- coding: utf-8 -*-
import xlsxwriter as xw


def xw_toExcel(data, fileName):  # xlsxwriter库储存数据到excel
    workbook = xw.Workbook(fileName)  # 创建工作簿
    worksheet1 = workbook.add_worksheet("sheet1")  # 创建子表
    worksheet1.activate()  # 激活表
    title = ['序号', '被评价人', '总分', '平均分']  # 设置表头
    worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头
    i = 2  # 从第二行开始写入数据
    for j in range(len(data)):
        insertData = [data[j]["id"], data[j]["name"],  data[j]["sum"], data[j]["avg"]]
        row = 'A' + str(i)
        worksheet1.write_row(row, insertData)
        i += 1
    workbook.close()  # 关闭表


# # "-------------数据用例-------------"
# testData = [
#     {"id": 1, "name": "立智", "sum": 100, "avg": 100},
#     {"id": 2, "name": "维纳", "sum": 200, "avg": 100},
#     {"id": 3, "name": "如家", "sum": 300, "avg": 100},
# ]
# fileName = 'write_data.xls'
# xw_toExcel(testData, fileName)
