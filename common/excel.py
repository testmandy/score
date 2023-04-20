# -*- coding: utf-8 -*-
# @Time    : 2022/5/10 14:59
# @Author  : Mandy

import xlrd

import conftest
import xlsxwriter as xw


class Excel(object):
    def __init__(self, filename):
        # filename = conftest.data_dir + '/source_data.xls'
        self.data = xlrd.open_workbook(filename)  # 文件名以及路径，如果路径或者文件名有中文给前面加一个 r
        self.table = self.data.sheets()[0]  # 通过索引顺序获取

    def get_nrows(self):
        nrows = self.table.nrows
        # print('当前表共有 %d 行' % nrows)
        return nrows

    def get_ncols(self):
        ncols = self.table.ncols
        print(ncols)
        return ncols

    def get_row(self, rowx):
        row_value = self.table.row_values(rowx)
        print(row_value)

    def get_col(self, colx):
        col_value = self.table.col_values(colx)
        print(col_value)

    def get_cell(self, rowx, colx):
        cell_value = self.table.cell_value(rowx, colx)
        return cell_value

    def get_rows_value(self):
        for i in range(self.get_nrows()):
            for j in range(self.get_ncols()):
                self.get_cell(i, j)

    def xw_toExcel(self, data_list, write_filename):  # xlsxwriter库储存数据到excel
        workbook = xw.Workbook(write_filename)  # 创建工作簿
        for data in data_list:
            worksheet1 = workbook.add_worksheet(str(list.index(data_list, data)+1))  # 创建子表
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

