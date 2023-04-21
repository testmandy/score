# -*- coding: utf-8 -*-
# @Time    : 2022/5/10 14:59
# @Author  : Mandy

import xlrd

import conftest
import xlsxwriter as xw


class Excel(object):
    def __init__(self, fileName, index=0):
        self.data = xlrd.open_workbook(fileName)  # 文件名以及路径，如果路径或者文件名有中文给前面加一个 r
        self.index = index

    def get_nsheets(self):
        nsheets = self.data.nsheets
        return nsheets

    def get_sheet(self):
        sheet = self.data.sheets()[self.index]  # 通过索引顺序获取
        return sheet

    def get_nrows(self):
        nrows = self.get_sheet().nrows
        # print('当前表共有 %d 行' % nrows)
        return nrows

    def get_ncols(self):
        ncols = self.get_sheet().ncols
        print(ncols)
        return ncols

    def get_row(self, rowx):
        row_value = self.get_sheet().row_values(rowx)
        print(row_value)

    def get_col(self, colx):
        col_value = self.get_sheet().col_values(colx)
        print(col_value)

    def get_cell(self, rowx, colx):
        cell_value = self.get_sheet().cell_value(rowx, colx)
        return cell_value

    def get_rows_value(self, index):
        for i in range(self.get_nrows(index)):
            for j in range(self.get_ncols()):
                self.get_cell(i, j, index)

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


if __name__ == '__main__':
    filename = conftest.data_dir + '/stats_data.xls'
    ex = Excel(filename, index=0)
    tables = ex.get_nsheets()
    print(tables)