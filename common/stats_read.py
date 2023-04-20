# -*- coding: utf-8 -*-
# @Time    : 2022/5/10 15:49
# @Author  : Mandy
import conftest
from common.excel import Excel


class StatsData(object):
    def __init__(self):
        filename = conftest.data_dir + '/stats_data.xls'
        self.read = Excel(filename)

    def qustion_id(self, rowx):
        data = self.read.get_cell(rowx, 0)
        return data

    def evaluate_type(self, rowx):
        data = self.read.get_cell(rowx, 1)
        return data

    def evaluated_name(self, rowx):
        data = self.read.get_cell(rowx, 2)
        return data

    def avg_score(self, rowx):
        data = self.read.get_cell(rowx, 3)
        return data




