# -*- coding: utf-8 -*-
# @Time    : 2022/5/10 15:49
# @Author  : Mandy
import conftest
from common.excel import Excel


class SourceData(object):
    def __init__(self):
        filename = conftest.data_dir + '/source_data.xls'
        self.read = Excel(filename)

    def submit_batch(self, rowx):
        data = self.read.get_cell(rowx, 0)
        return data

    def submit_time(self, rowx):
        data = self.read.get_cell(rowx, 1)
        return data

    def evaluator_id(self, rowx):
        data = self.read.get_cell(rowx, 2)
        return int(data)

    def evaluator_name(self, rowx):
        data = self.read.get_cell(rowx, 3)
        return data

    def evaluated_id(self, rowx):
        data = self.read.get_cell(rowx, 4)
        return data

    def evaluated_name(self, rowx):
        data = self.read.get_cell(rowx, 5)
        return data

    def questionnaire_id(self, rowx):
        data = self.read.get_cell(rowx, 6)
        return int(data)

    def questionnaire_name(self, rowx):
        data = self.read.get_cell(rowx, 7)
        return data

    def template_name(self, rowx):
        data = self.read.get_cell(rowx, 8)
        return int(data)

    def question_id(self, rowx):
        data = self.read.get_cell(rowx, 9)
        return int(data)

    def question_content(self, rowx):
        data = self.read.get_cell(rowx, 10)
        return data

    def grade(self, rowx):
        data = self.read.get_cell(rowx, 11)
        return data

    def score(self, rowx):
        grade = self.grade(rowx)
        if grade == 'S':
            score = 5
        elif grade == 'A':
            score = 4
        elif grade == 'B':
            score = 3
        elif grade == 'C':
            score = 2
        elif grade == 'D':
            score = 1
        else:
            score = 0
        return int(score)

# if __name__ == '__main__':
#     data = Data()
#     print(data.evaluator_name(1))
#     print(data.score(1))
#     print(data.evaluated_name(1))



