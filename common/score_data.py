# -*- coding: utf-8 -*-
# @Time    : 2022/5/10 15:49
# @Author  : Mandy
from common.read_excel import Excel


class Data(object):
    def __init__(self):
        self.read = Excel()

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

    def grade(self, rowx):
        data = self.read.get_cell(rowx, 4)
        return data

    def score(self, rowx):
        data = self.read.get_cell(rowx, 5)
        return int(data)

    def evaluated_name(self, rowx):
        data = self.read.get_cell(rowx, 6)
        return data

    def questionnaire_id(self, rowx):
        data = self.read.get_cell(rowx, 7)
        return int(data)

    def questionnaire_content(self, rowx):
        data = self.read.get_cell(rowx, 8)
        return data

    def question_id(self, rowx):
        data = self.read.get_cell(rowx, 9)
        return int(data)

    def question_content(self, rowx):
        data = self.read.get_cell(rowx, 10)
        return data


# if __name__ == '__main__':
#     data = Data()
#     print(data.evaluator_name(1))
#     print(data.score(1))
#     print(data.evaluated_name(1))



