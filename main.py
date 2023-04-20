# -*- coding: utf-8 -*-
# @Time    : 2023/4/19 11:07
# @Author  : Mandy
from common.excel import Excel

if __name__ == '__main__':
    ex = Excel()
    testData = [
        {"id": 1, "name": "立智", "sum": 100, "avg": 100},
        {"id": 2, "name": "维纳", "sum": 200, "avg": 100},
        {"id": 3, "name": "如家", "sum": 300, "avg": 100},
    ]
    ex.xw_toExcel(testData)