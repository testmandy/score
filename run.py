# -*- coding: utf-8 -*-
# @Time    : 2023/4/18 16:05
# @Author  : Mandy
import time

from common.source_write import write_stats
from common.stats_write import write_results

name_list = ["cherry", "Gary", "陈培挺", "崔莹新", "傅萌", "金鼎强", "黎焕", "李晨阳", "刘道熠", "柳阳", "马相连", "蒙仕彬",
             "毛伟伟", "沈韦婷", "沈振鹏", "谭敦钊", "万露", "汪建明", "王晓丹", "吴定康", "吴章华", "熊艺", "余帆"]


if __name__ == '__main__':
    question_name = "午休环评"
    write_stats(question_name, name_list)
    time.sleep(10)
    write_results(question_name, name_list)

