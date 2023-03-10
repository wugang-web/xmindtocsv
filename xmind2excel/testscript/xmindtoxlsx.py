#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import logging
import re
import os
import sys

Base_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0,Base_DIR)
from tools.csv_xlsx import csv_xlsx_tool
from tools.xmindparserutils import get_cases
from tools.conf import *
def xmind_to_excel_file(xmind_file):
    """将xmind文件转换成csv"""
    logging.info('开始转换...', xmind_file)
    # 从xmind获取每条用例信息 函数有问题
    testcases = get_cases(xmind_file)
    # 可用于修改导出表格文件的表头
    #fileheader = ["用例标题", "前置条件", "步骤", "预期", "用例类型", "用例目录", "用例状态", "优先级","用例类型"]
    #print(fileheader)
    #fileheader = ["所属产品", "所属模块", "用例标题", "前置条件", "步骤", "预期", "优先级", "用例类型", "适用阶段"]
    tapd_testcase_rows = []
    #
    for testcase in testcases:
        row = gen_a_testcase_row(testcase)
        tapd_testcase_rows.append(row)
        #print(tapd_testcase_rows)
    csv_xlsx_tool().list_to_xlsx(tapd_testcase_rows,fileheader,xmind_file)

# 按照规则转换生成每条表格用例
def gen_a_testcase_row(testcase_dict):

    # 用例标题
    case_title = testcase_dict['title'][-1]

    # 用例状态，默认是正常
    case_status = gen_case_status(testcase_dict["labels"])
    # 优先级
    case_priority = gen_case_priority(testcase_dict['makers'])

    # "前置条件", "步骤", "预期",关联需求 ----
    case_precondition, case_step, case_expected_result, demandID,note,date_time,product,module = gen_case_step_and_expected_result( testcase_dict['note'])

    # 用例的目录
    dir = ""
    # 标注
    callout=""
    row = [case_title, case_precondition, case_step, case_expected_result, demandID, date_time,product,module,case_priority]
    return row

def gen_case_step_and_expected_result(preconditions):
    """
    根据xmind的用例类型信息解析出：前置条件|步骤|预期|用例类型，对应的内容
    :param preconditions: xmind用例类型内容
    :return:
    """
    # 利用正则表达式，获取对应内容 "用例标题", "前置条件", "步骤", "预期", "用例类型",  "适用阶段","所属产品", "所属模块"
    #print("------12-%s",preconditions)
    listStep = re.split('前置条件|步骤|预期|用例类型|适用阶段|所属产品|所属模块', preconditions)
    #"前置条件", "步骤", "预期","用例类型","所属产品", "所属模块","适用阶段"
    #print("------",listStep)

    try:
        # case_precondition=re.sub("(\d)", r"\n\1", listStep[1].strip())
        case_precondition = re.sub("\n+", "\n", listStep[1].strip())
    except:
        case_precondition = None

    try:
        case_step = re.sub("\n+", "\n", listStep[2].strip())
    except:
        case_step = None

    try:
        case_expected_result = re.sub("\n+", "\n", listStep[3].strip())
    except:
        case_expected_result = None

    try:#date_time,product,module
        demandID =int( re.sub("\n+","\n",listStep[4].strip().strip()))
        #print("-----%s" % demandID)
    except:
        demandID = None

    date_time = listStep[5].strip().strip()

    # try:
    #     date_time =int( re.sub("\n+","\n",listStep[5].strip().strip()))
    #     print("-----%s" % date_time)
    # except:
    #     date_time = None
    #
    product = listStep[6].strip().strip()
    # try:
    #     product =int( re.sub("\n+","\n",listStep[6].strip().strip()))
    # except:
    #     product = None
    #
    module = listStep[7].strip().strip()
    # try:
    #     module =int( re.sub("\n+","\n",listStep[7].strip().strip()))
    # except:
    #     module = None
    #
    try:
        note =int( re.sub("\n+","\n",listStep[0].strip().strip()))
    except:
        note = None
    #print(listStep[1].strip(),listStep[2].strip(),listStep[3].strip(),listStep[4].strip().strip(),listStep[5].strip().strip(),listStep[6].strip().strip(),listStep[7].strip().strip())
    #print(case_precondition, case_step, case_expected_result, demandID,date_time,product,module)
    return case_precondition, case_step, case_expected_result, demandID,note,product,module,date_time


def gen_case_priority(priority):
    """
    转换测试优先级
    :param priority:
    :return:
    """
    mapping = {"priority-1": '高', "priority-2": '中', "priority-3": '低'}
    if priority in mapping.keys():
        return mapping[priority]
    else:
        return 'none'


def gen_case_status(status):
    """
    转换测试用例的状态
    :param priority:
    :return:
    """
    try:
        if "待更新" in status:
            return "待更新"
        else:
            return " "
    except:
        print(status)



if __name__ == '__main__':
    xmind_file = r'D:\work\学习记录\定制化\江苏银行\测试\江苏银行.xmind'
    tapd_csv_file = xmind_to_excel_file(xmind_file)

