# -*- coding: utf8 -*-

import requests
import json

import time
import xlrd
import xlwt
from xlutils.copy import copy

CaseFile = "TestCase/case.xlsx"
CurrentTime = time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))
ResultTitle = ["id", u'描述', u'请求地址', u'请求方式', u'请求参数', u'期望响应码', u'实际响应吗', u'测试结果', u'返回信息']
ResultFile = "Result_" + CurrentTime + ".xls"


# 获取用例
def get_case(test_case_file, case_no):
    testCase = xlrd.open_workbook(test_case_file)
    sheet = testCase.sheet_by_index(0)

    case_id = sheet.cell_value(case_no, 0)
    description = sheet.cell_value(case_no, 1)
    request_url = sheet.cell_value(case_no, 2)
    request_method = sheet.cell_value(case_no, 3)
    request_data = sheet.cell_value(case_no, 4)
    check_point = sheet.cell_value(case_no, 5)

    case = [case_id, description, request_url, request_method, request_data, check_point]
    return case


# 执行请求
def run_request(case_id, case_file, result_file):
    result = get_case(case_file, case_id)

    url = result[2]
    els_data = result[4]
    data_json = json.loads(els_data)

    # 根据请求方式执行请求
    if result[3] == 'POST':
        result = request_post(case_file, case_id, url, data_json)
        save_to_file(result, result_file)
    elif result[3] == 'GET':
        result = request_get(case_file, case_id, url, data_json)
        save_to_file(result, result_file)
    else:
        result.append('')
        result.append('no run')
        result.append(u'请求方式不支持')
        print(str(case_id) + '.' + str(result[3]) + ':' + '请求方式不支持')
        save_to_file(result, result_file)


# POST 请求
def request_post(case_file, case_id, url, data):
    result = get_case(case_file, case_id)

    rsp = requests.post(url=url, data=data)
    rsp_json = json.loads(rsp.text)

    code = rsp_json["resCode"]
    result.append(code)
    print(str(case_id) + "." + str(result[3]) + ':' + str(code))

    if code == result[5]:
        result.append("pass")
        result.append(rsp_json["resDesc"])
        result.append(rsp.text)
    else:
        result.append("fail")
        result.append(rsp_json["resDesc"])
        result.append(rsp.text)
    return result


# GET 请求
def request_get(case_file, case_id, url, data):
    result = get_case(case_file, case_id)

    rsp = requests.get(url=url, params=data)
    rsp_json = json.loads(rsp.text)

    code = rsp_json["resCode"]
    result.append(code)
    print(str(case_id) + "." + str(result[3]) + ':' + str(code))

    if code == result[5]:
        result.append("pass")
        result.append(rsp_json["resDesc"])
        result.append(rsp.text)
    else:
        result.append("fail")
        result.append(rsp_json["resDesc"])
        result.append(rsp.text)
    return result


# 保存结果
def save_to_file(rowlist, savefilename):
    old_excel = xlrd.open_workbook(ResultFile, formatting_info=True)
    old_sheet = old_excel.sheet_by_index(0)
    row = old_sheet.nrows

    new_excel = copy(old_excel)
    new_sheet = new_excel.get_sheet(0)

    for i in range(len(rowlist)):
        new_sheet.write(row, i, rowlist[i])

    new_excel.save(savefilename)


# 创建结果文件
def new_xls(filename, title):
    w = xlwt.Workbook()
    ws = w.add_sheet(u'Sheet1')

    for i in range(len(title)):
        ws.write(0, i, title[i])

    w.save(filename)


# 获取用例总数
def get_total_row(basic_excel):
    excel = xlrd.open_workbook(basic_excel)
    sheet = excel.sheet_by_index(0)
    return sheet.nrows


# 读取用例文件，循环执行
def run(caseFile, result, result_title):
    new_xls(result, result_title)
    for i in range(get_total_row(caseFile) - 1):
        run_request(i + 1, caseFile, result)


run(CaseFile, ResultFile, ResultTitle)
