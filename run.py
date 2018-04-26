# -*- coding: utf8 -*-

import requests
import json


import time
import xlrd
import xlwt
from xlutils.copy import copy

CaseFile = "TestCase/case.xlsx"
ResultTitleRow = \
    ["id", "Description", "Request_URL", "Method", "Request_Data", "ExpectedRspCode", "ActualRspCode", "TestResult",
     "RspMeessage"]
CurrentTime = time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))
ResultFile = "Result_" + CurrentTime + ".xls"


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


def run_request(caseId, case, result):
    case = get_case(case, caseId)

    url = case[2]
    els_data = case[4]
    data_json = json.loads(els_data)

    header = {"Wxid": "o79aixECshqXft8Cck5fMC7LdYZs",
              "Channel": "wx_anxinjiankang",
              "User-Agent": "micromessenger"}

    rsp = requests.post(url=url, data=data_json, headers=header)
    rsp_json = json.loads(rsp.text)
    code = rsp_json["code"]
    case.append(code)

    if code == case[5]:
        case.append("pass")
        case.append("null")
    else:
        case.append("fail")
        case.append(rsp_json["msg"])

    save_to_file(case, result)


def save_to_file(rowlist, savefilename):
    old_excel = xlrd.open_workbook(ResultFile, formatting_info=True)
    old_sheet = old_excel.sheet_by_index(0)
    row = old_sheet.nrows
    print(row)

    new_excel = copy(old_excel)
    new_sheet = new_excel.get_sheet(0)

    for i in range(len(rowlist)):
        new_sheet.write(row, i, rowlist[i])

    new_excel.save(savefilename)


def new_xls(filename, title):
    w = xlwt.Workbook()
    ws = w.add_sheet(u'Sheet1')

    for i in range(len(title)):
        ws.write(0, i, title[i])

    w.save(filename)


def get_total_row(basic_excel):
    excel = xlrd.open_workbook(basic_excel)
    sheet = excel.sheet_by_index(0)
    return sheet.nrows


def run(case, result, result_title):
    new_xls(result, result_title)
    for i in range(get_total_row(case) - 1):
        run_request(i + 1, case, result)


# run_request(1, CaseFile, ResultFile)
# run_request(2, CaseFile, ResultFile)
run(CaseFile, ResultFile, ResultTitleRow)
