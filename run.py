# -*- coding: utf8 -*-

import requests
import json

import sys
import xlrd
import xlwt
from xlutils.copy import copy


CaseFile = "TestCase/case.xlsx"
ResultFile = "Result.xls"


def get_case(test_case_file, case_no):
    testCase = xlrd.open_workbook(test_case_file)
    sheet = testCase.sheet_by_index(0)

    case_id = int(sheet.cell_value(case_no, 0))
    description = sheet.cell_value(case_no, 1)

    request_url = sheet.cell_value(case_no, 2)
    request_method = sheet.cell_value(case_no, 3)
    request_data = sheet.cell_value(case_no, 4)

    check_point = sheet.cell_value(case_no, 5)

    case = [case_id, description, request_url, request_method, request_data, check_point]
    return case


def get_sum_case(test_case_file):
    testCase = xlrd.open_workbook(test_case_file)
    sheet = testCase.sheet_by_index(0)

    return sheet.nrows


def run_request(caseId):
    case = get_case(CaseFile, caseId)
    print(type(case))

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
    else:
        case.append("fail")

    case.append(rsp_json)
    print(case)

    save_to_file(case, ResultFile)


def save_to_file(rowlist, savefilename):
    old_excel = xlrd.open_workbook(ResultFile, formatting_info=True)
    old_sheet = old_excel.sheet_by_index(0)
    row = old_sheet.nrows

    new_excel = copy(old_excel)
    new_sheet = new_excel.get_sheet(0)

    for i in range(len(rowlist)):
        new_sheet.write(row+1, i, rowlist[i])

    new_excel.save(savefilename)


def new_xls(filename, title):
    w = xlwt.Workbook()
    ws = w.add_sheet(u'Sheet1')

    for i in range(len(title)):
        ws.write(0, i, title[i])

    w.save(filename)


result_title = ["id", "Description", "Request_URL", "Method", "Request_Data", "Check_Point", "ActualResult",
                    "TestResult", "Response"]
new_xls(ResultFile, result_title)

run_request(1)

