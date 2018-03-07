# -*- coding: utf8 -*-

import requests
import json

import sys
import xlrd
import xlwt


CaseFile = "TestCase/case.xlsx"


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


def run_request():
    url = get_case(CaseFile, 1)[2]
    params = get_case(CaseFile, 1)[4]
    header = {"Wxid": "o79aixECshqXft8Cck5fMC7LdYZs",
              "Channel": "wx_anxinjiankang",
              "User-Agent": "micromessenger"}

    print(url, params)

    rsp = requests.get(url=url, params=params, headers=header)
    rsp_json = json.loads(rsp.text)

    print(rsp_json)


run_request()