# -*- coding: utf8 -*-

import requests
import json

url = 'http://test.anxinyisheng.com/home/question/detailQuestion'
params = {'questionId': '59c333037f1008310d8b4567'}
header = {
    "Wxid": "o79aixECshqXft8Cck5fMC7LdYZs",
    "Channel": "wx_anxinjiankang",
    "User-Agent": "micromessenger",
    "Auth": "57320a1b2a8e8cdfa121b417bd0ef547"}

rsp = requests.post(url, params=params, headers=header)
print(rsp)
