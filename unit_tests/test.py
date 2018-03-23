#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 3/23/18 5:07 PM
# @Author  : Miracle Young
# @File    : test.py


from lib.excel import Excel
from utils import settings

for _path, _sheetname in settings.SRC_DATA.items():
    _excel = Excel(_path)
    print(1)

