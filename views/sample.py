#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 3/23/18 4:32 PM
# @Author  : Miracle Young
# @File    : app.py

from lib.logger import StreamFileLogger
from lib.excel import Excel
from utils import settings

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()

_sflogger.info(1)
_sflogger.debug(2)

try:
    a
except Exception as e:
    _sflogger.error('Failed', exc_info=True)



excel_name = settings.SRC_DATA['a.xlsx']


excel = Excel(excel_name)

excel.get_columns('CAPS Industry KPIs New', 'A1', 'F28')
# get all sheetname
excel.read_excel_by_pos('CAPS Industry KPIs New', 'A1', 'F28')

#



