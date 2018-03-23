#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 3/14/18 9:33 AM
# @Author  : Miracle Young
# @File    : settings.py

import logging, pathlib

SETTING_FILE = pathlib.Path(__file__)
PROJECT_BASE_DIR = SETTING_FILE.parent.parent
CONFIG_DIR = PROJECT_BASE_DIR / 'etc'
LOG_CONFIG_FILE = CONFIG_DIR / 'logger.conf'
LOG_FILE = PROJECT_BASE_DIR / 'logs' / 'caps.log'
STREAM_LOG_LEVEL = logging.INFO
FILE_LOG_LEVEL = logging.DEBUG

EXCEL_PATH = PROJECT_BASE_DIR / 'src_data'

SRC_DATA = {
    (EXCEL_PATH / 'CAPS.xlsx').as_posix(): 'CAPS Industry KPIs New',
}
