#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 3/14/18 3:48 PM
# @Author  : Miracle Young
# @File    : excel.py

from openpyxl import load_workbook
from functools import wraps

from lib.logger import StreamFileLogger
from lib.handles import DBHandle
from utils import settings

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()


class Excel:
    def __init__(self, path):
        self._wb = load_workbook(filename=path, data_only=True)
        _sflogger.info('Open excel: {}'.format(path))

    def get_wb(self):
        return self._wb

    def get_sheet(self, sheetname):
        return self._wb[sheetname]

    def get_all_sheetnames(self):
        return self._wb.get_sheet_names()

    def get_max_row(self, sheetname):
        return self.get_sheet(sheetname).max_row

    def get_max_column(self, sheetname):
        return self.get_sheet(sheetname).max_column

    def get_dimensions(self, sheetname):
        return self.get_sheet(sheetname).dimensions.split(':')

    def get_column_names(self, sheetname, start=None, end=None, mapping=None):
        _start, _end = start, end if start and end else self.get_dimensions(sheetname)
        _sheet = self.get_sheet(sheetname)
        if mapping:
            _columns = ['[{}]'.format(mapping.get(_column.value)) for _column in _sheet[_start:_end][0]]
        else:
            _columns = ['[{}]'.format(_column.value) for _column in _sheet[_start:_end][0]]
        return _columns

    def excel2db(self, conn, schema, table, sheetname, start=None, end=None, mapping=None, size=100, **kwargs):
        '''
        load data from excel to db
        :param start, end: excel start pos, end pos
        :example: start='A1', end='Q5'
        '''
        _start, _end = start, end if start and end else self.get_dimensions(sheetname)
        _custom_columns = kwargs.get('custom_columns', None)
        _custom_values = kwargs.get('_custom_values', None)
        try:
            _sheet = self.get_sheet(sheetname)
            _sflogger.debug('Load <{}>[{}][{}] starting...'.format(sheetname, _start, _end))
            _columns = self.get_column_names(sheetname, _start, _end, mapping)
            _sflogger.debug('columns: {}'.format(_columns))
            _table_raw = _sheet[_start: _end][1:]
            _sflogger.debug('Load [{}].[{}] start...'.format(schema, table))
            DBHandle.bulk_insert(conn, schema, table, _columns, _table_raw, _custom_columns, _custom_values, size)
        except Exception as e:
            _sflogger.error('Execute failed.', exc_info=True)

    def read_excel_by_pos(self, sheetname, start=None, end=None, mapping=None):
        try:
            _start, _end = start, end if start and end else self.get_dimensions(sheetname)
            _sheet = self.get_sheet(sheetname)
            _sflogger.debug('Read {}{}{}: starting...'.format(sheetname, _start, _end))
            _columns = self.get_column_names(sheetname, _start, _end, mapping)
            _sflogger.debug('columns: {}'.format(_columns))
            _table_raw = _sheet[_start: _end][1:]
            _ret = []
            for _i, _row in enumerate(_table_raw):
                _value = tuple([_v.value for _v in _row])
                _ret.append(_value)
                if _i % 100 == 0 and _i != 0:
                    _sflogger.info('row {}: {}'.format(_i, tuple(_value)))
            _sflogger.debug('Read complete. Total counts: {}'.format(_i))
            return _ret
        except Exception as e:
            _sflogger.error('Execute failed.', exc_info=True)
            return []
