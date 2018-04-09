#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 3/20/18 3:21 PM
# @Author  : Miracle Young
# @File    : handles.py

from lib.logger import StreamFileLogger
from etc import settings

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()


class DBHandle:
    def __init__(self):
        pass

    @staticmethod
    def insertsql_fmt(schema, tgt_table, columns, values):
        _insert_sql = "insert into [{}].[{}] ({}) values {}"
        _sql = _insert_sql.format(schema, tgt_table, ','.join(columns), tuple(values))
        _sql = _sql.replace('values (', 'values ')[:-1]
        _sql = _sql.replace('None', 'NULL')
        if len(columns) == 1:
            _sql = _sql.replace(',)', ')')
        if len(values) == 1:
            _sql = _sql.rstrip(',')
        _sflogger.debug('Insert sql: {}'.format(_sql))
        return _sql

    @staticmethod
    def bulk_insert(conn, schema, table, columns, raw_data, custom_columns=None, custom_values=None, size=100):
        _count = len(raw_data)
        _values = []
        if custom_columns and isinstance(custom_columns, (list, tuple)):
            columns = columns + custom_columns
        for _i, _row in enumerate(raw_data):
            if custom_values and isinstance(custom_values, (list, tuple)):
                _value = tuple([str(_v.value) for _v in _row] + custom_values[_i])
            elif hasattr(_row[0], 'value'):
                _value = tuple([str(_v.value) for _v in _row])
            else:
                _value = _row
            _values.append(_value)
            # each size rows commit once.
            # execute rest data, insufficient to 100 rows.
            if (_i % size == 0 and _i != 0) or (_i + 1 == _count):
                _sql = DBHandle.insertsql_fmt(schema, table, columns, _values)
                _sflogger.debug('Row {}: {}'.format(_i, tuple(_value)))
                conn.execute(_sql)
                _sflogger.debug('Load {}-{} complete.'.format(_i - size, _i))
                _values.clear()
            if _i + 1 == _count:
                _sflogger.debug('Last row {}: {}'.format(_i, tuple(_value)))
        _sflogger.debug('Load {} complete. Total counts: {}'.format(table, _count))


