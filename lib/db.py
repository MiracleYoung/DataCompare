#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 3/9/18 11:50 AM
# @Author  : Miracle Young
# @File    : db.py

import queue

from sqlalchemy import create_engine

from lib.logger import StreamFileLogger
from etc import settings

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()


class MSSQL:
    def __init__(self, dbinfo, pool_size=5, max_overflow=0):
        self._q = queue.Queue(pool_size)
        self._dbinfo = dbinfo[settings.DB_ENV]
        self._schema = self._dbinfo['SCHEMA']
        self._connstr = 'mssql+pymssql://{}:{}@{}/{}'.format(self._dbinfo['USERNAME'], self._dbinfo['PASSWORD'],
                                                             self._dbinfo['HOSTNAME'], self._dbinfo['DB'])
        self._engine = create_engine(self._connstr, pool_size=pool_size, max_overflow=max_overflow)
        self._conns = [self._q.put(self._engine.connect()) for _ in range(pool_size)]

    def get_engine(self):
        _engine_url = self._engine.url
        _sflogger.info('Get engine {}:{}/{}'.format(_engine_url.host, _engine_url.port, _engine_url.username))
        return self._engine

    def get_schema(self):
        return self._schema

    def get_conn(self):
        if not self._q.empty():
            _conn = self._q.get()
            _sflogger.info('Connect DB success.')
            return _conn

    def close_conn(self):
        if not self._q.full():
            self._q.put(self._engine.connect())
