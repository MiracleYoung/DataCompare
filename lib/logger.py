#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 3/13/18 9:29 AM
# @Author  : Miracle Young
# @File    : logger.py

import logging.config
from functools import wraps

from etc.settings import STREAM_LOG_LEVEL, FILE_LOG_LEVEL


class Logger:
    def __init__(self, logger, level=logging.DEBUG):
        self._fmt_debug = '%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s'
        self._fmt_info = '%(asctime)s %(filename)-12s: %(message)s'
        self._fmt_dt = '%Y-%m-%d %H:%M:%S'
        self._logger = logging.getLogger(logger)
        self._logger.setLevel(level)

    def get_logger(self):
        return self._logger


class StreamLogger(Logger):
    def __init__(self, logger):
        super(StreamLogger, self).__init__(logger)
        self._handler = logging.StreamHandler()
        self._handler.setLevel(STREAM_LOG_LEVEL)
        self._fmt = logging.Formatter(self._fmt_info, self._fmt_dt)
        self._handler.setFormatter(self._fmt)
        self._logger.addHandler(self._handler)


class FileLogger(Logger):
    def __init__(self, path, logger):
        super(FileLogger, self).__init__(logger)
        self._handler = logging.FileHandler(path)
        self._handler.setLevel(FILE_LOG_LEVEL)
        self._fmt = logging.Formatter(self._fmt_debug, self._fmt_dt)
        self._handler.setFormatter(self._fmt)
        self._logger.addHandler(self._handler)


class StreamFileLogger(Logger):
    def __init__(self, path, logger):
        super(StreamFileLogger, self).__init__(logger)
        self._fhandler = logging.FileHandler(path)
        self._fhandler.setLevel(FILE_LOG_LEVEL)
        self._fmt = logging.Formatter(self._fmt_debug, self._fmt_dt)
        self._fhandler.setFormatter(self._fmt)
        self._logger.addHandler(self._fhandler)

        self._shandler = logging.StreamHandler()
        self._shandler.setLevel(STREAM_LOG_LEVEL)
        self._fmt = logging.Formatter(self._fmt_info, self._fmt_dt)
        self._shandler.setFormatter(self._fmt)
        self._logger.addHandler(self._shandler)


def dec_step_log(step, table, logger):
    def _wrap(fn):
        @wraps(fn)
        def __wrap(*args, **kwargs):
            logger.info('Step {}: Loading {} table...'.format(step, table))
            ret = fn(*args, **kwargs)
            logger.info('Step {} completed.'.format(step))
            return ret
        return __wrap
    return _wrap
