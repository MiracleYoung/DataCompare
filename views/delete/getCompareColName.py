from lib.excel import Excel
from utils import settings
from lib.logger import StreamFileLogger

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()

def get_match_columns(sheetname):
    _srcpath = settings.SRC_FILE_PATH
    _tgtpath = settings.TGT_FILE_PATH
    _srcexcel = Excel(_srcpath)
    _tgtexcel = Excel(_tgtpath)
    _srccolumn =_srcexcel.get_column_names(sheetname)
    _tgtcolumn =_tgtexcel.get_column_names(sheetname)
    _matchcolumn = []
    for _item in _srccolumn:
        if _item in _tgtcolumn and _item is not None:
            _matchcolumn.append(_item)

    _sflogger.info('matched column: {}'.format(_matchcolumn))
    return _matchcolumn

def get_del_columns(sheetname):
    _path = settings.SRC_FILE_PATH
    _srcexcel = Excel(_path)
    _srccolumn = _srcexcel.get_column_names(sheetname)
    _match_columns = get_match_columns(sheetname)
    for _item in _match_columns:
        if _item in _srccolumn and _item is not None:
            _srccolumn.remove(_item)
    _sflogger.info('deleted column: {}'.format(_srccolumn))
    return _srccolumn

def get_add_columns(sheetname):
    _path = settings.TGT_FILE_PATH
    _tgtexcel = Excel(_path)
    _tgtcolumn = _tgtexcel.get_column_names(sheetname)
    _match_columns = get_match_columns(sheetname)
    for _item in _match_columns:
        if _item in _tgtcolumn and _item is not None:
            _tgtcolumn.remove(_item)
    _sflogger.info('added column: {}'.format(_tgtcolumn))
    return  _tgtcolumn


