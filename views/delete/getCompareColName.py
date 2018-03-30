from lib.excel import Excel
from utils import settings
from lib.logger import StreamFileLogger

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()
def get_match_columns(srcpath,tgtpath,sheetname):
    _srcexcel = Excel(srcpath)
    _tgtexcel = Excel(tgtpath)
    _srccolumn =_srcexcel.get_column_names(sheetname)
    _tgtcolumn =_tgtexcel.get_column_names(sheetname)
    _matchcolumn = []
    for _item in _srccolumn:
        if _item in _tgtcolumn:
            _matchcolumn.append(_item)

    _sflogger.debug('matched column: {}'.format(_matchcolumn))
    return _matchcolumn

def get_del_columns(path,sheetname):

    _srcexcel = Excel(path)
    _srccolumn = _srcexcel.get_column_names(sheetname)
    _match_columns = get_match_columns(sheetname)
    for _item in _match_columns:
        if _item in _srccolumn:
            _srccolumn.remove(_item)
    _sflogger.debug('deleted column: {}'.format(_srccolumn))
    return _srccolumn

def get_add_columns(path,sheetname):
    _tgtexcel = Excel(path)
    _tgtcolumn = _tgtexcel.get_column_names(sheetname)
    _match_columns = get_match_columns(sheetname)
    for _item in _match_columns:
        if _item in _tgtcolumn:
            _tgtcolumn.remove(_item)
    _sflogger.debug('added column: {}'.format(_tgtcolumn))
    return  _tgtcolumn





