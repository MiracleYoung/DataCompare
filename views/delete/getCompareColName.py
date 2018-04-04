from utils import settings
from lib.logger import StreamFileLogger

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()

def get_match_columns(srcexcel,tgtexcel,sheetname,idx=None):

    _srccolumn =srcexcel.get_column_names(sheetname)
    _tgtcolumn =tgtexcel.get_column_names(sheetname)
    _matchcolumn = []
    for _item in _srccolumn:
        if _item in _tgtcolumn and _item is not None:
            _matchcolumn.append(_item)
    if idx is not None:
        for i in _matchcolumn:
            if i in idx:
                _matchcolumn.remove(i)
    _matchcolumn = filter(None, _matchcolumn)
    _sflogger.info('matched column: {}'.format(_matchcolumn))

    return _matchcolumn

def get_del_columns(srcexcel,tgtexcel,sheetname):

    _srccolumn = srcexcel.get_column_names(sheetname)
    _match_columns = get_match_columns(srcexcel,tgtexcel,sheetname)
    for _item in _match_columns:
        if _item in _srccolumn and _item is not None:
            _srccolumn.remove(_item)

    _sflogger.info('deleted column: {}'.format(_srccolumn))
    return _srccolumn

def get_add_columns(srcexcel,tgtexcel,sheetname):

    _tgtcolumn = tgtexcel.get_column_names(sheetname)
    _match_columns = get_match_columns(srcexcel,tgtexcel,sheetname)
    for _item in _match_columns:
        if _item in _tgtcolumn and _item is not None:
            _tgtcolumn.remove(_item)
    _tgtcolumn = list(filter(None,_tgtcolumn))
    _sflogger.info('added column: {}'.format(_tgtcolumn))
    return  _tgtcolumn



# get_compare_colNum('CAPS Industry KPIs New',['Name'])