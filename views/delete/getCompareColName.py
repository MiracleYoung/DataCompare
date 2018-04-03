from lib.excel import Excel
from utils import settings
from lib.logger import StreamFileLogger

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()

def get_match_columns(sheetname,idx=None):
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
    if idx is not None:
        for i in _matchcolumn:
            if i in idx:
                _matchcolumn.remove(i)

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

def get_compare_colNum(sheetname,idx):
    _srcpath = settings.SRC_FILE_PATH
    _tgtpath = settings.TGT_FILE_PATH
    _srcexcel = Excel(_srcpath)
    _tgtexcel = Excel(_tgtpath)
    _srccolumn = _srcexcel.get_column_names(sheetname)
    _tgtcolumn = _tgtexcel.get_column_names(sheetname)

    #getcompare column position for both sides
    _matchcolumn = []
    for _item in _srccolumn:
        if _item in _tgtcolumn and _item is not None:
            _matchcolumn.append(_item)
    for _item in _matchcolumn:
        if _item in idx:
            _matchcolumn.remove(_item)
    _sheadnum = []
    _theadnum = []
    for _maccol in _matchcolumn:
        _curshead = _srcexcel.convert_col2header(sheetname, _maccol)
        _sheadnum.append(_curshead)
    for _maccol in _matchcolumn:
        _curthead = _tgtexcel.convert_col2header(sheetname, _maccol)
        _theadnum.append(_curthead)
    _compareColZips = zip(_sheadnum, _theadnum)

    #getcompare row position for both sides

    return _compareColZips





# get_compare_colNum('CAPS Industry KPIs New',['Name'])