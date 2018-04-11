from etc import settings
from lib.logger import StreamFileLogger
from lib.excel import Excel
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

    # _sflogger.info('matched column: {}'.format(_matchcolumn))

    return _matchcolumn

# def get_del_columns(srcexcel,tgtexcel,sheetname):
#
#     _srccolumn = srcexcel.get_column_names(sheetname)
#     _match_columns = get_match_columns(srcexcel,tgtexcel,sheetname)
#     for _item in _match_columns:
#         if _item in _srccolumn and _item is not None:
#             _srccolumn.remove(_item)
#     print(_srccolumn)
#     _sflogger.info('deleted column: {}'.format(_srccolumn))
#     return _srccolumn

def get_add_columns(srcexcel,tgtexcel,sheetname):

    _tgtcolumn = tgtexcel.get_column_names(sheetname)
    _match_columns = get_match_columns(srcexcel,tgtexcel,sheetname)
    for _item in _match_columns:
        if _item in _tgtcolumn and _item is not None:
            _tgtcolumn.remove(_item)
    _tgtcolumn = list(filter(None,_tgtcolumn))
    # _sflogger.info('added column: {}'.format(_tgtcolumn))
    return  _tgtcolumn

def conver_header(sheet, column_name):
    for _row in sheet.rows:
        for _j, _cell in enumerate(_row):
            if _cell.value.upper() == column_name.upper():
                return _cell.column
        break
    return ''

# get_compare_colNum('CAPS Industry KPIs New',['Name'])

# def test():
#     _srcpath = settings.SRC_FILE_PATH
#     _tgtpath = settings.TGT_FILE_PATH
#     _srcexcel = Excel(_srcpath)
#     _tgtexcel = Excel(_tgtpath)
    # getMsgData.get_srcdata_message(_srcexcel,_tgtexcel,'CAPS Industry KPIs New','PRIMARY CONTACT_EMAIL')
    # get_del_columns(_srcexcel,_tgtexcel,'CAPS Industry KPIs New')
