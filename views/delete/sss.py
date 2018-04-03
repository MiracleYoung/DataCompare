import views.delete.getCompareColName as compareData
from utils import settings
from lib.excel import Excel
from lib.logger import StreamFileLogger

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()
def get_data_message(path,sheetname,idx=None):

    #initial index into list
    _headername = []

    #store message data
    _alldata = []
    _excel = Excel(path)
    if idx is None:
        _mactch_column_name = compareData.get_match_columns(sheetname)
        for columnname in _mactch_column_name:
            _curcells = _excel.convert_col2header(sheetname,columnname)
            _headername.append(_curcells)
    else:
        _indexCols = idx.split(',')
        _mactch_column_name = compareData.get_match_columns(sheetname, _indexCols)
        for columnname in _indexCols:
            _curcells = _excel.convert_col2header(sheetname, columnname)
            _headername.append(_curcells)

    _cursheet = _excel.get_sheet(sheetname)

    for _row in range(2, _cursheet.max_row+1):
        _rowdata = []
        for _column in _headername:
            _cellname = "{}{}".format(_column, _row)
            #get current cell line number and line column

            _cellvalue = _cursheet[_cellname].value
            _cellvalue = str(_cellvalue).upper()
            _rowdata.append(_cellvalue.strip())
            #upper all values

        _alldata.append(_rowdata)
    for _item in _alldata[::-1]:
        if _item in _alldata:
            _getcount = _alldata.count(_item)
            _item.append(_getcount)
    return _alldata


def get_compare_colNum(sheetname,idx):
    _indexCols = idx.split(',')
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
        if _item in _indexCols:
            _matchcolumn.remove(_item)
    _sheadnum = []
    _theadnum = []
    for _maccol in _matchcolumn:
        _curshead = _srcexcel.convert_col2header(sheetname, _maccol)
        _sheadnum.append(_curshead)
    for _maccol in _matchcolumn:
        _curthead = _tgtexcel.convert_col2header(sheetname, _maccol)
        _theadnum.append(_curthead)
    _compareCols = list(zip(_sheadnum, _theadnum))
    #getcompare row position for both sides
    return _compareCols