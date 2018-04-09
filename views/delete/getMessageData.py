import views.delete.getCompareColName as compareData
from utils import settings
from lib.logger import StreamFileLogger
from lib.excel import Excel
_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()
def get_srcdata_message(srcexcel,tgtexcel,sheetname,idx=None):

    #initial index into list
    _headername = []

    #store message data
    _alldata = []
    if idx is None:
        _mactch_column_name = compareData.get_match_columns(srcexcel,tgtexcel,sheetname)
        for columnname in _mactch_column_name:
            _curcells = srcexcel.convert_col2header(sheetname,columnname)
            _headername.append(_curcells)
    else:
        _indexCols = idx.split(',')
        for columnname in _indexCols:
            _curcells = srcexcel.convert_col2header(sheetname, columnname)
            _headername.append(_curcells)

    _cursheet = srcexcel.get_sheet(sheetname)

    #get start row number
    _startNumber = 0
    for _row in _cursheet.iter_rows():
        for _cell in _row:
            if(_cell.value is not None):
                _startNumber = _cell.row
                break
        if(_startNumber != 0):
            break
    # real data start from header number + 1
    _startNumber = _startNumber + 1

    for _row in range(_startNumber, _cursheet.max_row+1):
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

    return _alldata,_startNumber,10

def get_tgtdata_message(srcexcel,tgtexcel,sheetname,idx=None):

    #initial index into list
    _headername = []

    #store message data
    _alldata = []
    if idx is None:
        _mactch_column_name = compareData.get_match_columns(srcexcel,tgtexcel,sheetname)
        for columnname in _mactch_column_name:
            _curcells = tgtexcel.convert_col2header(sheetname,columnname)
            _headername.append(_curcells)
    else:
        _indexCols = idx.split(',')

        for columnname in _indexCols:
            _curcells = tgtexcel.convert_col2header(sheetname, columnname)
            _headername.append(_curcells)

    _cursheet = tgtexcel.get_sheet(sheetname)

    # get start row number
    _startNumber = 0
    for _row in _cursheet.iter_rows():
        for _cell in _row:
            if (_cell.value is not None):
                _startNumber = _cell.row
                break
        if (_startNumber != 0):
            break
     # real data start from header number + 1
    _startNumber = _startNumber + 1

    for _row in range(_startNumber, _cursheet.max_row):
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
    return _alldata,_startNumber



def get_compare_colNum(srcexcel,tgtexcel,sheetname,idx):
    _indexCols = idx.split(',')
    _srccolumn = srcexcel.get_column_names(sheetname)
    _srcsheet  = srcexcel.get_sheet(sheetname)
    _tgtcolumn = tgtexcel.get_column_names(sheetname)
    _tgtsheet  = tgtexcel.get_sheet(sheetname)
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
    _sflogger.info('get_column_names start1:')
    for _maccol in _matchcolumn:
        _cursrcshead = compareData.conver_header(_srcsheet, _maccol)
        _sheadnum.append(_cursrcshead)
    _sflogger.info('get_column_names start2:')
    for _maccol in _matchcolumn:
        _curtgtshead = compareData.conver_header(_tgtsheet, _maccol)
        _theadnum.append(_curtgtshead)
    _sflogger.info('get_column_names end:')
    _compareCols = list(zip(_sheadnum, _theadnum))


    #getcompare row position for both sides
    return _compareCols




# get_srcdata_message('CAPS Industry KPIs New','Name')

# def test():
#     _srcpath = settings.SRC_FILE_PATH
#     _tgtpath = settings.TGT_FILE_PATH
#     _srcexcel = Excel(_srcpath)
#     _tgtexcel = Excel(_tgtpath)
#     get_srcdata_message(_srcexcel,_tgtexcel,'CAPS Industry KPIs New','PRIMARY CONTACT_EMAIL')
#
# test()

