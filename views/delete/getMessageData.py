import views.delete.getCompareColName as compareData
from utils import settings
from lib.excel import Excel
from lib.logger import StreamFileLogger

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()
def get_data_message(path,sheetname):

    _excel = Excel(path)
    _mactch_column_name = compareData.get_match_columns(sheetname)
    if _mactch_column_name :
        _headername = []
        for columnname in _mactch_column_name:
            _curcells = _excel.convert_col2header(sheetname,columnname)
            _headername.append(_curcells)

    else:
        ''
    _cursheet = _excel.get_sheet(sheetname)

    _alldata = []

    for _row in range(2, _cursheet.max_row+1):
        _rowdata = []
        for _column in _headername:
            _cellname = "{}{}".format(_column, _row)
            #get current cell line number and line column

            _cellvalue = _cursheet[_cellname].value

            _rowdata.append(str(_cellvalue).upper())
            #upper all values

        _alldata.append(_rowdata)

    for _item in _alldata:
        _getcount = _alldata.count(_item)
        _item.append(_getcount)

    return _alldata

# get_data_message(path = 'C:/Users/jliu409/DataCompare/src_data/CAPS.xlsx',sheetname = 'CAPS Industry KPIs New')