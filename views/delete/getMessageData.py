
import views.delete.getCompareColName as compareData
from utils import settings
from lib.excel import Excel
from lib.logger import StreamFileLogger

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()
def get_data_message(path,sheetname):

    _excel = Excel(path)
    _mactch_column_name = compareData.get_match_columns(sheetname)
    if _mactch_column_name :
        _headername = ''
        for columnname in _mactch_column_name:
            _curcells = _excel.convert_col2header(sheetname,columnname)
            _headername = _headername + _curcells
    else:
        ''
    _cursheet = _excel.get_sheet(sheetname)
    print(_cursheet.max_row)
    _alldata = []
    for _row in range(2, _cursheet.max_row):
        _rowdata = []
        for _column in _headername:
            _cellname = "{}{}".format(_column, _row)
            _cellvalue = _cursheet[_cellname].value
            _rowdata.append(_cellvalue)
        _alldata.append(_rowdata)
    _list_dict_output =[]
    for _item in _alldata:
        _getcount = 0
        for _count in _alldata:
            if _item == _count:
                _getcount = _getcount+1
        _dictitem ={'rowdata':_item,'count':_getcount}
        _list_dict_output.append(_dictitem)
    _sflogger.debug('matchedData: {}'.format(_list_dict_output))