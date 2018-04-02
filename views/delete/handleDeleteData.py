import views.delete.getMessageData as getMsgData
from utils import settings
from lib.excel import Excel
from lib.logger import StreamFileLogger


def get_diff_rowdata(sheetname):
    _srcpath = settings.SRC_FILE_PATH
    _tgtpath = settings.TGT_FILE_PATH
    _srcData = getMsgData.get_data_message(_srcpath,sheetname)
    _tgtData = getMsgData.get_data_message(_tgtpath,sheetname)
    _numlist = []
    lineNum = 2
    for _tgtitem in _srcData:
        if (_tgtitem not in _tgtData):
            _numlist.append(lineNum)
        lineNum += 1
    print(_numlist)


get_diff_rowdata('CAPS Industry KPIs New')
