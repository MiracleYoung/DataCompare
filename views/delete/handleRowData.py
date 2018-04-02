import views.delete.getMessageData as getMsgData
from utils import settings
from lib.excel import Excel
from lib.logger import StreamFileLogger
from openpyxl.styles import PatternFill

def get_diff_rowdata(sheetname):
    _srcpath = settings.SRC_FILE_PATH
    _tgtpath = settings.TGT_FILE_PATH
    _srcData = getMsgData.get_data_message(_srcpath,sheetname)
    _tgtData = getMsgData.get_data_message(_tgtpath,sheetname)
    _numlist = []
    lineNum = 1
    for _tgtitem in _tgtData:
        if (_tgtitem not in _srcData):
            _numlist.append(lineNum)
        lineNum += 1
    return _numlist


def setBgColor(sheetname):
    _getSetList = get_diff_rowdata(sheetname)
    _filepath = settings.END_FILE_PATH
    _excel = Excel(_filepath)
    _wb = _excel.get_wb()
    _ws = _excel.get_sheet(sheetname)
    for curitem in _ws.iter_rows():

        if curitem[0].row in _getSetList:
            for cell in curitem:
                cell.fill = PatternFill(fgColor = 'FF0000', fill_type = 'solid')
    _wb.save(settings.END_FILE_PATH)

# setBgColor('CAPS Industry KPIs New')
