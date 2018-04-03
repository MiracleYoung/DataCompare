import views.delete.getMessageData as getMsgData
from utils import settings
from lib.excel import Excel
from lib.logger import StreamFileLogger
from openpyxl.styles import PatternFill

def get_diff_rowdNum(sheetname,idx=None):
    _srcpath = settings.SRC_FILE_PATH
    _tgtpath = settings.TGT_FILE_PATH
    if idx is None:
        _srcData = getMsgData.get_data_message(_srcpath,sheetname)
        _tgtData = getMsgData.get_data_message(_tgtpath,sheetname)
    else:
        _srcData = getMsgData.get_data_message(_srcpath, sheetname,idx)
        _tgtData = getMsgData.get_data_message(_tgtpath, sheetname,idx)
    _numlist = []
    lineNum = 2
    for _tgtitem in _tgtData:
        if (_tgtitem not in _srcData):
            _numlist.append(lineNum)
        lineNum += 1
    return _numlist

def get_matchIdx_rowdNum(sheetname,idx=None):
    _srcpath = settings.SRC_FILE_PATH
    _tgtpath = settings.TGT_FILE_PATH
    if idx is None:
        _srcData = getMsgData.get_data_message(_srcpath,sheetname)
        _tgtData = getMsgData.get_data_message(_tgtpath,sheetname)
    else:
        _srcData = getMsgData.get_data_message(_srcpath, sheetname,idx)
        _tgtData = getMsgData.get_data_message(_tgtpath, sheetname,idx)

    #store match rowNum in target file
    tgt_numlist = []
    _tgtlineNum = 2
    for _tgtitem in _tgtData:
        if (_tgtitem  in _srcData):
            tgt_numlist.append(_tgtlineNum)
        _tgtlineNum += 1
    # store match rowNum in src file
    _src_numlist = []
    _srclineNum = 2
    for _srcitem in _srcData:
        if (_srcitem in _tgtData):
            _src_numlist.append(_srclineNum)
        _srclineNum += 1
    _compareRowList = list(zip(_src_numlist,tgt_numlist))
    print(_compareRowList)
    return _compareRowList


def setBgColor(sheetname):
    _getrowsNum = get_diff_rowdNum(sheetname)
    _filepath = settings.TGT_FILE_PATH
    _excel = Excel(_filepath)
    _wb = _excel.get_wb()
    _ws = _excel.get_sheet(sheetname)
    for curitem in _ws.iter_rows():

        if curitem[0].row in _getrowsNum:
            for cell in curitem:
                cell.fill = PatternFill(fgColor = 'FF0000', fill_type = 'solid')
    _wb.save(settings.END_FILE_PATH)

def setBgColorIdx(sheetname,idx):
    _getrowsNum = get_matchIdx_rowdNum(sheetname,idx)
    _filepath = settings.TGT_FILE_PATH
    _comparepath = settings.SRC_FILE_PATH
    _excel = Excel(_filepath)
    _srcexcel = Excel(_comparepath)

    _wb = _excel.get_wb()
    _ws = _excel.get_sheet(sheetname)
    # only flag target file

    _srcws = _srcexcel.get_sheet(sheetname)
    _getZips= getMsgData.get_compare_colNum(sheetname,idx)
    #set color in same index but different cell value
    for _row in _getrowsNum:
        for _zip in _getZips:
            #getbothCellsName _zip[0] is srcrownum,_zip[1] is tgrrownum.row is same
            _srccellname = "{}{}".format(_zip[0], _row[0])
            _tgtcellname = "{}{}".format(_zip[1], _row[1])
            _srclvalue = _srcws[_srccellname].value
            _tgtlvalue = _ws[_tgtcellname].value
            _srclvalue = str(_srclvalue).upper()
            _tgtlvalue = str(_tgtlvalue).upper()
            if(_srclvalue !=_tgtlvalue):
                _ws[_tgtcellname].fill = PatternFill(fgColor = 'FF0000', fill_type = 'solid')
    # set different index ,highlight all cell color
    _getdiffrowsNum = get_diff_rowdNum(sheetname,idx)
    for curitem in _ws.iter_rows():

        if curitem[0].row in _getdiffrowsNum:
            for cell in curitem:
                cell.fill = PatternFill(fgColor = 'FF0000', fill_type = 'solid')

    _wb.save(settings.END_FILE_PATH)










setBgColorIdx('CAPS Industry KPIs New','PRIMARY CONTACT_EMAIL')
