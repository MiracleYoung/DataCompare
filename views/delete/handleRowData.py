import views.delete.getMessageData as getMsgData
from utils import settings
from lib.excel import Excel
from lib.logger import StreamFileLogger
from openpyxl.styles import PatternFill

def get_diff_rowdNum(srcexcel,tgtexcel,sheetname,idx=None):
    if idx is None:
        _srcData = getMsgData.get_srcdata_message(srcexcel,tgtexcel,sheetname)
        _tgtData = getMsgData.get_tgtdata_message(srcexcel,tgtexcel,sheetname)
    else:
        _srcData = getMsgData.get_srcdata_message(srcexcel,tgtexcel,sheetname,idx)
        _tgtData = getMsgData.get_tgtdata_message(srcexcel,tgtexcel,sheetname,idx)
    _numlist = []
    lineNum = 2
    for _tgtitem in _tgtData:
        if (_tgtitem not in _srcData):
            _numlist.append(lineNum)
        lineNum += 1
    return _numlist

def get_matchIdx_rowdNum(srcexcel,tgtexcel,sheetname,idx=None):
    if idx is None:
        _srcData = getMsgData.get_srcdata_message(srcexcel, tgtexcel, sheetname)
        _tgtData = getMsgData.get_tgtdata_message(srcexcel, tgtexcel, sheetname)
    else:
        _srcData = getMsgData.get_srcdata_message(srcexcel, tgtexcel, sheetname, idx)
        _tgtData = getMsgData.get_tgtdata_message(srcexcel, tgtexcel, sheetname, idx)

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


def setBgColor(srcexcel,tgtexcel,sheetname):
    _getrowsNum = get_diff_rowdNum(srcexcel,tgtexcel,sheetname)
    _wb = tgtexcel.get_wb()
    _ws = tgtexcel.get_sheet(sheetname)
    for curitem in _ws.iter_rows():

        if curitem[0].row in _getrowsNum:
            for cell in curitem:
                cell.fill = PatternFill(fgColor = 'FF0000', fill_type = 'solid')
    _wb.save(settings.END_FILE_PATH)

def setBgColorIdx(srcexcel,tgtexcel,sheetname,idx):
    _getrowsNum = get_matchIdx_rowdNum(srcexcel,tgtexcel,sheetname,idx)
    _wb = tgtexcel.get_wb()
    _ws = tgtexcel.get_sheet(sheetname)
    # only flag target file
    _srcws = srcexcel.get_sheet(sheetname)
    _getZips= getMsgData.get_compare_colNum(srcexcel,tgtexcel,sheetname,idx)
    #set color in same index but different cell value
    for _row in _getrowsNum:
        for _zip in _getZips:
            #getbothCellsName _zip[0] is srcrownum,_zip[1] is tgrrownum.row is same
            _srccellname = "{}{}".format(_zip[0], _row[0])
            _tgtcellname = "{}{}".format(_zip[1], _row[1])
            _srclvalue = _srcws[_srccellname].value
            _tgtlvalue = _ws[_tgtcellname].value
            _srclvalue = str(_srclvalue).strip().upper()
            _tgtlvalue = str(_tgtlvalue).strip().upper()
            if(_srclvalue !=_tgtlvalue):
                _ws[_tgtcellname].fill = PatternFill(fgColor = 'FF0000', fill_type = 'solid')
    # set different index ,highlight all cell color
    _getdiffrowsNum = get_diff_rowdNum(srcexcel,tgtexcel,sheetname,idx)
    for curitem in _ws.iter_rows():

        if curitem[0].row in _getdiffrowsNum:
            for cell in curitem:
                cell.fill = PatternFill(fgColor = 'FF0000', fill_type = 'solid')

    _wb.save(settings.END_FILE_PATH)


def test():
    _srcpath = settings.SRC_FILE_PATH
    _tgtpath = settings.TGT_FILE_PATH
    _srcexcel = Excel(_srcpath)
    _tgtexcel = Excel(_tgtpath)
    setBgColorIdx(_srcexcel,_tgtexcel,'CAPS Industry KPIs New','PRIMARY CONTACT_EMAIL')


test()
