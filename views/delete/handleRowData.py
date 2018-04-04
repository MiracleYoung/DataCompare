import views.delete.getMessageData as getMsgData
import views.delete.getCompareColName as getColumn
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
    print(_numlist)
    return _numlist

def get_matchIdx_rowdNum(srcexcel,tgtexcel,sheetname,idx=None):
    if idx is None:
        _srcData = getMsgData.get_srcdata_message(srcexcel, tgtexcel, sheetname)
        _tgtData = getMsgData.get_tgtdata_message(srcexcel, tgtexcel, sheetname)
    else:
        _srcData = getMsgData.get_srcdata_message(srcexcel, tgtexcel, sheetname, idx)
        _tgtData = getMsgData.get_tgtdata_message(srcexcel, tgtexcel, sheetname, idx)

    #store match rowNum in both file
    _numlist = []
    for _i in range(0,len(_srcData)-1):
        for _j in range(0,len(_tgtData)-1):
            if _srcData[_i] == _tgtData[_j]:
                list1 = (str(_i+2)+','+str(_j+2)).split(',')
                _numlist.append(list1)
                break

    print(_numlist)
    return _numlist


def setBgColorRow(srcexcel,tgtexcel,sheetname):
    _getrowsNum = get_diff_rowdNum(srcexcel,tgtexcel,sheetname)
    _wb = tgtexcel.get_wb()
    _ws = tgtexcel.get_sheet(sheetname)
    for curitem in _ws.iter_rows():

        if curitem[0].row in _getrowsNum:
            for cell in curitem:
                cell.fill = PatternFill(fgColor = 'FF0000', fill_type = 'solid')
    _wb.save(settings.END_FILE_PATH)

def setBgColorRowIdx(srcexcel,tgtexcel,sheetname,idx):
    _getrowsNum = get_matchIdx_rowdNum(srcexcel,tgtexcel,sheetname,idx)
    _wb = tgtexcel.get_wb()
    _ws = tgtexcel.get_sheet(sheetname)
    # only flag target file
    _srcws = srcexcel.get_sheet(sheetname)
    _getZips= getMsgData.get_compare_colNum(srcexcel,tgtexcel,sheetname,idx)
    print('loop start')
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

        if curitem[0].row  in _getdiffrowsNum:
            for cell in curitem:
                cell.fill = PatternFill(fgColor = 'FF0000', fill_type = 'solid')

    #set add columns color
    _addColumn = getColumn.get_add_columns(srcexcel,tgtexcel,sheetname)
    if _addColumn is not None:
        #convert add column name into excel head(A B C D AA...)
        for i in range(0, len(_addColumn)):
            _addColumn[i]  = tgtexcel.convert_col2header(sheetname, _addColumn[i])
        print(_addColumn)

        for _row in _ws.iter_rows():
            for _cellitem in _row:
                if _cellitem.column in _addColumn:
                    _cellitem.fill = PatternFill(fgColor='FF0000', fill_type='solid')

    _wb.save(settings.END_FILE_PATH)


def test():
    _srcpath = settings.SRC_FILE_PATH
    _tgtpath = settings.TGT_FILE_PATH
    _srcexcel = Excel(_srcpath)
    _tgtexcel = Excel(_tgtpath)
    # getMsgData.get_srcdata_message(_srcexcel,_tgtexcel,'CAPS Industry KPIs New','PRIMARY CONTACT_EMAIL')
    setBgColorRowIdx(_srcexcel,_tgtexcel,'CAPS Industry KPIs New','PRIMARY CONTACT_EMAIL')

test()
