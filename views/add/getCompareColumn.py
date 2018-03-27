from openpyxl import load_workbook
from lib.excel import Excel
from utils import settings
from lib.logger import StreamFileLogger
import pathlib



def src_column():
    for path, sheetname in settings.SRC_DATA.items():
        srcpath = path
        for srcsheetname in sheetname :
            srcexcel = Excel(srcpath)
            source_column_names = srcexcel.get_column_names(sheetname = srcsheetname)
            return source_column_names

def tgt_column():
    for path, sheetname in settings.TGT_DATA.items():
        tgtpath = path
        for tgtsheetname in sheetname :
            tgtexcel = Excel(tgtpath)
            tgt_column_names = tgtexcel.get_column_names(sheetname = tgtsheetname)
            return tgt_column_names

def compare_result():
    src_list = src_column()
    tgt_list = tgt_column()
    match_list = []
    for srcitem in src_list:
        for tgtitem in tgt_list:
            if (tgtitem == srcitem):
                match_list.append(srcitem)
    return match_list
    # print (match_list)

compare_result()




