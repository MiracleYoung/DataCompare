from lib.logger import StreamFileLogger
import openpyxl
from openpyxl.styles import colors,PatternFill
from openpyxl import Workbook
from lib.excel import Excel
from utils import settings

_sflogger = StreamFileLogger(settings.LOG_FILE, __file__).get_logger()

def read_src_excel(sheetname):
    excel = Excel(settings.SRC_FILE_PATH)
    srccolnames = excel.get_column_names(sheetname)
    src_colnames = []
    for src_col_item in srccolnames:
        if src_col_item is not None:
            src_colnames.append(src_col_item)
    return src_colnames

def read_tgt_excel(sheetname):
    excel = Excel(settings.TGT_FILE_PATH)
    tgtcolnames = excel.get_column_names(sheetname)
    tgt_colnames = []
    for tgt_col_item in tgtcolnames:
        if tgt_col_item is not None:
            tgt_colnames.append(tgt_col_item)
    return tgt_colnames

def get_same_columns(sheetname):
    src_items = read_src_excel(sheetname)
    tgt_items = read_tgt_excel(sheetname)
    same_items = []

    for item in tgt_items:
        if item in src_items:
            same_items.append(item.upper())
    return same_items
