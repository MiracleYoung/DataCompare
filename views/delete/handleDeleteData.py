import views.delete.getMessageData as getMsgData
from utils import settings
from lib.excel import Excel
from lib.logger import StreamFileLogger


def get_diff_rowdata(sheetname):
    _srcpath = settings.SRC_FILE_PATH
    _tgtpath = settings.TGT_FILE_PATH
    srcData = getMsgData.get_data_message(_srcpath,sheetname)
    tgtData = getMsgData.get_data_message(_tgtpath,sheetname)
