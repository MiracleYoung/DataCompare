
import views.delete.getCompareColumns as compareData
from utils import settings



def get_source_message():
    for path, sheetname in settings.SRC_DATA.items():
        for srcsheetname in sheetname :
            compareData.get_match_columns(srcsheetname)
            compareData.get_del_columns(srcsheetname)
            compareData.get_add_columns(srcsheetname)


