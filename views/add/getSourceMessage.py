
from lib.excel import Excel
from utils import settings
import pathlib
import views.add.getCompareColumn as Getcompare


# pathlidef getsrcMeg():
for path, sheetname in settings.SRC_DATA.items():
    srcpath = path
    for srcsheetname in sheetname :
        excel = Excel(srcpath)
        dim_list = excel.get_dimensions(srcsheetname)
        strStart = dim_list[0]
        endStart = dim_list[1]

        source_data = excel.read_excel_by_pos(sheetname = srcsheetname)
        source_column_names = excel.get_column_names(sheetname = srcsheetname)

        new_dict = {}
        list_source = []
        for singledata in source_data:
            new_dict = dict([(source_column_names[i],singledata[i] ) for i in range(len(source_column_names))])
            list_source.append(new_dict)

        for myiter in  list_source:
            print(myiter)

Getcompare.compare_result()


