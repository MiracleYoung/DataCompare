from lib.excel import Excel
from utils import settings

def get_src_columns():
    source_column_names = []
    for path, sheetname in settings.SRC_DATA.items():
        srcpath = path
        for srcsheetname in sheetname :
            srcexcel = Excel(srcpath)
            column_names = srcexcel.get_column_names(sheetname = srcsheetname)
            src_dict_item = dict([(srcsheetname, column_names)])
            source_column_names.append(src_dict_item)


    # print(source_column_names)
    return  source_column_names

def get_tgt_columns():
    tgt_column_names = []
    for path, sheetname in settings.TGT_DATA.items():
        tgtpath = path
        for tgtsheetname in sheetname:
            tgtexcel = Excel(tgtpath)
            column_names = tgtexcel.get_column_names(sheetname=tgtsheetname)
            tgt_dict_item = dict([(tgtsheetname, column_names)])
            tgt_column_names.append(tgt_dict_item)

    # print(tgt_column_names)
    return tgt_column_names

def get_match_columns():
    match_list_all = []
    src_columns = get_src_columns()
    tgt_columns = get_tgt_columns()
    for src_item in src_columns:
        for tgt_item in tgt_columns:
            for srck,srcv in src_item.items():
                item = []
                for tgtk,tgtv in tgt_item.items():
                    if (srck == tgtk):
                        for sv in srcv:
                            for tv in tgtv:
                                if(sv ==tv):
                                    item.append(sv)
            match_list_all.append(dict([(srck, item)]))

            # print(item)
    # print(match_list_all)
    return match_list_all
    #print (match_list)

def get_del_columns():
    src_list = get_src_columns()
    mactch_list = get_match_columns()
    del_list = src_list
    del_list_all = []
    print (del_list)
    print (mactch_list)
    for matchitem in mactch_list:
        for srcitem in del_list:
            for mack, macv in matchitem.items():
                item = []
                for srck, srcv in srcitem.items():
                    if (srck == mack):
                        for mv in macv:
                            for sv in srcv:
                                if(sv == mv):
                                    srcv.remove(sv)
            del_list_all.append(dict([(srck, srcv)]))
    print(del_list_all)
    return del_list_all

def get_add_columns():
    tgt_list = get_tgt_columns()
    mactch_list = get_match_columns()
    add_list = tgt_list
    add_list_all = []
    print (add_list)
    print (mactch_list)
    for matchitem in mactch_list:
        for tgtitem in add_list:
            for mack, macv in matchitem.items():
                item = []
                for tgtk, tgtv in tgtitem.items():
                    if (tgtk == mack):
                        for mv in macv:
                            for tv in tgtv:
                                if(tv == mv):
                                    tgtv.remove(tv)
            add_list_all.append(dict([(tgtk, tgtv)]))
    print(add_list_all)
    return add_list_all






