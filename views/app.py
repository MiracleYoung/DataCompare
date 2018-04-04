#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 4/2/18 3:37 PM
# @Author  : Miracle Young
# @File    : app.py


import argparse
from utils import settings, config
from views.delete.handleRowData import setBgColorRowIdx

_desc = '''
        The sheet name should like "Sheet1,Sheet2,Sheet3", use double quote include sheetnames
        balabala
    '''
_parser = argparse.ArgumentParser(description=_desc)
_parser.add_argument('--src', type=str, help='Source file')
_parser.add_argument('--src-sheets', type=str, help='Source file sheets')
_parser.add_argument('--tgt', type=str, help='Target file')
_parser.add_argument('--tgt-sheets', type=str, help='Target file sheets')

# _args = _parser.parse_args(
#     ['--src', 'abc.xlsx', '--src-sheets', 'CAPS,APQC,3rdparty', '--tgt', 'def.xlsx', '--tgt-sheets', 'qwe,asd,zxc'])
_args = _parser.parse_args()
print(_args)

def main():
    setBgColorRowIdx(_args.src, _args.src_sheets, _args.tgt, _args.tgt_sheets)


if __name__ == '__main__':

    _config = config.Config()
    setattr(_config, 'SRC_DATA', {_args.src: _args.src_sheets})
    setattr(_config, 'TGT_DATA', {_args.tgt: _args.tgt_sheets})

    main()

    # print(config.Config)
