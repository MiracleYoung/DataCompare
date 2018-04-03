#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 3/23/18 5:07 PM
# @Author  : Miracle Young
# @File    : test.py


if __name__ == '__main__':
    import argparse

    _desc = '''
        The sheet name should like "Sheet1,Sheet2,Sheet3", use double quote include sheetnames
        balabala
    '''
    parser = argparse.ArgumentParser(description=_desc)
    parser.add_argument("--foo-qwe", type=str)
    parser.add_argument("-v", "--verbosity", dest='verbosity', nargs='?')
    args = parser.parse_args()
    print(args)
    # print(args.qwe)
    print(args.verbosity)
