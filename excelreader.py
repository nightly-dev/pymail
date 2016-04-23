#!/usr/bin/python
# -*- coding: utf-8 -*-
import xlrd
import os


def get_xls_name():
    files = os.listdir(os.getcwd())
    xls_pool = [f for f in files if f.endswith('xls') or f.endswith('xlsx')]
    if not len(xls_pool):
        print u'未找到相关的excel文件'
        return None
    if len(xls_pool) > 1:
        print u'发现多个excel文件，请保留一个.如当前excel正在打开，请关闭再试'
        return None
    return xls_pool[0]

def read_xls():
    fn = get_xls_name()
    if not fn:
        print u'出错啦，请检查'
        return None, None, None, None
    try:
        data = xlrd.open_workbook('%s' % fn)
        table = data.sheet_by_index(0)
        nrows = table.nrows
        ncols = table.ncols
        return nrows, ncols, table.row_values, fn
    except:
        print u'excel文件读取失败'
    return None, None, None, None

def check_xls(offset):
    ret = True
    fn = get_xls_name()
    if not fn:
        print u'出错啦，请检查'
        return False
    try:
        data = xlrd.open_workbook('%s' % fn)
        table = data.sheet_by_index(0)
        nrows = table.nrows
        ncols = table.ncols
        title_count = len(table.row_values(offset))
        for r in range(nrows):
            row_value = table.row_values(r)
            for c in range(len(row_value)):
                if table.cell(r, c).ctype == 1 and table.cell(r, c).value.isspace():
                    print u'第%d行第%d列为空格，请删除后重新运行本程序!' % (r+1, c+1)
                    ret = False
            if len(row_value) > title_count:
                print u'第%d行过长，请修正!' % (r+1)
                ret = False
            if len(row_value) < title_count:
                print u'第%d行过短，请修正!' % (r+1)
                ret = False
    except:
        print u'excel文件读取失败'
        ret = False
    return ret

