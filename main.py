# !/usr/bin/env python
# -*- coding: utf-8 -*-

__author__ = "calmXia"

import sys

from excel.excel import Excel


# C:\Users\calm.xia\Desktop\UNISOCLogParse\UNISOCLogDatabase_V1.1.xlsx
def main():
    LogDataBase=sys.argv[1]
    tables = Excel(LogDataBase)
    print(tables.getdata_by_module(module="Sensor"))
    #print(tables.getdata_by_col_index(4, end_rowx = 4))

if __name__ == "__main__":
    main()
