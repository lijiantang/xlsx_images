#! /usr/bin/env python
# -*- coding: utf-8 -*-
# vim:fenc=utf-8

"""
下载xlsx图片
"""
from openpyxl.reader.excel import load_workbook
import re
import wget
import sys
import datetime
import os
import time


#取第一张表
table = sys.argv[1]
wb = load_workbook('%s.xlsx' % table)
sheetnames = wb.get_sheet_names()
ws = wb.get_sheet_by_name(sheetnames[0])

print ("表名:",ws.title)
print ("总行数:", ws.max_row)
print ("总列数:",ws.max_column)

####创建文件夹
today = datetime.datetime.now()
todayStr = today.strftime("%Y%m%d%H%M")
if not os.path.exists(todayStr):
    os.mkdir(todayStr)

for rx in range(2, ws.max_row+1):
   # temp_list = []
    keys = ws.cell(row=rx, column=1).value
    values = ws.cell(row=rx, column=2).value
    url = str(values)
    if url == 'None':
        continue
    strinfo = re.compile('60x60')
    newurl = strinfo.sub('600x600',url)
    jpg_name = todayStr + '/' + str(keys) + '.jpg'
    wget.download(newurl, jpg_name)
    time.sleep (1)
