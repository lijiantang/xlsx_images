#! /usr/bin/env python
# -*- coding: utf-8 -*-
# vim:fenc=utf-8

"""
下载xlsx图片
"""
from openpyxl.reader.excel import load_workbook
import re
import wget
#取第一张表
wb = load_workbook('数据源.xlsx')
sheetnames = wb.get_sheet_names()
ws = wb.get_sheet_by_name(sheetnames[0])

print ("Title:",ws.title)
print ("Row:", ws.max_row)
print ("Cols:",ws.max_column)

#建立字典存储表格内容
data_dic = {} 

for rx in range(2, ws.max_row+1):
   # temp_list = []
    pid = rx
    w1 = ws.cell(row=rx, column=1).value
    w2 = ws.cell(row=rx, column=2).value
    #temp_list = [w1] 
    data_dic[w1] = w2


for keys,values in data_dic.items():
   # print (keys,values)
    a = str(values)
    if a == 'None':
        continue
    strinfo = re.compile('60x60')
    b = strinfo.sub('600x600',a)
    #print (b)
    c_out = 'picture/' + str(keys) + '.jpg'
    wget.download(b, c_out)

