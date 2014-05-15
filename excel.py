# -*- coding: utf-8 -*-
import xlrd

data_all = {}
data = ()

rb = xlrd.open_workbook('/home/rusto/public/excel/wr.xls',formatting_info=True)
sheet = rb.sheet_by_index(0)
print 'cols -', sheet.ncols, ' rows -', sheet.nrows

printrow = -1

for rownum in range(sheet.nrows):
    for colnum in range(sheet.ncols):
        if sheet.cell_value(rownum, colnum) == u'Наименование точки поставки': printrow = rownum
        if (rownum == printrow) or (rownum == printrow + 1) or (rownum == printrow + 2):
            if sheet.cell_value(rownum, colnum) <> '':
                print colnum, sheet.cell_value(rownum, colnum)
        else: printrow = -1

colnum = 14
for rownum in range(sheet.nrows):
    if sheet.cell_value(rownum, colnum) <> '':
        print rownum, sheet.cell_value(rownum, colnum)
