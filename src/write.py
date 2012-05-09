#!/usr/bin/env python
#-*- coding:utf-8 -*-
"""
@version: d$
@author: U{sam han<mailto:samrain.han@gmail.com>}
@license:
@contact:
@see:
"""
from xlwt import *
import xlwt
import sqlite3
import xlrd

book = xlwt.Workbook()

sheet1 = book.add_sheet("overview")

"""
    复制一页
"""
xlsfile = xlrd.open_workbook('【TG-IT(1203-2)】任务执行跟踪表(岗位职能).xls')
worksheet = xlsfile.sheet_by_name('overview')
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = -1
while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    curr_cell = -1
    while curr_cell < num_cells:
        curr_cell += 1
        sheet1.write(curr_row, curr_cell,worksheet.cell_value(curr_row, curr_cell))

"""
    打开一个数据库
"""
conn = sqlite3.connect('taskdb.sqlite')
cur = conn.cursor()

ldata=[]

cur.execute('select taskgroup as "任务组",count(1) as "任务总数",sum(ET)/480 as "累计预计工时",sum(onetime+twotime+threetime+fourtime+fivetime+sixtime+seventime)/480 as "累计实际工时" from task group by taskgroup')
rows = cur.fetchall()
r = 0
for row in rows:
    ldata.append(row[0])
    sheet1.write(r,0,row[0])
    sheet1.write(r,1,row[1])
    sheet1.write(r,7,row[2])
    sheet1.write(r,8,row[3])
    r += 1
rrr = r


cur.execute('select taskgroup,count(1) from task where flag4finish = 1 group by taskgroup')
rows = cur.fetchall()
r = 0
for row in rows:
    while r < rrr:
        if ldata[r]==row[0]:
            sheet1.write(r,2,row[1])
        r += 1

cur.execute('select taskgroup,count(1) from task where result = "合格" group by taskgroup')
rows = cur.fetchall()
r = 0
for row in rows:
    while r < rrr:
        if ldata[r]==row[0]:
            sheet1.write(r,3,row[1])
        r += 1

cur.execute('select taskgroup,count(1) from task where flag4delay = 1 group by taskgroup')
rows = cur.fetchall()
r = 0
for row in rows:
    while r < rrr:
        if ldata[r] == row[0]:
            sheet1.write(r,4,row[1])
        r += 1

cur.execute('select taskgroup,count(1) from task where flag4trac = 1 group by taskgroup')
rows = cur.fetchall()
r = 0
for row in rows:
    while r < rrr:
        if ldata[r]==row[0]:
            sheet1.write(r,5,row[1])
        r += 1

cur.execute('select taskgroup,count(1) from task where flag4check = 1 group by taskgroup')
rows = cur.fetchall()
r = 0
for row in rows:
    while r < rrr:
        if ldata[r]==row[0]:
            sheet1.write(r,6,row[1])
        r += 1

cur.close()
conn.close()
book.save("writeTest.xls")
