#!/usr/bin/env python
#-*- coding:utf-8 -
"""
@version:  0.1
@author: U{sam han<mailto:samrain.han@gmail.com>}
@license:
@contact:
@see:
"""
import xlrd
import sqlite3
import time
import datetime
from datetime import date
from datetime import timedelta
import os
from xlwt import *
import xlwt
"""
    打开一个数据库
"""
conn = sqlite3.connect('taskdb.sqlite')
#conn = sqlite3.connect(":memory:")'数据库放在内存中
cur = conn.cursor()

"""
    打开一个目录下所有xls文件
"""
dir_name = "/home/rain/下载/trac"
startdate = datetime.date(2012,5,14)
file_list = [f_name for f_name in os.listdir(dir_name) if f_name.endswith('xls')]
for f_in_name in file_list:
    print f_in_name
    xlsfile = xlrd.open_workbook(os.path.join(dir_name,f_in_name))
    """
        根据sheet位置，打开XLS文件中某个Sheet
    """
    worksheet = xlsfile.sheet_by_index(0)
    """
        取得每个Cell的值
    """
    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1
    """
        从第3行开始取数据
    """
    curr_row = 2
    """
        循环取数据
    """
    taskinfo = []
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        curr_cell = -1
        rowinfo = []
        flag4finish = 0#默认没有完成
        flag4delay = 1#默认没有逾期
        flag4trac = 1#默认尚未开展跟踪
        flag4check = 1#默认没有检查
        flag4qualty = 1 #默认质量分评定正确
        while curr_cell < num_cells:
            curr_cell += 1
            cell_type = worksheet.cell_type(curr_row, curr_cell)
            cell_value = worksheet.cell_value(curr_row, curr_cell)
#            if cell_type == 3:
#                tuplex4xls = list(xlrd.xldate_as_tuple(cell_value,0))
#                tuplex4xls.append(23)
#                tuplex4xls.append(59)
#                tuplex4xls.append(59)
#                cell_value = time.strftime('%Y-%m-%d %H:%M:%S',tuplex4xls)
#                print cell_value
            if curr_cell in (10,12,14,16,18,20,22) and cell_value == 100:#表示完成
                flag4finish = 1
                if rowinfo[2] < date.isoformat(startdate+timedelta(curr_cell-11)):#逾期
#                    print curr_cell,rowinfo[1].encode('utf-8'),rowinfo[2],date.isoformat(startdate+timedelta(curr_cell-11))
                    flag4delay = 0
                if rowinfo[5] == '':#没有记录位置
                    flag4trac = 0
                elif rowinfo[7] == '' and rowinfo[5] not in('N/A','NA'): #如果有质量记录时没有打质量分
                    flag4qualty = 0
                if rowinfo[6] == '': #没有打效果分
                    flag4check = 0
            rowinfo.append(cell_value)
        rowinfo.append(flag4delay)
        rowinfo.append(flag4trac)
        rowinfo.append(flag4check)
        rowinfo.append(flag4qualty)
        rowinfo.append(flag4finish)
        cur.execute('insert into checktrac values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',rowinfo)

book2 = xlwt.Workbook()
sheet2 = book2.add_sheet("checktrac")
cur.execute('select plan,count(1),sum(ET)/60,sum(firsttime+secondetime+thirdtime+fourthtime+fifthtime+sixthtime+seventhtime)/60 from checktrac group by plan')
r = 1
sheet2.write(0,0,u'计划名称')
sheet2.write(0,1,u'任务数')
sheet2.write(0,2,u'小计：预计工时')
sheet2.write(0,3,u'小计：实际工时')
for row in cur.fetchall():
    sheet2.write(r,0,row[0])
    sheet2.write(r,1,row[1])
    sheet2.write(r,2,row[2])
    sheet2.write(r,3,row[3])
    r += 1

cur.execute('select plan,count(1) from checktrac where flag4finish = 1 group by plan')
r += 1
sheet2.write(r,0,u'计划名称')
sheet2.write(r,1,u'完成任务数')
r += 1
for row in cur.fetchall():
    sheet2.write(r,0,row[0])
    sheet2.write(r,1,row[1])
    r += 1

cur.execute('select plan,count(1) from checktrac where flag4delay = 0 group by plan')
r += 1
sheet2.write(r,0,u'计划名称')
sheet2.write(r,1,u'逾期任务数')
r += 1
for row in cur.fetchall():
    sheet2.write(r,0,row[0])
    sheet2.write(r,1,row[1])
    r += 1

cur.execute('select plan,count(1) from checktrac where flag4trac = 0 group by plan')
r += 1
sheet2.write(r,0,u'计划名称')
sheet2.write(r,1,u'记录位置为空的任务数')
r += 1
for row in cur.fetchall():
    sheet2.write(r,0,row[0])
    sheet2.write(r,1,row[1])
    r += 1

cur.execute('select plan,count(1) from checktrac where flag4check = 0 group by plan')
r += 1
sheet2.write(r,0,u'计划名称')
sheet2.write(r,1,u'未跟踪检查的任务数')
r += 1
for row in cur.fetchall():
    sheet2.write(r,0,row[0])
    sheet2.write(r,1,row[1])
    r += 1

cur.execute('select plan,count(1) from checktrac where flag4qualty = 0 group by plan')
r += 1
sheet2.write(r,0,u'计划名称')
sheet2.write(r,1,u'未评定质量的任务数')
r += 1
for row in cur.fetchall():
    sheet2.write(r,0,row[0])
    sheet2.write(r,1,row[1])
    r += 1

cur.close()
conn.commit()
conn.close()
book2.save("/home/rain/下载/执行状况一览表.xls")
