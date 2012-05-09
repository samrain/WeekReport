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

book = xlwt.Workbook()
sheet1 = book.add_sheet("1")

"""
    打开一个数据库
"""
conn = sqlite3.connect('taskdb.sqlite')
cur = conn.cursor()

cur.execute('SELECT a.plan,a.name taskname,c.name,date(a.startdate,a.strdur) enddate,(select resources.name from resources where resources.id = a.pmid) FROM taskfromgantproj a,allocations b,resources c where a.id = b.taskid and b.resourceid = c.id')
rows = cur.fetchall()
r = 1
for row in rows:
    plan = row[0]
    sheet1.write(r,0,plan)
    sheet1.write(r,1,row[1])
    sheet1.write(r,2,row[2])
    sheet1.write(r,3,row[3])
    sheet1.write(r,4,row[4])
    r+=1
cur.close()
conn.close()

sheet1.write(0,0,u'计划名称')
sheet1.write(0,1,u'任务名称')
sheet1.write(0,2,u'执行人')
sheet1.write(0,3,u'截止时间')
sheet1.write(0,4,u'检查人')
sheet1.write(0,5,u'预计工期')

book.save("/home/rain/下载/"+plan.encode('utf-8')+".xls")
