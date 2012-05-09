#!/usr/bin/env python
#-*- coding:utf-8 -*-
"""
@version: d$
@author: U{sam han<mailto:samrain.han@gmail.com>}
@license:
@contact:
@see:
"""

import xml.etree.cElementTree as ET
import sqlite3
import os
from xlwt import *
import xlwt

def printelement(element,parentname,projectname,pmid,tasklist):
    """
        递归方法得到ganttproject文件中Task任务树，并将儿子名称改为父亲的名称+儿子名称
        @param p:
        @type v:
        @return:
        @rtype v:
    """
    for subelement in element.getchildren():
        if subelement.tag == 'task':
            name = parentname+subelement.attrib['name']
#            print subelement.attrib['id'],name.encode('utf-8'),subelement.attrib['start'],subelement.attrib['duration']
            subtasklist=(subelement.attrib['id'],name,subelement.attrib['start'],subelement.attrib['duration'],projectname,'+'+str(int(subelement.attrib['duration'])-1)+' day',pmid)
            tasklist.append(subtasklist)
            printelement(subelement,name+'.',projectname,pmid,tasklist)


"""
    打开一个目录下所有ganttproject文件
"""
dir_name = "/home/rain/下载/gan"
file_list = [f_name for f_name in os.listdir(dir_name) if f_name.endswith('gan')]

"""
    打开一个数据库
"""
conn = sqlite3.connect(":memory:")
#conn = sqlite3.connect("taskdb.sqlite")

cur = conn.cursor()
"""
    生成3张表
    计划任务表  taskfromgantproj
    资源表  resources
    资源分配表  allocations
"""
cur.execute('CREATE TABLE "taskfromgantproj" ("id" INTEGER, "name" VARCHAR, "startdate" DATETIME, "duration" INTEGER,"plan" VARCHAR,"strdur" VARCHAR,"pmid" INTEGER)')
cur.execute('CREATE TABLE "resources" ("id" INTEGER, "name" VARCHAR)')
cur.execute('CREATE TABLE "allocations" ("taskid" INTEGER, "resourceid" INTEGER)')

for f_in_name in file_list:
    tree = ET.ElementTree(file=os.path.join(dir_name,f_in_name))
#    print tree.getroot().tag,tree.getroot().attrib
    projectname = tree.getroot().attrib['name']
    tasks = tree.getroot()[4]
    resources = tree.getroot()[5]
    allocations = tree.getroot()[6]
    tasklist = []
    listresources = []
    listallocations = []
    for subelement in resources.getchildren():
        """
            得到资源信息
        """
    #    print subelement.attrib['id'],subelement.attrib['name'].encode('utf-8')
        if subelement.attrib['function'] == 'Default:1': #得到项目经理的资源id
            pmid = subelement.attrib['id']        
        listresources.append([subelement.attrib['id'],subelement.attrib['name']])

    for subelement in allocations.getchildren():
        """
            得到任务和资源的关联关系
        """
        listallocations.append([subelement.attrib['task-id'],subelement.attrib['resource-id']])
    """
        得到任务信息
    """
    printelement(tasks,'',projectname,pmid,tasklist)
    
    """
        插入数据库,将以上得到的信息插入到各自表中
    """
    cur.execute('delete from taskfromgantproj')
    cur.execute('delete from resources')
    cur.execute('delete from allocations')
    
    cur.executemany('insert into taskfromgantproj values(?,?,?,?,?,?,?)',tasklist)
    cur.executemany('insert into resources values(?,?)',listresources)
    cur.executemany('insert into allocations values(?,?)',listallocations)

    """
        导出计划任务分配表
    """
    cur.execute('SELECT a.plan,a.name taskname,c.name,date(a.startdate,a.strdur) enddate,(select resources.name from resources where resources.id = a.pmid) FROM taskfromgantproj a,allocations b,resources c where a.id = b.taskid and b.resourceid = c.id and a.startdate >= ? and date(a.startdate,a.strdur)<= ?',['2012-05-14','2012-05-18'])
    rows = cur.fetchall()
    r = 1
    book = xlwt.Workbook()
    sheet1 = book.add_sheet("1")
    for row in rows:
        plan = row[0]
        sheet1.write(r,0,plan)
        sheet1.write(r,1,row[1])
        sheet1.write(r,2,row[2])
        sheet1.write(r,3,row[3])
        sheet1.write(r,4,row[4])
        r+=1
    sheet1.write(0,0,u'计划名称')
    sheet1.write(0,1,u'任务名称')
    sheet1.write(0,2,u'执行人')
    sheet1.write(0,3,u'截止时间')
    sheet1.write(0,4,u'检查人')
    sheet1.write(0,5,u'预计工期')

    book.save(os.path.join(dir_name,projectname.encode('utf-8'))+'plan.xls')

cur.close()
#conn.commit()
conn.close()
