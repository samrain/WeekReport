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
import os
#import pdb
"""
    打开一个数据库
"""
conn = sqlite3.connect('taskdb.sqlite')
#conn = sqlite3.connect(":memory:")'数据库放在内存中
cur = conn.cursor()

"""
    打开一个XLS文件
"""
#xlsfile = xlrd.open_workbook('test.xls')

"""
    打开一个目录下所有xls文件
"""
dir_name = "/home/rain/下载"
file_list = [f_name for f_name in os.listdir(dir_name) if f_name.endswith('xls')]
#print file_list
for f_in_name in file_list:
    print f_in_name
    xlsfile = xlrd.open_workbook(os.path.join(dir_name,f_in_name))
    """
        根据sheet名称，打开XLS文件中某个Sheet
    """
    worksheet = xlsfile.sheet_by_name('listdetail')
    """
        取得每个Cell的值
    """
    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1
    #curr_row = -1
    """
        从第3行开始取数据
    """
    curr_row = 2
    """
        循环取数据
    """
    #startdate = time.strftime('%Y-%m-%d',(2012,3,12,0,0,0,0,0,0))
    taskinfo = []
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        print 'Row:', curr_row
    #    print row
        curr_cell = -1
        rowinfo = []
        flag4finish = 0#默认都是没有完成
        flag4delay = 0#默认没有逾期
        flag4trac = 2#默认尚未开展跟踪
        flag4check = 0#默认没有检查
        while curr_cell < num_cells:
            curr_cell += 1
            # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
            cell_type = worksheet.cell_type(curr_row, curr_cell)
            cell_value = worksheet.cell_value(curr_row, curr_cell)
            startdate = datetime.date(2012,3,12)
            d = datetime.timedelta(days=curr_cell-7)
            if cell_type == 3:
                tuplex4xls = list(xlrd.xldate_as_tuple(cell_value,0))
                tuplex4xls.append(0)
                tuplex4xls.append(0)
                tuplex4xls.append(0)
                cell_value = time.strftime('%Y-%m-%d',tuplex4xls)
    #        elif cell_type == 1:
    #            print cell_value.encode('utf-8')#如果是字符串，转换成utf-8
            if curr_cell in (10,12,14,16,18,20,22):
    #            pdb.set_trace()
                if cell_value == 100:
                    flag4finish = 1#表示完成
    #                print datetime.datetime.strftime(t,'%Y-%m-%d')
    #                print rowinfo[2],startdate + d
                    if rowinfo[2] >= datetime.datetime.strftime((startdate + d),'%Y-%m-%d'):
                        flag4delay = 1#逾期
                    if rowinfo[7] == '':
                        flag4trac = 0#跟踪失败，没有记录位置
                    else:
                        flag4trac = 1#跟踪成功
                        if rowinfo[5] == '':
                            flag4check = 0#没有检查
                        else:
                            flag4check = 1#做过检查
    #        print '    ', cell_type, ':', cell_value
            rowinfo.append(cell_value)
    #    print rowinfo
        rowinfo.append(flag4finish)
        rowinfo.append(flag4delay)
        rowinfo.append(flag4trac)
        rowinfo.append(flag4check)
        cur.execute('insert into task values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',rowinfo)
cur.close()
conn.commit()
conn.close()
