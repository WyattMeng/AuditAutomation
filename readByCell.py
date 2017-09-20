# -*- coding: utf-8 -*-
"""
Created on Tue Sep 19 14:34:33 2017

@author: Maddox.Meng
"""

from openpyxl import load_workbook
import re

from openpyxl.utils import get_column_letter, column_index_from_string# 根据列的数字返回字母


def matchDate(date):
    
    # 将正则表达式编译成Pattern对象
    pattern = re.compile(r'2\d\d\d年')
     
    # 使用Pattern匹配文本，获得匹配结果，无法匹配时将返回None
    match = pattern.match(date.encode('utf-8')) #u'2017\u5e7412\u670831\u65e5'
     
    if match:
        # 使用Match获得分组信息
        #print match.group()
        return True
    if not match:
        return False


path = r'C:\Workspace\测试\risk table.xlsx'.decode('utf-8')
wb = load_workbook(path)
#ws = wb.get_sheet_names() #[u'BS&PL', u'CF']

'''xxxx'''
ws = wb.get_sheet_by_name(u'信用风险')

#openpyxl worksheet[x][y], x从1开始计算，y从0开始计算
y=0

xlist=[]
sheetnames=[]

for x in range(1, ws.max_row):
    if ws[x][y].value != None:
        xlist.append(x)
        sheetnames.append(ws[x][y].value)
        
        #xdict[k] = {'sheetname': }
        #print x, ws[x][y].value, type(ws[x][y].value)
xlist.append(ws.max_row)    

#底稿其中一个sheet的x范围
xdict = {}        
k=0
for sheetname in sheetnames:
    xdict[k] = {'sheetname':sheetname, 'x_min':xlist[k], 'x_max':xlist[k+1]}
    k+=1


for k in xdict:
    
    #对于一个一级科目（一个GET sheet）的Area，判断第3 or 第4列是写有二三级科目的列
    y=2
    xitems_a = []
    #print 'Below is [' + xdict[k]['sheetname'] + ']:'
    for x in range(xdict[k]['x_min'], xdict[k]['x_max']):
        if (
            ws[x][y].value != None and 
            ws[x][y].value != u'项目' and
            ws[x][y].value != u'附注六' and
            ws[x][y].value != u'股东名称' and
            ws[x][y].value != u'关联方名称' and
            matchDate(ws[x][y].value) is False
            ):
            xitem = ws[x][y].value
            x     = ws[x][y].row
            xitems_a.append([xitem, x])
            
    y=3
    xitems_b = []
    #print 'Below is [' + xdict[k]['sheetname'] + ']:'
    for x in range(xdict[k]['x_min'], xdict[k]['x_max']):
        #正则表达式去掉2xxx年的匹配cell
        if (
            ws[x][y].value != None and 
            isinstance(ws[x][y].value,long) is False and 
            isinstance(ws[x][y].value,float) is False and
            matchDate(ws[x][y].value) is False
            ):
            xitem = ws[x][y].value
            x     = ws[x][y].row
            xitems_b.append([xitem, x])

    if len(xitems_a) <= len(xitems_b):
        xitems = xitems_b
        yrows = xitems[0][1] - xdict[k]['x_min'] - 1
        print 'xitems_b'
    else:
        xitems = xitems_a
        print 'xitems_a'  

    yrows = xitems[0][1] - xdict[k]['x_min'] - 1
    
    
    yrowslist = []
    for i in range(0, yrows):
        yrowslist.append(xdict[k]['x_min'] + i)
        #print yrows,xitems[0][1],xdict[k]['x_min']                  
    
    '''yitems'''
    #去掉项目、附注六、股东名称、关联方名称、年份, yitems行数 = 二级科目行 - 一级科目行 - 1
    print yrowslist
    
    ylist =[]
    for x in yrowslist:
        for y in range(3, ws.max_column): #预知从列3开始
            if ws[x][y].value != None:
                ylist.append(ws[x][y].column)
                print ws[x][y].coordinate, ws[x][y].value 
                print ylist
                print list(set(ylist)).sort()
            
for item in ws.merged_cell_ranges:
    print item.split(':')[0]            