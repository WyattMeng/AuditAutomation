# -*- coding: utf-8 -*-
"""
Created on Tue Sep 19 14:34:33 2017

@author: Maddox.Meng
"""

from openpyxl import load_workbook

path = r'C:\Workspace\测试\A3.xlsx'.decode('utf-8')
wb = load_workbook(path)
#ws = wb.get_sheet_names() #[u'BS&PL', u'CF']

'''xxxx'''
ws = wb.get_sheet_by_name(u'BS&PL')

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

xdict = {}        
k=0
for sheetname in sheetnames:
    xdict[k] = {'sheetname':sheetname, 'x_min':xlist[k], 'x_max':xlist[k+1]}
    k+=1


for k in xdict:
    
    y=2
    xitems_a = []
    #print 'Below is [' + xdict[k]['sheetname'] + ']:'
    for x in range(xdict[k]['x_min'], xdict[k]['x_max']):
        if ws[x][y].value != None and ws[x][y].value != u'项目':
            xitem = ws[x][y].value
            x     = ws[x][y].row
            xitems_a.append([xitem, x])
            
    y=3
    xitems_b = []
    #print 'Below is [' + xdict[k]['sheetname'] + ']:'
    for x in range(xdict[k]['x_min'], xdict[k]['x_max']):
        if ws[x][y].value != None and isinstance(ws[x][y].value,long) is False and isinstance(ws[x][y].value,float) is False:
            xitem = ws[x][y].value
            x     = ws[x][y].row
            xitems_b.append([xitem, x])

    if len(xitems_a) <= len(xitems_b):
        xitems = xitems_b
        print 'xitems_b'
    else:
        xitems = xitems_a
        print 'xitems_a'                    
    
    print xitems            
            