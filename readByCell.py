# -*- coding: utf-8 -*-
"""
Created on Tue Sep 19 14:34:33 2017

@author: Maddox.Meng
"""

'''之前用xlrd、pandas读取，这次尝试用单一模块openpyxl搞定一切

'''

from openpyxl import load_workbook
import re

from openpyxl.utils import get_column_letter, column_index_from_string# 根据列的数字返回字母
#column_index_from_string('aa') == 27

'''正则表达式match日期'''
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


'''get一级条目的单元格value list，rownumber list。进而从底稿get每个一级条目的高度。'''
#openpyxl worksheet[x][y], x从1开始计算，y从0开始计算
def getCatalogBands(worksheet, catalogColNo):
    rownumbers=[]
    sheetnames=[]
    y = catalogColNo                      #假定已知第0列是存放底稿一级条目（GET sheetname）的列
    for x in range(1, worksheet.max_row): #openpyxl的行从1起计，所以这里range从1开始。
        if worksheet[x][y].value != None: #append非空单元格value、rownumber
            rownumbers.append(x)
            sheetnames.append(worksheet[x][y].value)
            
    rownumbers.append(worksheet.max_row)  #最后append上max_row，为了下面dict好计算。因为最后一行一级条目的最大行好就是整个worksheet的最大行号。

    xCatalogBands = {}    #底稿其中一个一级条目的value（也是GET的sheetname）、起始rownumber、结束rownumber字典    
    k=0
    for sheetname in sheetnames:
        xCatalogBands[k] = {'sheetname':sheetname, 'x_min':rownumbers[k], 'x_max':rownumbers[k+1]}
        k+=1
    return xCatalogBands   


def getXtitles(xCatalogBands):
    '''得到xtitles，包括title name和title行号'''    
    y=2  #对于一个一级科目（一个GET sheet）的Area，判断第3 or 第4列是写有二三级科目的列
    xtitles_a = []
    #print 'Below is [' + xdict[k]['sheetname'] + ']:'
    for x in range(xCatalogBands[k]['x_min'], xCatalogBands[k]['x_max']):        
        if (
            ws[x][y].value != None and 
            ws[x][y].value != u'项目' and
            ws[x][y].value != u'附注六' and
            ws[x][y].value != u'股东名称' and
            ws[x][y].value != u'关联方名称' and
            matchDate(ws[x][y].value) is False
            ):#去掉项目、附注六、股东名称、关联方名称、年份, yitems行数 = 二级科目行 - 一级科目行 - 1
            xitem = ws[x][y].value
            x     = ws[x][y].row
            xtitles_a.append([xitem, x])
            
    y=3  #对第4列循环
    xtitles_b = []
    #print 'Below is [' + xdict[k]['sheetname'] + ']:'
    for x in range(xCatalogBands[k]['x_min'], xCatalogBands[k]['x_max']):
        #正则表达式去掉2xxx年的匹配cell
        if (
            ws[x][y].value != None and 
            isinstance(ws[x][y].value,long) is False and 
            isinstance(ws[x][y].value,float) is False and
            matchDate(ws[x][y].value) is False
            ):
            xitem = ws[x][y].value
            x     = ws[x][y].row
            xtitles_b.append([xitem, x])

    if len(xtitles_a) <= len(xtitles_b):  #如果列3非空个数 小于等于 列4非空非数字个数，则列4为title列
        xtitles = xtitles_b
    else:                                 #反之，列2为title列
        xtitles = xtitles_a
        
    return xtitles    


def getYRows(xtitles, xCatalogBands): 
    
    '''得到ytitles行数（在后面起到作用）'''
    yrows = xtitles[0][1] - xCatalogBands[k]['x_min'] - 1  #根据excel模板规律，ytitle行数 = 第1个xtitle行号 - 所在catalog行号 - 1
    
    '''得到ytitles，包括title name和title列号'''
    yrowslist = []
    for i in range(0, yrows):
        yrowslist.append(xCatalogBands[k]['x_min'] + i)
    return yrowslist           


def getYColumnNo():
    for x in yrowslist:
        for y in range(4, ws.max_column): #ytitle都在第5、6列起始。特例：R利率敏感性分析，没有ytitle。
            
            



'''----------------------------------------------------------------------------------------------------'''
path = r'C:\Workspace\测试\risk table.xlsx'.decode('utf-8') #r和decode utf-8组合使用可以解决中文路径的issue
wb = load_workbook(path)
#ws = wb.get_sheet_names() #[u'BS&PL', u'CF']
ws = wb.get_sheet_by_name(u'信用风险')
xCatalogBands = getCatalogBands(ws,0)

for k in xCatalogBands:
    xtitles = getXtitles(xCatalogBands)
    yrowslist = getYRows(xtitles, xCatalogBands)
    for merged_cells in ws.merged_cell_ranges:
#        print merged_cells
#        print merged_cells.split(':')[0][1:],xCatalogBands[k]['x_min'], xCatalogBands[k]['x_max'] 
        if int(merged_cells.split(':')[0][1:]) in range(xCatalogBands[k]['x_min'],  xCatalogBands[k]['x_max']):
            print merged_cells

    print '==========='
    
    
#    '''yitems'''    
#    ylist =[]
#    for x in yrowslist:
#        for y in range(3, ws.max_column): #预知从列3开始
#            if ws[x][y].value != None:
#                ylist.append(ws[x][y].column)
#                print ws[x][y].coordinate, ws[x][y].value 
#                print ylist
#                print list(set(ylist)).sort()
            
#for item in ws.merged_cell_ranges:
#    print item.split(':')[0]            