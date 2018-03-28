#coding:utf-8

from openpyxl import load_workbook #可以用来载入已有数据表格

from openpyxl import Workbook #可以用来处理新的数据表格

import datetime #可以用来处理时间相关的数据

def combine():
    '''
    该函数可以用来处理原数据文件：
    1. 合并表格写入的combine表中
    2.保存原数据文件
    '''
    #载入excel文件
    wb1 = load_workbook('courses.xlsx')
    #获取excel文件中的所有sheet名的列表
    sheet_list = wb1.get_sheet_names()
    #读取名为students和time的 sheet 页
    sheet1 = wb1['students']
    sheet2 = wb1['time']
    #如果文当中没有combine这个sheet，则创建
    if 'combine' not in sheet_list:
        wb1.create_sheet('combine',index=2)
    
    #读取最大行数，最大列数
    max_row_s1 = sheet1.max_row
    max_column_s1 = sheet1.max_column

    max_row_s2 = sheet2.max_row
    max_column_s2 = sheet2.max_column

    for i in range(1,max_row_s1+1):
        for j in range(1,max_column_s1+1):  #chr(97)='a'
           # n = chr(j)
           # bh='%s%d'%(j,i)
            print(sheet1.cell(row=i,column=j).value,end=" ")
        print()


    #test：获取1，10行中的第2列数据
#    li=[]
#    for row_num in range(1,10):
#        li.append(sheet1.cell(row=row_num,column=2).value)
#    print(li)

    


def split():
    '''
    该函数可以用来分割文件:
    1.读取combine表中的数据
    2.将数据按时间分割
    3.写入不同的数据表中
    '''
    pass


#执行
if __name__ == '__main__':
    combine()
    split()

