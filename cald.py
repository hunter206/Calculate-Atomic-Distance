#-*- coding: UTF-8 -*- 
# ==============================================================================
#
#       Filename:  cald.py
#    Description: atomic data excel operat
#        Created:  20200709
#         Author:  hunter qq 770896174
#
# ==============================================================================
import xlsxwriter #加载包
import numpy as np
import math

myWorkbook = xlsxwriter.Workbook('test.xlsx')     #opath为目录名，file_name为excel文件名，表示在某路径下创建一个excel文件
sheet_file_realname = myWorkbook.add_worksheet('new') #在文件中创建一个名为file_realname_new的sheet,不加名字默认为sheet1
bold= myWorkbook.add_format({'bold':True})                        #设置一个加粗的格式对象---单元格格式
## 其他格式

'''
para_1 = 10
para_2 = 20
sheet_file_realname.write(0, 0, para_1, bold) #写入方式一，也可采用循环写入
sheet_file_realname.write(0, 1, para_2, bold)
sheet_file_realname.write(0, 2, 'good', bold) #写入字符
sheet_file_realname.write(0, 3, 'people', bold)
sheet_file_realname.write(0, 4, 'hello', bold)
'''
'''第一行标题'''
col = 1
row = 0
zeroarray = np.zeros(251)
#print('zeroarray:', zeroarray)
first_flag = False
data_range = np.linspace(0,5,251)
sheet_file_realname.write(0, 0, 'E', bold)
for i in data_range:
    sheet_file_realname.write(0, col, str(i), bold)
    col += 1

f = open("Au_transfer.xyz")
while 1:
    lines = f.readlines(1)
    line_str = str(lines)
    line_str_sp = line_str.split()
    if line_str.find('E') == 2:#如果是能量值行
        row += 1
        sheet_file_realname.write(row, 0, line_str[6:20], bold)
        first_flag = True
        zeroarray = np.zeros(251)
        zeroarray[0] = 1#每组第一个数是基准
        sheet_file_realname.write(row, 1, str(zeroarray[0]), bold)
    elif line_str.find('A') != -1:#如果是原子坐标行
        auxyz = lines[0].split()
        if first_flag == True:
            X0 = float(auxyz[1])
            Y0 = float(auxyz[2])
            Z0 = float(auxyz[3])
            first_flag = False
        X = float(auxyz[1])
        Y = float(auxyz[2])
        Z = float(auxyz[3])
        distance = math.sqrt((X-X0)**2 + (Y-Y0)**2 + (Z-Z0)**2)
        #判断/分类
        if (distance <=5):
            for j in range(0, 249):
                if ((distance> 0.02*j)&(distance<= 0.02*(j+1))):
                    zeroarray[j] += 1
                    #sheet_file_realname.write(row, i+1, str(distance), bold)
                    sheet_file_realname.write(row, j+1, str(zeroarray[j]), bold)
        else:
                zeroarray[250] += 1
                sheet_file_realname.write(row, 251, str(zeroarray[250]), bold)
        #row += 1
        print(row, np.sum(zeroarray), X,Y,Z)
        
    if not lines:
        break
print('总行数：', row)
f.close()
myWorkbook.close()
