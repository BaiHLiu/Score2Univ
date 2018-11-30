#-*-coding:utf-8-*-
from aip import AipImageClassify
import os.path
from aip import AipOcr
from urllib import request
from bs4 import BeautifulSoup
from urllib.parse import quote
import string
import numpy as np
import os,shutil
from wxpy import *
from aip import AipFace
import base64
import shutil
import xlrd
import xlwt
import xlutils.copy
import time
import shutil


data = xlrd.open_workbook('test.xls')
database=xlrd.open_workbook('database.xls')
database = database.sheet_by_name('sheet1')


table = data.sheets()[0]   #工作表索引
nrows = table.nrows
ncols = table.ncols
row_num = 0
col_num = 0
#for row in range(nrows):
for row in range(21):
    row_num = row_num+1
    #print(table.row_values(row))   #每行返回一个list
    for col in table.row_values(row):
        col_num = col_num+1
        if not str(col).strip()=='':
            if str(col)[0] == '总':
                #print(nrows)
                #msg.reply('服务器正在高速处理中，大约需要'+str(2*nrows)+'秒，请耐心等待..\n*本消息由服务器自动发出，若不慎打扰请谅解')
                #msg.reply(2*nrows)
                #msg.reply('1')
                score_col_num=col_num-1    #得到总分所在列索引
                score_row_num=row_num  #得到第一个总分所在行索引    
                break
                row_num = 0
    col_num = 0

#print(score_col_num)
#print(score_row_num)

int_excel = xlutils.copy.copy(data)     #使用xlutils.copy追加数据
new_table=int_excel.get_sheet(0)            #获取第一张工作表
for title_num in range(1,21):
    new_table.write(score_row_num-1,ncols-1+title_num,'预测大学'+str(title_num))         #寻找总成绩所在位置，写入表头
new_file_name = 'ok.xls'

print(new_file_name)
int_excel.save(new_file_name)


for per_row_num in range(score_row_num,nrows):
    per_score = table.cell(per_row_num,score_col_num).value
    print(per_score)
    
    univ_num=0
    for db_univ_name in database.row_values(int(per_score)):
        univ_num=univ_num+1

        new_table.write(per_row_num,ncols-1+univ_num,db_univ_name)


    int_excel.save(new_file_name)
