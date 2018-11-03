#-*-coding:utf-8-*-
import xlrd
from urllib import request
from bs4 import BeautifulSoup
from urllib.parse import quote
import string
import xlwt
import xlutils.copy
from wxpy import *
import os,shutil


bot = Bot(console_qr=True,cache_path=True)
bot.groups(update=True, contact_only=False)
@bot.register(Friend,ATTACHMENT)

def user_msg(msg):     #定义一个接收消息的函数
    image_name = msg.file_name
    friend = msg.chat
    print(msg.chat)
    print('接收文件:',msg.file_name)
    msg.get_file(''+msg.file_name)
    
    data = xlrd.open_workbook(msg.file_name)
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
                    #print(col_num)
                    score_col_num=col_num-1    #得到总分所在列索引
                    score_row_num=row_num  #得到第一个总分所在行索引    
                    break
                    row_num = 0
        col_num = 0

    print(score_col_num)
    print(score_row_num)

    int_excel = xlutils.copy.copy(data)     #使用xlutils.copy追加数据
    new_table=int_excel.get_sheet(0)            #获取第一张工作表
    for title_num in range(1,21):
        new_table.write(score_row_num-1,ncols-1+title_num,'预测大学'+str(title_num))         #寻找总成绩所在位置，写入表头
    int_excel.save('处理后_'+msg.file_name)


    for per_row_num in range(score_row_num,nrows):
        per_score = table.cell(per_row_num,score_col_num).value
        print(per_score)

        user_score = str(int(per_score))

        url = quote('http://kaoshi.edu.sina.com.cn/college/estimate/view/?local=1&wl=2&score='+user_score+'&provid=0&typeid=1', safe=string.printable)
        headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'}
        page = request.Request(url, headers=headers)

        page_info = request.urlopen(page).read()
        page_info = page_info.decode('utf-8')
        soup = BeautifulSoup(page_info, 'html.parser')
        #print(soup)
        univ_num = 0
        for tag_a in soup.find_all('a'):

            if tag_a.get('style')=="color: #0071ff;":
                #print(tag_a)
                if len(tag_a.get_text())<=8:
                    univ_num = univ_num+1
                    print(tag_a.get_text())
                    new_table.write(per_row_num,ncols-1+univ_num,tag_a.get_text())

        int_excel.save('处理后_'+msg.file_name)
    shutil.move('处理后_'+msg.file_name, '/www/wwwroot/www.defender.ink/score_processed')
    msg.reply('文件处理成功，下载地址:\n'+quote('https://www.defender.ink/score_processed/处理后_'+msg.file_name, safe=string.printable))

bot.join()
