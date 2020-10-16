#-*- coding : utf-8 -*-
import xlrd
from . import Time_Deal

def get_xlsdata(path):
    try:
        data = xlrd.open_workbook(path)#打开需要读取的excel表
        return data
    except Exception as e:
        print('Warning：表格配置数据读取出错！')
        with open('error.txt','a') as code:
            code.write(Time_Deal.getTimeNow())
            code.write("get_xlsdata()出错：" + str(e)+"\n\n")
        return ''
    #customers_information = data.sheet_by_name('Sheet1') #提取第0个活页博，即excel中首个活页博