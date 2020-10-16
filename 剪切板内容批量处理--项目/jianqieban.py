# -*- coding:utf-8 -*-

import win32con,xlrd,xlwt
import win32clipboard as wincld

def setCopy(content):
    result_text = content.strip()
    wincld.OpenClipboard()
    wincld.EmptyClipboard()
    wincld.SetClipboardData(win32con.CF_UNICODETEXT, result_text)
    wincld.CloseClipboard()
    
def get_xlsdata(path):
    try:
        data = xlrd.open_workbook(path)#打开需要读取的excel表
        return data
    except Exception as e:
        print('Warning：表格配置数据读取出错！')
        with open('error.txt','a') as code:
            code.write("get_xlsdata()出错：" + str(e)+"\n\n")
        return ''
        
if __name__ == '__main__':
    try:


        #读取原始表格数据
        datas = get_xlsdata(input('\n请输入需要处理的excel表格文件名称，包含后缀: '))

        #原始数据：
        data_all = datas.sheets()[0]

        #功能选择：
        options = input('\n请选择功能：表格单元格内容长度控制，请输入1；表格单元格内容依次复制，请输入2。\n')

        if options == '1':

            #指定需要控制长度的列号和控制长度
            control_liehao = [num.split(',') for num in input('\n请输入需要控制长度的列号和对应的控制长度（举例:3,20#4,15）:').split('#')]

            #新建表格
            myWorkbook = xlwt.Workbook()
            mySheet = myWorkbook.add_sheet('0') # 添加活页博

            for data  in range(data_all.ncols):
                content = data_all.col_values(data)
                for liehao_num,length in control_liehao: 
                    if data == int(liehao_num):
                        content_dealed = [mcontent[0:int(length)] for mcontent in content]
                for num in range(len(content_dealed)):
                    mySheet.write(num,data,content_dealed[num])

            myWorkbook.save('长度处理后' + '.xls') #保存excle数据表。

            print("\n\n表格中单元格内容长度控制已经处理完成！\n\n")

        elif options == '2':

            #指定剪切复制的列号数据
            liehao = int(input('\n请输入需要依次复制的数据列号，整数: ')) - 1

            print('\n\n 程序已经开始运行，数据依次复制进行中...\n\n')

            for data  in range(data_all.nrows):
                content = data_all.row_values(data)[liehao].strip()
                setCopy(content)
                print(data + 1,data_all.row_values(data),'已复制到剪切板！请粘贴后按ENTER键继续。')
                input()
            print('\n程序已经执行完成，请点击ENTER结束，谢谢！\n')

    except Exception as e:
        print('程序出错，已结束运行！请查看error文件确认错误原因。')
        with open('error.txt','a') as code:
            code.write('\n' + str(e) + '\n')
    
        
        
    