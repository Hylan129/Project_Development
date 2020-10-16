# -*- coding:utf-8  -*-
import re,xlwt,xlrd,time,traceback

def verifyEmail(email):
    pattern = r'^[\.a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$'
    if re.match(pattern,email) is not None: return True
    else: return False

def prefixDeal(prefix):
    content = re.findall('[a-zA-Z0-9_]+',prefix)
    newcontent = [m.capitalize() for m in content]
    return ' '.join(newcontent)

def emailDeal(email):
    if verifyEmail(email):
        qianzhui,houzhui= email.split('@')
        qianzhui = prefixDeal(qianzhui)
        houzhui = houzhui.split('.')[-1]
        return [email,qianzhui,houzhui,'邮箱地址正常'],True
    else:
        print(" 地址不合法：" + str(email))
        return [email,'邮箱地址异常'],False
    
if __name__ == '__main__':
    try:
        print("\n\n程序已经开始运行！使用过程中有任何问题，请联系VX:hylan129\n\n")
        
        print("倒计时5秒开始数据处理！如需结束运行，请及时点击关闭窗口(避免误点击)。\n\n")
        for i in range(5):
            print(5-i)
            time.sleep(1)
        print("\n\n数据正在处理中.......\n\n")
        
        start_time = time.time()
        
        #读取xls原始数据：
        data = xlrd.open_workbook('待处理表格.xls',formatting_info=False)#打开需要读取的excel表
        table = data.sheets()[0] #提取第0个活页博，即excel中首个活页博
        col_data = table.col_values(0) #取出第4列的数据，生成数组。

        #数据处理，提取前缀和后缀
        data_dealed = []
        data_fail = []
        for data in col_data:
            result,flag = emailDeal(data.strip())
            if flag:
                data_dealed.append(result)
            else:
                data_fail.append(result)
        #数据处理，排序
        final_data = sorted(data_dealed,key = lambda x:x[2])
        
        #异常地址计数：
        unnormal_number = len(data_fail)
        
        #写入xls信息
        myWorkbook = xlwt.Workbook()
        mySheet = myWorkbook.add_sheet('Emails_Information') # 添加活页博
        mySheet1 = myWorkbook.add_sheet('Emails_Abnormal') # 添加活页博

        title = ['No.','邮件地址','用户名称','国家名别','地址合法性判定结果']
        title1 = ['No.','邮件地址','地址合法性判定结果']
        
        #数据写入，写入标题
        for num,content in enumerate(title):
            mySheet.write(0, num, content) 
        
        for num,content in enumerate(title1):
            mySheet1.write(0, num, content)

        #数据循环写入
        for num_row,contents in enumerate(final_data): #content_list_all为需写入的数组数据
            for num_col,content in enumerate(contents):
                mySheet.write(num_row+1,num_col+1,content)
            mySheet.write(num_row+1,0,num_row+1)
        
        for num_row,contents in enumerate(data_fail): #content_list_all为需写入的数组数据
            for num_col,content in enumerate(contents):
                mySheet1.write(num_row+1,num_col+1,content)
            mySheet1.write(num_row+1,0,num_row+1)
        
        myWorkbook.save('Emails_Information_After_Dealed' + '.xls') #保存excle数据表。
        
        end_time = time.time()
        
        print("***总处理邮件地址数量：{}个，异常邮件地址：{}个，耗时：{}秒。\n\n***数据处理详情请查看Emails_Information_After_Dealed.xls表格。".format(
                len(col_data),unnormal_number,round(end_time-start_time,2))
             )
        input("\n程序已执行完成，按Enter键退出！")
    except Exception as e:
        traceback.print_exc(file=open('error.txt','a')) 
