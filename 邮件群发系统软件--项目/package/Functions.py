# -*- coding:utf-8 -*-
import random,csv

def Replace_strchar(string,char,target):

    """
    替换字符串中的变量。
    """
    m,n = string.find(char),string.rfind(char)
    string = list(string)
    string[m],string[n] = target
    string[m+1],string[n+1] ='',''
    return ''.join(string)

def Monitor_User(mails_lists):
    """
    根据发件人信息生成监控参数字典
    """
    user_monitor ={}
    for user_para in mails_lists:
        #获取配置时间间隔原始参数
        range_num = [float(nu) for nu in user_para[-1].split('-->')]

        #生成时间间隔数列
        range_nums = [round(random.uniform(*range_num),3) for k in range(int(user_para[5]))]

        #根据发件账号生成参数字典：

        user_monitor[user_para[1]] = {'TotalTimes':int(user_para[5]),'Per_Count':int(user_para[4]),'Remain_Times':int(user_para[5]),'TimeGap_Lists':range_nums,'Last_SendTime':''}
    return user_monitor

def Check_Possiblity(user_monitor):
    #按照配置策略可发邮件数量：
    total_count = 0
    total_time = 0
    for value in user_monitor.values():
        total_count += value['TotalTimes'] * value['Per_Count']
        total_time += sum(value['TimeGap_Lists'])
    return total_count,round(total_time/60)

def Record_CSV(file_name,datas):
    with open('data_record/' + file_name + '.csv','a',newline='') as data:
        data_write = csv.writer(data,dialect='excel')
        data_write.writerow(datas) #datas为数组内容。

def Record_TXT(path,file_name,datas):
    with open(path + "/" + file_name + '.txt','a') as data:
        data.write(datas) #datas为数组内容。