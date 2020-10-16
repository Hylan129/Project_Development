# -*- coding:utf-8 -*-
import xlrd,re,csv,traceback,time

if __name__ == '__main__':
    try:
        print("\n\n程序已经开始运行！使用过程中有任何问题，请联系VX:hylan129\n\n")
        
        print("倒计时5秒开始数据处理！如需结束运行，请及时点击关闭窗口。\n\n")
        for i in range(5):
            print(5-i)
            time.sleep(1)
        print("\n\n数据正在处理中.......\n\n")
        
        start_time = time.time()
        
        pattern = r'([\.a-zA-Z0-9_-]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+)'
        data_raw = xlrd.open_workbook('邮箱并排处理.xlsx')#打开需要读取的excel表
        table = data_raw.sheets()[0] #提取第0个活页博，即excel中首个活页博

        datas = {}
        number = 0
        for i in range(table.nrows):
            company,mails = table.row_values(i)
            with open('Email_Lists.csv','a',newline='') as code:
                m = csv.writer(code,dialect='excel')
                for mail,back in re.findall(pattern,mails):
                    m.writerow([number+1,company,mail]) #写入数据
                    number += 1
        
        end_time = time.time()
        
        print('\n数据处理完毕！！！\n\n')
        print("***总邮件地址数量：{}个，处理耗时：{}秒。\n\n***数据处理详情请查看Emails_Lists.csv表格。".format(
            number,round(end_time-start_time,2)))
        input("\n程序已执行完成，按Enter键退出！")
    except Exception as e:
        traceback.print_exc(file=open('error.txt','a')) 
        print('\n程序出错，已结束运行！具体原因清查看error文件。\n')
