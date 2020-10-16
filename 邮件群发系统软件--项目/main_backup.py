# -*- coding:utf-8 -*-
 
from package import Excel_Read,Time_Deal,Email_Custom,Pdf_Attach,Picture_Attach,IP_Check,Functions
from email.utils import formataddr
import traceback,csv,prettytable
import sys

if __name__ == "__main__":


    sys.setrecursionlimit(100000)
    flag_para = True #配置参数配置情况
    flag_sendmail_fail = True #发件出错情况
    try:
        print("\n\n程序已经开始运行！使用过程中有任何问题，请联系VX:hylan129\n\n")
                    
        print("倒计时5秒开始检查软件参数配置情况！如需结束运行，点击关闭窗口即可。\n\n")
        for i in range(5):
            print(5-i)
            Time_Deal.time.sleep(1)
        print("\n\n开始检查相关配置参数.......\n\n")

        #读取表格    
        rawdata = Excel_Read.get_xlsdata("settings/setting_informations.xls")

        #读取活页簿 
        data_sheets = ['mail_content','mailto_lists','mails_lists','socks_ip']
        mail_content,mailto_lists,mails_lists,socks_ip = [rawdata.sheet_by_name(name) for name in data_sheets]

        #邮件内容活页簿处理，数据提取
        mail_subject,mail_content,attach_pictures,attach_pdfs = mail_content.col_values(1)[0:4]

        #邮件正文内容数据处理，及正文格式确定
        if mail_content == '':
            content = open("settings/content.html",encoding='utf-8').read().replace("\n",'')
            content_class = 'html'
        else:
            content = mail_content
            content_class = 'plain'

        #邮件附件数据处理，添加所有附件
        attach_pictures,attach_pdfs= [pic for pic in attach_pictures.split("#")],[pic for pic in attach_pdfs.split("#")],
        #添加所有图片
        picture_attaches = [Picture_Attach.uploadPicture("attach_files/" + pic,'') for pic in attach_pictures]
        #添加所有pdf文件
        pdf_attaches = [Pdf_Attach.uploadPdf("attach_files/" + pdf)  for pdf in attach_pdfs]

        #收件人清单
        mailto_lists = [mailto_lists.row_values(num) for num in range(1,mailto_lists.nrows)]

        #发件人清单
        mails_lists =  [mails_lists.row_values(num) for num in range(1,mails_lists.nrows)]

        #代理清单
        socks_ip =  [socks_ip.row_values(num) for num in range(1,socks_ip.nrows)]

        #发件正文内容进行判定，是发html文件，还是纯文本。
        if mail_content == '{}{}':
            content = open("settings/content.html",encoding='utf-8').read().replace("\n",'')
            content_class = 'html'
        else:
            content = mail_content
            content_class = 'plain'

        #数据存储路径预检查：
        Functions.Record_TXT('data_record','data_record','#'*20 + Time_Deal.getTimeNow()+ "文档存储路径配置确认OK。" + '#'*20 + "\n")
        Functions.Record_TXT('Error','error','#'*20 + Time_Deal.getTimeNow()+ "软件log存储路径配置确认OK。" + '#'*20 + "\n")

        # 初始化文档存储csv文件
        file_record_names = ['发件成功记录','发件失败记录','邮件群发整体报告']
        file_name_start_time = Time_Deal.getTimeNow()[0:-3].replace(':','时') + "分"
        file_records = [file_name_start_time + "_"+name for name in file_record_names]
        #成功发件记录
        Functions.Record_CSV(file_records[0],['No.','发件时间','使用外网IP地址','发件账号地址','收件人数量','收件账号地址清单','发件情况','备注'])
        #失败发件记录
        Functions.Record_CSV(file_records[1],['No.','发件时间','发件账号','收件人数量','收件账号地址清单','发件情况','原因','备注'])
        #发件整体情况报告
        Functions.Record_CSV(file_records[2],['No','发件账号','发件成功地址数量','配置发送次数剩余数量','备注'])

        #生成监各发件账号的发件情况监控字典；email主体，参数总发次数，每次发送数据，余次，时间间隔，上次发送时间戳，
        user_monitor = Functions.Monitor_User(mails_lists)

        total_count,total_time = Functions.Check_Possiblity(user_monitor)
        #软件配置情况评估说明：
        instructions = "\n{}\n\n发件配置策略评估：\n1、总需求发件对象邮件地址数量：{}个；发件箱配置数量：{}个；\n2、根据发件配置策略评估发件相关情况：\n    总共可发件数量：{}个邮件地址，完成{}个地址发件总耗时约{}分钟；\n{}\n\n{}\n"
        
        if total_count >= len(mailto_lists):
            conclusion = "\n结论：可发件数量大于目标发件数量，配置策略可行。请点击Enter键开始发件，如需重新配置发件策略，请直接关闭窗口！"
        else :
            conclusion = "\n结论：可发件数量小于目标发件数量，建议优化配置策略。\n\n如需重新配置发件策略，请直接关闭窗口；如坚持采用该策略，请点击Enter键开始发件"

        print(instructions.format('#'*60,len(mailto_lists),len(mails_lists),total_count,total_count,total_time,conclusion,'#'*60))
        
    except Exception as e:
        flag_para = False
        Functions.Record_TXT('.','setting_error','*'*40 + Time_Deal.getTimeNow() + '*'*40 + '\n'+str(e) + "\n")
        traceback.print_exc(file=open('setting_error.txt','a'))
        Functions.Record_TXT('.','setting_error','*'*108 + '\n')
        print("软件参数配置出错，请重新检查！具体出错原因请查看setting_error.txt log文件！\n\n")
        input('请点击Enter键退出程序！\n')
    
    try:
        if flag_para:
            input('请点击Enter键继续！')
            print("\n\n开始邮件发送.......\n\n")
            #开始邮件发送：

            #检查网络情况：
            current_ip = IP_Check.getcurrentip()
            if current_ip:
                print("当前网络外网IP地址为：" + current_ip +"，网络正常！")
            else:
                input("当前网络异常，如确认网络正常，请点击Enter继续！")
            #发件成功次数，发件失败次数
            Succeed_Number,Failed_Number = 0,0
            while mailto_lists:
            
                #循环检查每一个发件箱
                for mails in mails_lists:

                    #发件人信息：发件地址，授权码，发件人昵称
                    user,password,user_name = mails[1],mails[2],mails[3]
                    
                    #确定发件账号是否有余次，如果没有余次则进行下一个账号；
                    if not user_monitor[user]['Remain_Times'] > 0: continue
                    
                    #判定是否符合发件条件
                    #条件有两种情况：a、没发过，首次发；b、发过且时间间隔已到
                    if user_monitor[user]['Last_SendTime'] == '' or (Time_Deal.time.time() - user_monitor[user]['Last_SendTime'] >= user_monitor[user]['TimeGap_Lists'][0]) :
                        
                        #发件人昵称处理
                        user_name = formataddr([user_name,user])
                        
                        #发件箱单次发件收件人数量
                        person_count = user_monitor[user]['Per_Count']
                        #符合发件条件后，区分是单发还是群发；
                        if  person_count == 1:

                            #取出收件人信息，一个收件人
                            customs_one = mailto_lists[0]
                            
                            #收件人处理，个性化收件人名称
                            touser = [formataddr([customs_one[1],customs_one[2]])]
                            
                            #个性化邮件主题，添加变量
                            custom_subject = mail_subject.format(*customs_one[3:5])
                            #个性化正文内容，添加变量
                            custom_content = Functions.Replace_strchar(content,'{}',customs_one[5:])

                            try:
                                #邮件正式开始发送
                                Email_Custom.sendEmail(user.split('@')[-1],user_name,user,password,custom_subject,custom_content,content_class,touser,picture_attaches,pdf_attaches)

                            except Exception as e:
                                flag_sendmail_fail = False
                                #写入失败发件的数据：
                                Failed_Number += 1
                                Functions.Record_CSV(file_records[1],[Failed_Number,Time_Deal.getTimeNow(),IP_Check.getcurrentip(),user,person_count,customs_one[2],'Failed！',str(e),'1对1发件！'])

                                #写入错误信息
                                Functions.Record_TXT('Error','error','*'*50 + Time_Deal.getTimeNow() + '*'*50 + '\n'+str(e) + '\n')
                                traceback.print_exc(file=open('Error/error.txt','a'))
                                Functions.Record_TXT('Error','error','*'*119 + '\n')

                            
                            if flag_sendmail_fail:

                                Succeed_Number +=1

                                print(Succeed_Number,Time_Deal.getTimeNow(),"From:"+user,"To:"+customs_one[2],"1对1发送成功！")
                                    
                                #写入成功发件的数据
                                Functions.Record_CSV(file_records[0],[Succeed_Number,Time_Deal.getTimeNow(),IP_Check.getcurrentip(),user,person_count,customs_one[2],'Succeeded','1对1发件！'])
                                
                                #修改账号参数
                                user_monitor[user]['Remain_Times'] -= 1
                                user_monitor[user]['Last_SendTime'] = Time_Deal.time.time()
                                user_monitor[user]['TimeGap_Lists'].pop(0)
                                mailto_lists = mailto_lists[1:]
                            
                            else:
                                #修改账号参数
                                user_monitor[user]['Last_SendTime'] = Time_Deal.time.time()
                                mailto_lists = mailto_lists[1:]

                        
                        elif person_count > 1:
                        
                            #取出收件人信息，Percount个收件人
                            customs_multiple = mailto_lists[0:person_count]
                            
                            #收件人处理，个性化收件人名称
                            tousers = [formataddr([customs_one[1],customs_one[2]]) for customs_one in customs_multiple]
                            tousers_noname = [customs_one[2] for customs_one in customs_multiple]
                            #个性化邮件主题，添加变量
                            custom_subject = mail_subject.format('','')
                            #个性化正文内容，添加变量
                            custom_content = Functions.Replace_strchar(content,'{}',['',''])

                            try:
                                #邮件正式开始发送
                                Email_Custom.sendEmail(user.split('@')[-1],user_name,user,password,custom_subject,custom_content,content_class,tousers,picture_attaches,pdf_attaches)

                            except Exception as e:
                                flag_sendmail_fail = False
                                #写入失败发件的数据：
                                Failed_Number += 1
                                Fail_Time = Time_Deal.getTimeNow()
                                IP = IP_Check.getcurrentip()
                                for mail_fail in tousers_noname:
                                    Functions.Record_CSV(file_records[1],[Failed_Number,Fail_Time,IP,user,person_count,mail_fail,'Failed！',str(e),'1对多发件！'])
                                #写入错误信息
                                Functions.Record_TXT('Error','error','*'*50 + Time_Deal.getTimeNow() + '*'*50 + '\n'+str(e) + '\n')
                                traceback.print_exc(file=open('Error/error.txt','a'))
                                Functions.Record_TXT('Error','error','*'*119 + '\n')

                            if flag_sendmail_fail :

                                Succeed_Number +=1

                                print(Succeed_Number,Time_Deal.getTimeNow(),"From:"+user,"1对"+ str(person_count) + "发送成功！")

                                #写入成功发件的数据
                                Functions.Record_CSV(file_records[0],[Succeed_Number,Time_Deal.getTimeNow(),IP_Check.getcurrentip(),user,person_count,';'.join(tousers_noname),'Succeeded','1对多发件！'])

                                #修改账号参数
                                user_monitor[user]['Remain_Times'] -= 1
                                user_monitor[user]['Last_SendTime'] = Time_Deal.time.time()
                                user_monitor[user]['TimeGap_Lists'].pop(0)
                                mailto_lists = mailto_lists[person_count:]

                            else:
                                user_monitor[user]['Last_SendTime'] = Time_Deal.time.time()
                                mailto_lists = mailto_lists[person_count:]
                            
                        if not mailto_lists: break
    except Exception as e:

        print('软件程序因故障已停止运行,请检查Error文件查看具体错误原因！谢谢！')
        #写入错误信息
        Functions.Record_TXT('Error','error','*'*50 + Time_Deal.getTimeNow() + '*'*50 + '\n'+str(e) + '\n')
        traceback.print_exc(file=open('Error/error.txt','a'))
        Functions.Record_TXT('Error','error','*'*119 + '\n')

    finally :
        try:
            if 'user_monitor' in vars():
                informations = []
                informations.append(['No','From_Email','Su_Qtys','T_Times','Rem_Times','Remark'])
                for num,information in enumerate(user_monitor.keys()):
                    Send_Times = user_monitor[information]['TotalTimes'] - user_monitor[information]['Remain_Times']

                    if user_monitor[information]['Remain_Times'] == 0:
                        marknote = "All used!"
                    else:
                        marknote = "Part used!"
                    informations.append([num+1,information,user_monitor[information]['Per_Count'] * Send_Times,user_monitor[information]['TotalTimes'],user_monitor[information]['Remain_Times'],marknote])

                results = prettytable.PrettyTable()
                results.field_names = informations[0]
                for info in informations[1:]:
                    Functions.Record_CSV(file_records[2],info)
                    results.add_row(info)
                
                count_jing = 70
                print("\n"+"#"* count_jing + "\n\n" + "本次邮件发件汇总情况如下：\n")
                print(results)
                print("\n"+"#"*count_jing + "\n" +"#"*count_jing)
                input("本次发件完毕，请点击Enter键退出！") 
        except Exception as e:
            print('软件程序已运行完毕！谢谢！')
            #写入错误信息
            Functions.Record_TXT('Error','error','*'*40 + Time_Deal.getTimeNow() + '*'*40 + '\n'+str(e) + "\n")
            traceback.print_exc(file=open('Error/error.txt','a'))
            Functions.Record_TXT('Error','error','*'*108 + '\n')
            input("请点击Enter键退出！")