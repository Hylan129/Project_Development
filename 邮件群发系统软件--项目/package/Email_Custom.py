#-*- coding : utf-8 -*-
import smtplib,socks
from . import Time_Deal
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.utils import formataddr

def sendEmail(email_class,email_name,email,password,email_subject,content_page,towhos,pictures=None,pdfs=None,proxies=None):
    """
    ##sendEmail函数参数注释
        #email_class --> 发件箱使用的客户端，qq还是163；字符串；
        #email --> 发件人账号，字符串
        #password -->账号客户端授权码，字符串
        #email_subject -->邮件主题，字符串
        #html_page -->邮件正文，html 
        #towhos --> 发件人清单，为数组list
        #pictures --> 图片附件
        #pdfs --> pdf文件附件
        #proxies -- > socks代理ip及端口，例如：('60.174.233.195', 1080)
    
    """
    # 第三方 SMTP 服务
    mail_host="smtp."+email_class  #设置服务器
    mail_user= email   #用户名
    mail_pass= password   #口令 
    
    message = MIMEMultipart()
    sender = email
    receivers = towhos # 接收邮件，可设置为你的QQ邮箱或者其他邮箱
    
    message['From'] = email_name
    message['To'] = ';'.join(receivers)

    #添加设置邮件主题
    subject = email_subject
    message['Subject'] = Header(subject, 'utf-8')

    #添加设置邮件正文，网页表单。
    
    message.attach(MIMEText(content_page, 'html', 'utf-8'))

    if pictures:
        for pic in pictures:
            message.attach(pic)
    if pdfs:
        for pdf in pdfs:
            message.attach(pdf)
    
    if email_class == "qq.com":
        hostport = 25
    else:
        hostport = 25

    if not proxies:
        smtpObj = smtplib.SMTP() 
        smtpObj.connect(mail_host, hostport)    #qq:465,163:25
        smtpObj.login(mail_user,mail_pass)
        smtpObj.sendmail(sender, receivers, message.as_string())
    else:
        proxy_ip,proxy_port = proxies
        socks.setdefaultproxy(socks.PROXY_TYPE_SOCKS4, proxy_ip, proxy_port)
        socks.wrapmodule(smtplib)
        smtpObj = smtplib.SMTP()
        smtpObj.connect(mail_host, hostport)    #qq:465,163:25
        smtpObj.login(mail_user,mail_pass)
        smtpObj.sendmail(sender, receivers, message.as_string())
    
    return True