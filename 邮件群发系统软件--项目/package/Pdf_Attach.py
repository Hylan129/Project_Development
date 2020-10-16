#-*- coding : utf-8 -*-

from email.mime.application import MIMEApplication

def uploadPdf(path):
    """
    参数说明：path为本地图片路径,id_name为邮件附件呈现的附件名称，自己指定。
    """

    pdf_part = MIMEApplication(open(path, 'rb').read())
    pdf_part.add_header('Content-Disposition', 'attachment', filename=path.split('/')[-1])
    
    #msgImage["Content-Disposition"] = 'attachment; filename="' + id_name +'.' + path.strip().split('.')[-1] + '"'
    return pdf_part