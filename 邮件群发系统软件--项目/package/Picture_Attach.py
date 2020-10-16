#-*- coding : utf-8 -*-

from email.mime.image import MIMEImage

def uploadPicture(path,id_name):
    """
    参数说明：path为本地图片路径,id_name为邮件正文调用图片进行网页呈现时的调用名称，自己指定。
    """

    # 指定图片为当前目录
    fp = open(path, 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    # 定义图片 ID，在 HTML 文本中引用
    msgImage.add_header('Content-ID', '<'+id_name+'>')
    msgImage.add_header('Content-Disposition', 'attachment', filename=path.split('/')[-1])
    #msgImage["Content-Disposition"] = 'attachment; filename="' + path.split('/')[-1] + '"'
    return msgImage