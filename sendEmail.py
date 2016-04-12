# -*- coding: utf-8 -*-

import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import xlsxFormatSetting as settings
import time

reload(sys)
sys.setdefaultencoding('utf8')
reload(sys)

def sendMail():
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    import smtplib

    #创建一个带附件的实例
    msg = MIMEMultipart()

    #构造附件
    for series in settings.carseries:
        print series.decode('UTF8').encode('gbk')
        file_name = series.encode('gbk')+'-报价日报_'.encode('gbk')+localtime+'.xlsx'
        file_path = settings.filePath+'\\report\\'+file_name
        print file_path
        att1 = MIMEText(open('%s' % file_path, 'rb').read(), 'base64', 'gb2312')
        att1["Content-Type"] = 'application/octet-stream'
        att1["Content-Disposition"] = 'attachment; filename="%s"' % file_name
        msg.attach(att1)

    #加邮件头
    msg['to'] = 'peter.zhang@mathartsys.com'
    msg['from'] = 'peter.zhang@mathartsys.com'
    msg['subject'] = '雪佛兰新车上市日报'.encode('GBK')
    #发送邮件
    try:
        server = smtplib.SMTP()
        server.connect('smtp.qiye.163.com')
        server.login('peter.zhang@mathartsys.com','Zz651454')
        server.sendmail(msg['from'], msg['to'],msg.as_string())
        server.quit()
        print '发送成功'.encode('GBK')
    except Exception, e:
        print str(e)


if __name__ == '__main__':

    localtime = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    sendMail()

