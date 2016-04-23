#!/usr/bin/python
# -*- coding: utf-8 -*-
#from email import encoders
from email.header import Header
from email.mime.text import MIMEText
#from email.utils import parseaddr, formataddr
import smtplib
import time

acc_info = {'from': 'cloud-dev@qq.com', 'password': 'xxx', 'smtp': 'smtp.qq.com', 'ssl': True}

def err_log(msg):
    try:
        t = time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(time.time()))
        with open('errlog.txt', 'a+') as f:
            f.write((t + ': ' + msg+'\r\n').encode('utf-8'))
    except:
        pass

def smart_mail_send(acc, mailto, subject, content, ssl=True):
    ret = True
    from_addr = acc['from']
    password = acc['password']
    to_addr = mailto
    smtp_server = acc['smtp']
    
    try:
        msg = MIMEText(u'%s' %content, 'html', 'utf-8')
        msg['From'] = '%s' % from_addr
        msg['To'] = mailto
        msg['Subject'] = Header(u'%s' %subject, 'utf-8').encode()

        if ssl:
            server = smtplib.SMTP_SSL(smtp_server)
        else:
            server = smtplib.SMTP(smtp_server)
        #server.set_debuglevel(1)
        server.login(from_addr, password)
        server.sendmail(from_addr, [to_addr], msg.as_string())
    except Exception as e:
        str = u'向用户%s发送邮件失败' %mailto
        print str
        err_log(str)
        err_log(repr(e))
        ret = False
    if ret:
        try:
            server.quit()
        except Exception as e:
            print u'warning: server quit fail'
            err_log(repr(e))
    return ret

if __name__ == '__main__':
    smart_mail_send(acc_info, 'zhuyijing168@163.com', u'当月工资单', u'本月的工资为20000RMB')
