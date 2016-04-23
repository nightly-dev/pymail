#!/usr/bin/python
# -*- coding: utf-8 -*-
import xmail
import cfgreader
import excelreader
import mylog
import time
import os
import copy
import webbrowser

print u'''
****************************************
*                                      *
*         欢迎使用本程序！             *
*   邮件未发完之前请不要关闭本窗口     *
*                           author@nick*
****************************************

'''

templates = '''
    <!DOCTYPE html>
    <html>
    <head>
    </head>
    <body>

    <table border="1" style="border-collapse: collapse;border-width: 1px;padding: 8px;">
      ###
    </table>
    </body>
    </html>

'''

##########helper functions###########################
def prog_exit():
    os.system('pause')
    quit()

def fstr(value):
    str = '%s' % value
    if str.endswith('+'):
        str = str.rstrip('+')
    if str.endswith('-'):
        str = str.rstrip('-')
    if '$' in str:
        str = str.replace('$', '')
    if str.endswith('.0'):
        str = str[:-2]
    return str

maxstyle = 100
colors = ['black']*maxstyle
bold = [False]*maxstyle

def format_table(title_row, value_row):
    tbl = ''
    slip = 1
    max_per_line = 5
    count = len(title_row) - slip
    if max_per_line > count:
        max_per_line = count
    for i in range(slip):
        title_row.pop()    #pop email address
        value_row.pop()    #pop email address
    x = count/max_per_line
    if count % max_per_line != 0:
        y = max_per_line - (count % max_per_line)
    else:
        y = 0
    for i in range(y):
        title_row.append(' ')   #align title with multiple of max_per_line
        value_row.append(' ')   #align value with multiple of max_per_line
    count = len(title_row)  #count after align, should be multiple of max_per_line already

    for i in range(0, count, max_per_line):
        line = ''
        for j in range(max_per_line):
            line = line + '<td>%s</td>' % fstr(title_row[i+j])
            if title_row[i+j].endswith('+'):
                colors[i+j] = 'blue'
            elif title_row[i+j].endswith('-'):
                colors[i+j] = 'red'
            if '$' in title_row[i+j]:
                bold[i+j] = True
        tbl = tbl + '<tr bgcolor="LightSkyBlue">%s</tr>' % line
        line = ''
        for j in range(max_per_line):
            if not bold[i+j]:
                line = line + '<td><font color="%s">%s</font></td>' % (colors[i+j], fstr(value_row[i+j]))
            else:
                line = line + '<td><font color="%s"><b>%s</b></font></td>' % (colors[i+j], fstr(value_row[i+j]))
        tbl = tbl + '<tr>%s</tr>' % line
    return tbl
            
##########helper functions###########################


##########global varibles############################
accs = None
interval = 2        #default interval between mail sends

success = 0         #number of mail sends successfully
skip = 0            #number of no mail address supplied
failed = 0          #number of mail send failure
##########global varibles############################

#load and check configurations
data = cfgreader.load_cfg()
if data:
    accs = data['accounts']
    interval = data['interval']
else:
    print u'配置文件不正确，请检查mycfg.json'
    prog_exit()

#employee info is begin from offset, define in mycfg.json, must be equal or greater than 1
offset = data['offset']

#validate the excel file
if not excelreader.check_xls(offset):
    prog_exit()

#load data from excel files
nrows, ncols, row_values, filename = excelreader.read_xls()
if not nrows:
    print u'读取excel文件失败'
    prog_exit()
else:
    print u'此excel文件共有%d行%d列!' % (nrows, ncols)

row_title = row_values(offset-1)

subject = filename.decode('cp936')
subject = subject.strip('.xlsx').strip('.xls')

#how many mail accounts are used as sender, defined in mycfg.json, load balance purpose
acc_num = len(accs)

#clear mylog.html
mylog.clear_log()

for i in range(offset, nrows):
    row_value = row_values(i)
    tblstr = format_table(copy.deepcopy(row_title), copy.deepcopy(row_value))
    content = templates.replace('###', tblstr)

    #the last column is mail address 
    mailto = row_values(i)[-1]
    #the first column is user name
    username = row_values(i)[0]
    #the second colum is user id
    userid = fstr(row_values(i)[1])

    acc = accs[(i-offset)%acc_num]
    ssl_enabled = acc['ssl']
    
    if '@' in mailto:
        ret = xmail.smart_mail_send(acc, mailto, subject, content, ssl=ssl_enabled)
        if ret:
            str = u'(%d/%d) 已成功发送至用户 %s %s' %(i, nrows-offset, username, userid)
            success = success + 1
            color = 'green'
        else:
            str = u'(%d/%d) 向用户 %s %s发送邮件失败，原因未知' %(i, nrows-offset, username, userid)
            failed = failed + 1
            color = 'red'
        print str
    else:
        str = u'(%d/%d) 请提供用户 %s %s的合法邮箱' %(i, nrows-offset, username, userid)
        print str
        skip = skip + 1
        color = 'blue'
    mylog.write_log('<font color=%s>' %color +str+ '</font>')
    time.sleep(interval)

#conclusion message
mylog.write_log(u'<br><hr /><font size=5>成功发送%d封邮件，%d位用户未提供邮件, %d位用户发送失败<font>' %(success, skip, failed))

#open log in default browser
logfile = os.path.join(os.getcwd(), 'mylog.html')
webbrowser.open(logfile)

os.system('pause')