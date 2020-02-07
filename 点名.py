#-*- encoding: utf-8 -*-

#主模块

#参考
#1. Python 用IMAP接收邮件 https://www.cnblogs.com/zixuan-zhang/p/3402825.html
#2. Python学习：通过IMAP收邮件 https://blog.csdn.net/pengzhi5966885/article/details/75019099

import configs

import os
import sys
import pandas
from imapclient import IMAPClient
import email

#读取本班同学信息
classmateExcelData=pandas.read_excel(configs.roster)
classmates=classmateExcelData.get('姓名').tolist()
#放到字典中（在为假，不在为真）
classmateCheck=set(classmates)
#for i in range(len(classmates)):
#    classmateCheck.add(classmates[i])
#收邮件
subjects=[]
print('正在连接……')
# 通过以下方式连接smtp服务器，没有考虑异常情况，详细请参考官方文档
c = IMAPClient(host = configs.imapServer, ssl= True) 
try:
    c.login(configs.username, configs.passwd) #登录个人帐号
except c.Error as err:
    print('Could not log in')
    sys.exit(1)
else:
    #c.select_folder('INBOX', readonly = True) 
    c.select_folder('INBOX',readonly=False)
    result = c.search('UNSEEN')
    print('有%d封未读邮件。'%len(result))
    #msgdict = c.fetch(result, ['BODY.PEEK[]'] )
    msgdict=c.fetch(result,['BODY[]'])
    msglist=msgdict.items()
    for message_id, message in msglist:
        #print(type(message))
        listMsg=message.items()
        strMsg=None
        for p in listMsg:
            if p[0]==b'BODY[]':
                strMsg=p[1]
        #e = email.message_from_string(message[b'BODY[]']) # 生成Message类型
        #e=email.message_from_bytes(message.items()[0])
        e=email.message_from_bytes(strMsg)
        # 由于'From', 'Subject' header有可能有中文，必须把它转化为中文
        subject = email.header.make_header(email.header.decode_header(e['SUBJECT']))
        subjects.append(subject.__str__())
finally:
    c.logout()
print('已读取邮件，正在解析……')
#解析邮件标题
for s in subjects:
    strArray=s.split(maxsplit=1)
    if len(strArray)>1 and configs.mark in strArray[0] and strArray[1] in classmateCheck:
        classmateCheck.remove(strArray[1])
#显示结果
listClassmateCheck=list(classmateCheck)
print('\n以下%d名同学未报数：'%len(listClassmateCheck))
for classmate in listClassmateCheck:
    print(classmate)
print()
