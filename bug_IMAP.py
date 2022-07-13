from datetime import datetime
import os
import imaplib
import email
from email.header import decode_header
import email.parser
from imap_tools.imap_utf7 import encode, decode
from email.header import decode_header
import re

Account = '' #輸入email account
PassWord = '' #輸入email password
Since_Date = datetime.now().strftime("%d-%b-%Y")

server = imaplib.IMAP4_SSL('imap.gmail.com') #輸入imap的伺服器
server.login(Account,PassWord)
server.select("INBOX")
typ, data = server.search(None, '(FROM "" SUBJECT "" SINCE "%s")' % (Since_Date)) #指定收取哪一個帳號傳的email

ids = data[0]
id_list = ids.split()

def decode_data(bytes, added_encode=None):
    def _decode(bytes, encoding):
        try:
            return str(bytes, encoding=encoding)
        except Exception as e:
            return None
    encodes = ['UTF-8', 'GBK', 'GB2312']
    if added_encode:
        encodes = [added_encode] + encodes
    for encoding in encodes:
        str_data = _decode(bytes, encoding)
        if str_data is not None:
            return str_data
    return None

def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,decode=True)

for i in id_list:
    typ, data = server.fetch(i, '(RFC822)')
    emailbody = data[0][1]
    mail = email.message_from_bytes(emailbody)
    mail_encode = decode_header(mail.get("Subject"))[0][1]
    mail_title = decode_data(decode_header(mail.get("Subject"))[0][0], mail_encode) #email主旨
    print(mail_title,"\n")
    mail_body = decode_data(get_body(mail)) #email內文
    pattern = re.compile('<[^>]+>')
    result = pattern.sub('',mail_body)
    result = result.replace("l&nbsp;", "")
    result = result.replace("&nbsp;", "")
    print(result)

    fileNames = []
    for part in mail.walk():        
        fileName = part.get_filename()
        pattern2 = re.compile('[^ \t\n\r\f\v]')
        if type(fileName) == str:
            fileName = pattern2.sub('', fileName)
        try:
            fileName = decode_header(fileName)[0][0].decode(decode_header(fileName)[0][1])
        except:
            pass
        if fileName:
            dirPath = os.path.join('下載的東西') #新增目錄
            if not os.path.exists(dirPath): #如果沒有就創一個
                os.makedirs(dirPath)
            filePath = os.path.join(dirPath, fileName) #放入要存取的目錄
            try:
                if not os.path.isfile(filePath):
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                    print("附件下載成功，文件為：" + fileName)
                else:
                    print("附件已存在，文件為：" + fileName)
            except Exception as e:
                print(e)