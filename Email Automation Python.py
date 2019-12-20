import imaplib, email, os
import xlrd

import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders

user = 'cs.chshekhar13@gmail.com'
password = 'password'
imap_url = 'imap.gmail.com'
attachment_dir = 'C:/Users/hp/Desktop/attachments/'

def auth(user,password,imap_url):
    con = imaplib.IMAP4_SSL(imap_url)
    con.login(user,password)
    return con

def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)

def get_attachments(msg):
    for part in msg.walk():
        if part.get_content_maintype()=='multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()

        if bool(fileName):
            filePath = os.path.join(attachment_dir, fileName)
            with open(filePath,'wb') as f:
                f.write(part.get_payload(decode=True))
        return fileName

def search(key,value,con):
    result, data  = con.search(None,key,'"{}"'.format(value))
    return data

def get_emails(result_bytes):
    msgs = []
    for num in result_bytes[0].split():
        typ, data = con.fetch(num, '(RFC822)')
        msgs.append(data)
    return msgs

con = auth(user,password,imap_url)
con.select('INBOX')

result, data = con.fetch(b'59','(RFC822)')
raw = email.message_from_bytes(data[0][1])
fileName = get_attachments(raw)

#search('FROM','cs.chshekhar13@gmail.com',con)
#msgs = get_emails((search('FROM','cs.chshekhar13@gmail.com',con)))
#for msg in msgs:
#   print(get_body(email.message_from_bytes(msg[0][1])))    

  
loc = ("C:/Users/hp/Desktop/attachments/"+fileName) 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 

data1 = []
for i in range(sheet.nrows): 
    data1.append( sheet.cell_value(i, 0))

#print(data1)


with open('C:/Users/hp/Desktop/attachments/'+fileName.split('.')[0]+'.txt', 'w+') as f:
    for item in data1:
        f.write("%s\n" % item)
 
   
fromaddr = "cs.chshekhar13@gmail.com"
toaddr = "deepan.sinha@capgemini.com"
   

msg = MIMEMultipart() 
  
msg['From'] = fromaddr 
  
msg['To'] = toaddr 

msg['Subject'] = "Python Assignment: Email (with attachment) Automation"

body = "Hello Deepan,\n\nThis mail is the Python assignment I was given. This is a python-script generated email.\n\nRegards,\nChandra Shekhar\nchandra.c.shekhar@capgemini.com"

msg.attach(MIMEText(body, 'plain')) 

filename = "Names.txt"
attachment = open("C:/Users/hp/Desktop/attachments/" + filename, "rb") 

p = MIMEBase('application', 'octet-stream') 

p.set_payload((attachment).read()) 

encoders.encode_base64(p) 
   
p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 

msg.attach(p)

s = smtplib.SMTP('smtp.gmail.com', 587) 

s.starttls() 

s.login(fromaddr, "password") 

text = msg.as_string() 

s.sendmail(fromaddr, toaddr, text) 

s.quit()
