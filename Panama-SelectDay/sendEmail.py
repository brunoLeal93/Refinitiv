
# In addition to record final files locally,
# Application also send the results by email.

import mimetypes
import os, glob
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date, datetime, timedelta
import itertools
from Config import sentmail, recipients

rcpts = recipients()
TO = rcpts[0]
CC = rcpts[1]

def addAttachment( msg, filename):
    ctype, encoding = mimetypes.guess_type(filename)

    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'

    maintype, subtype = ctype.split('/', 1)

    if maintype == 'text':
        with open(filename) as f:
            mime = MIMEText(f.read(), _subtype=subtype)
        f.close()
    elif maintype == 'csv':
        with open(filename, 'rb') as f:
            mime = MIMEBase(f.read(), _subtype=subtype)
    else:
        with open(filename, 'rb') as f:
            mime = MIMEBase(maintype, subtype)
            mime.set_payload(f.read())

        encoders.encode_base64(mime)
    name = filename.replace('\\','/').split('/')
    length = len(name)
    mime.add_header('Content-Disposition', 'attachment', filename=name[length-1])
    msg.attach(mime)

def sendBulks( date, attachment):
    
    date = datetime.strftime(datetime.strptime(date, '%Y-%m-%d'), '%d-%b-%Y')
    sender = '<EMAIL GMAIL ACCOUNT>'
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = ', '.join(TO)
    msg['Cc'] = ', '.join(CC)
    msg['Subject'] = 'Update for Panama ' + date

    html= 'Update for Panama ' + date

    body = MIMEText(html, 'html')
    msg.attach(body)

    aux= []
    
    # Check on 'Sent Email' spreadsheet if files have already sent
    aux1 = sentmail('r')
    fsent = [x[1] for x in aux1]
    fsent = [x.split('/') for x in fsent]
    fsent = [x[len(x)-1] for x in fsent]

    attc = [x.split('/') for x in attachment]
    attc = [x[len(x)-1] for x in attc]

    for i in range(len(attachment)):

        if attc[i] not in fsent:
            addAttachment(msg, attachment[i]) 
            aux.append([str(datetime.now()), attachment[i]])

    #  Send if there are files have never sent
    if len(aux)>0:
        
        raw = msg.as_string()
        smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        smtp.login(sender, '<PASSWORD GMAIL ACCOUNT>')
        smtp.sendmail(sender, TO + CC, raw)
        smtp.quit()
        sentmail('w', aux1+aux)
        print('\nFiles Sent by email.\nYou also may access through "DATA" folder.')
    else:
        print('\nThe generated files have already been sent by email on the date informed.\nYou also may access through "DATA" folder.')

