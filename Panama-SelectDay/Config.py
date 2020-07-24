
# Taking advantage of the Google API and the need to have a Front End 
# to control some functions of the application, the use of Googles Sheets was implemented 
# as a simple way to achieve the solution given the current situation.

import gspread, os
from oauth2client.service_account import ServiceAccountCredentials

localDir= os.getcwd().replace('\\','/')

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(localDir+ '/service_account.json',scope)
gc = gspread.authorize(credentials)

ss = gc.open("Panama")
ws_email = ss.worksheet('Recipients')
ws_holid = ss.worksheet('Holidays')
ws_sentmail = ss.worksheet('Sent Email')


def createDir(folder):
    path = ''
    for i in range(len(folder)):

        path = path + '/' + folder[i]

        if not os.path.exists(localDir + path):
            os.makedirs(localDir + path)

    return localDir + path

def recipients():
    To = []
    Cc = []
    dic = ws_email.get_all_records()

    for d in dic:
        if d['To'] != '':
            To.append(d['To'])

    for d in dic:
        if d['Cc'] != '':
            Cc.append(d['Cc'])
    
    return [To, Cc]

def holidays():
    holidays = []

    dic = ws_holid.get_all_records()
    
    for d in dic:
        if d['Date'] != '':
            holidays.append(d['Date'])
    
    return holidays

def sentmail(func, content=None):
    
    if func == 'r':
        return ws_sentmail.get_all_values()

    elif func == 'w' and content != None:
        ws_sentmail.update(content)



# print(holidays())