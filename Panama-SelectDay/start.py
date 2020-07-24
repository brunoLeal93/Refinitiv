
from GoogleAPI import getAttcEmailGmail
from CreateBulk import createBulk
from sendEmail import sendBulks
from datetime import datetime
import os
from Config import createDir, holidays

def main(date):
    date = datetime.strftime(datetime.strptime(date, '%d/%m/%Y'), '%Y-%m-%d')
    
    pathRoot = createDir(['DATA', date])
    pathFileXLSX =  getAttcEmailGmail(date, pathRoot)
    
    if type(pathFileXLSX) == list:
        for i in range(len(pathFileXLSX)):
            pathFileBulk = createBulk(date, pathRoot, pathFileXLSX[i])
            sendBulks(date, pathFileBulk)

    elif type(pathFileXLSX) == str:
        pathFileBulk = createBulk(date, pathRoot, pathFileXLSX)
        sendBulks(date, pathFileBulk)

    else:
        print('\nNo email to update.')

print('\n\t\t========== Panama ==========\n')
date = input('Enter a date (DD/MM/YYYY) to generate day bulks: ')
while True:
    if date not in holidays() and  datetime.strptime(date, '%d/%m/%Y').weekday() < 5:
        main(date)
        date = input('Enter with (x) to close app or enter another date (DD/MM/YYYY) to continue: ')
        if date == 'x':
            break
            
        else:
            continue
    else:
        print('Email no available to this day, try again!')
        date = input('Enter with (x) to close app or enter another date (DD/MM/YYYY) to continue: ')
        if date == 'x':
            break
            
        else:
            continue
