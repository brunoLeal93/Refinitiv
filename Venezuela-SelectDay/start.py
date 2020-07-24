import pandas as pd
from datetime import datetime
from ReadEmail import getNew
import os, time, sys
from caz  import createCAZ
from sendEmail import sendEmail
from report import createReport
# from pprint import pprint
from Config import holidays, sentmail
from CreateBulk import *

print('\n\t\t========== Venezuela ==========\n')
date = input('Enter a date (DD/MM/YYYY) to generate day bulks: ')


while True:
    if date not in holidays() and  datetime.strptime(date, '%d/%m/%Y').weekday() < 5:
        dt = datetime.strftime(datetime.strptime(date, '%d/%m/%Y'), '%d-%b-%y')

        # # Check if already exist the files locally, if there are, detele them to create new files
        steml = sentmail('r')
        fsent = [x[1] for x in steml]
        for fs in fsent:
            if os.path.exists(fs) and (dt in fs):
                os.remove(fs)

        # # GetNews(), fuction to take email for especific date 
        # # dt, Date definied before  
        data = getNew(dt)

        listAttachament= []
        loop = 1

        dfBonds = None
        dfEquities = None

        for d in data: 

            df = pd.DataFrame(d[0])

            dfBonds = df[df['typeInst'] == 'Bond']
            
            dfEquities = df[df['typeInst'] == 'Equity']

            if dfBonds.empty == False:
                dfBonds = createBulkGedaBonds(dfBonds, d[1])
                listAttachament.append(dfBonds[1])
                            
                
            if dfEquities.empty == False:
                dfEquities = createBulksEquities(dfEquities, d[1])
                listAttachament.append(dfEquities[1])
                listAttachament.append(dfEquities[2])
        
        # # Clean df with possible duplicates values and make Report and CAZ
        try:
            if dfBonds[0].empty == False:
                listAttachament.append(createReport(dfBonds[0], 'Bonds', dt))
                listAttachament.append(createCAZ(dfBonds[0], 'Bonds', dt))
        except:
            pass
        try: 
            if dfEquities[0].empty == False:
                listAttachament.append(createReport(dfEquities[0], 'Equities', dt))
                listAttachament.append(createCAZ(dfEquities[0], 'Equities', dt))
        except:
            pass
        
        listAttachament = list(dict.fromkeys(listAttachament))
        if listAttachament == []:
            print('Email no available to this day, try again!\n')
            date = input('\nEnter with (x) to close app or enter another date (DD/MM/YYYY) to continue: ')
            if date == 'x':
                break
            else:
                continue

        else:
            print('')
            x = [print(i) for i in listAttachament]
            se = sendEmail()
            se.sendFilesVezuela(dt, listAttachament)

            date = input('\nEnter with (x) to close app or enter another date (DD/MM/YYYY) to continue: ')
            if date == 'x':
                break
            else:
                continue
             
    else:
        print('\nThis day is Weekend or Holiday!\n')
        date = input('Enter with (x) to close app or enter another date (DD/MM/YYYY) to continue: ')
        if date == 'x':
            break
            
        else:
            continue