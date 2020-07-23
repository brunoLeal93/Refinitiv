import pandas as pd
from datetime import datetime
from ReadEmail import getNew
import os, time, sys
from caz  import createCAZ
from sendEmail import sendEmail
from report import createReport
from pprint import pprint
from Config import createDir, holidays, sentmail

def transformDsplyNm(nm):
    nm = nm.upper().replace('.', ' ').replace('-', ' ').replace('/', ' ').replace(' S A', '').replace('Á','A').replace('Â','A').replace('À','A').replace('Ã','A').replace('Ä','A').replace('É','E').replace('Ê','E').replace('Ë','E').replace('È','E').replace('Í','I').replace('Î','I').replace('Ì','I').replace('Ï','I').replace('Ó','O').replace('Ô','O').replace('Ò','O').replace('Ø','O').replace('Ú','U').replace('Û','U').replace('Ù','U').replace('Ü','U').replace('Ç','C').replace('Ñ','N').replace('Ý','Y').strip()

    if len(nm)<=16:
        return nm
    else:
        return nm[:16]

def transformOrgNm(nm):
    nm = nm.upper().replace('.', ' ').replace('-', ' ').replace('/', ' ').replace(' S A', '').replace('Á','A').replace('Â','A').replace('À','A').replace('Ã','A').replace('Ä','A').replace('É','E').replace('Ê','E').replace('Ë','E').replace('È','E').replace('Í','I').replace('Î','I').replace('Ì','I').replace('Ï','I').replace('Ó','O').replace('Ô','O').replace('Ò','O').replace('Ø','O').replace('Ú','U').replace('Û','U').replace('Ù','U').replace('Ü','U').replace('Ç','C').replace('Ñ','N').replace('Ý','Y').strip()

    if len(nm)+4<=36:
        return nm
    else:
        return nm[:32]

def createBulkGedaBonds(df, dt):

    localSaveFile = createDir(['DATA', dt, 'Bonds', 'CreateGeda'])

    try:
        dfAux = pd.read_csv(localSaveFile+'/ADD_Bonds_GEDA_'+dt+'.txt', sep='\t')
    except:
        pass
        # print('Failed')

    dfFinal = pd.DataFrame()
    dfFinal.loc[:,'RIC'] = df['SYMBOL'].map(lambda x: x+'=CR')
    dfFinal.loc[:,'SYMBOL'] = df['SYMBOL']
    dfFinal.loc[:,'DSPLY_NAME'] = df['SYMBOL']
    dfFinal.loc[:,'OFFCL_CODE'] = df['ISIN']
    dfFinal.loc[:,'COUPN_RATE'] = '0'
    dfFinal.loc[:,'ISSUE_DATE'] = df['FechaInicioSesion'].map(lambda x: datetime.strftime(datetime.strptime(x, '%d/%m/%Y'), '%d-%m-%Y'))
    dfFinal.loc[:,'ORG_NAME'] = df['emissor'].map(lambda x: x.upper())
    dfFinal.loc[:,'MATUR_DATE'] = df['FechaVencimientoSerie'].map(lambda x: datetime.strftime(datetime.strptime(x, '%d/%m/%Y'), '%d-%m-%Y'))
    dfFinal.loc[:,'NAME_ROOT'] = df['SYMBOL'].map(lambda x: x+'@')
    dfFinal.loc[:,'ISSUE_DETAILS'] = df['ISIN']
    dfFinal.loc[:,'#INSTMOD_TDN_SYMBOL'] = df['SYMBOL']
    dfFinal.loc[:,'EXL_NAME'] = 'CCSER_CCSEBOND'
    dfFinal.loc[:,'CHAIN_RIC'] = 'BVCBONDS.CR'

    try:
        dfFinal = pd.concat([dfFinal, dfAux])
    except:
        pass
        # print('Failed')
    dfFinal.drop_duplicates(inplace=True)
    dfFinal.to_csv(localSaveFile+'/ADD_Bonds_GEDA_'+dt+'.txt',sep='\t', index=False, encoding='utf-8-sig')
    #dfFinal.to_csv(localSaveFile+'ADD_Bonds_GEDA_'+dt+'.csv',sep=',', index=False)
    return [dfFinal, localSaveFile+'/ADD_Bonds_GEDA_'+dt+'.txt']

def createBulksEquities(df, dt):

    localSaveFile = createDir(['DATA', dt, 'Equities', 'CreateGeda'])

    try:
        dfAux = pd.read_csv(localSaveFile+'/ADD_Equities_GEDA_'+dt+'.txt', sep='\t')
    except:
        pass

    dfFinal = pd.DataFrame()

    dfFinal.loc[:,'RIC'] = df['SYMBOL'].map(lambda x: x+'.CR' if '.' not in x else x[:3]+x[-1:].lower()+'.CR')
    dfFinal.loc[:,'SYMBOL'] = df['SYMBOL']
    dfFinal.loc[:,'DSPLY_NAME'] = df['emissor'].map(lambda x: transformDsplyNm(x))
    dfFinal.loc[:,'OFFCL_CODE'] = df['ISIN']
    dfFinal.loc[:,'DDS_SYMBOL'] = df['SYMBOL']
    dfFinal.loc[:,'NAME_ROOT'] = df['emissor'].map(lambda x: x.upper()+'@') # Necessary correct manually
    dfFinal.loc[:,'ISSUE_DETAILS'] = df['SYMBOL'].map(lambda x: 'STK' if '.' not in x else 'STK '+x[-1:])
    dfFinal.loc[:,'ORG_NAME'] = df['emissor'].map(lambda x: transformOrgNm(x)+' ORD')
    dfFinal.loc[:,'#INSTMOD_ISIN_CODE'] = df['ISIN']
    dfFinal.loc[:,'#INSTMOD_TDN_ASSET_SUB_TYPE'] = '#NULL#'
    dfFinal.loc[:,'#INSTMOD_TDN_CURRENCY'] = 'VES'
    dfFinal.loc[:,'#INSTMOD_TDN_ISSUE_DESC'] = 'ORD'
    dfFinal.loc[:,'#INSTMOD_TDN_SYMBOL'] = df['SYMBOL']
    dfFinal.loc[:,'EXL_NAME'] = 'CCSER_CCSEEQUITY'

    try:
        dfFinal = pd.concat([dfFinal, dfAux])
    except:
        pass
    dfFinal.drop_duplicates(inplace=True)

    pathBulkNDA = createBulkNdaEquities(dfFinal,dt)

    dfFinal.to_csv(localSaveFile+'/ADD_Equities_GEDA_'+dt+'.txt',sep='\t', index=False, encoding='utf-8-sig')
    return [dfFinal, localSaveFile+'/ADD_Equities_GEDA_'+dt+'.txt', pathBulkNDA]

def createBulkNdaEquities(df, dt):
    localSaveFile = createDir(['DATA', dt, 'Equities', 'CreateNda'])
    try:
        dfAux = pd.read_csv(localSaveFile+'/ADD_Equities_GEDA_'+dt+'.txt', sep='\t')
    except:
        pass
    dfFinal = pd.DataFrame()

    dffinal.loc[:,'RIC'] = df['RIC']
    dffinal.loc[:,'ASSET SHORT NAME'] =  df['DSPLY_NAME']
    dffinal.loc[:,'ASSET COMMON NAME'] = df['DSPLY_NAME']
    dffinal.loc[:,'TICKER SYMBOL'] = df['SYMBOL']
    dffinal.loc[:,'TAG'] = '199'
    dffinal.loc[:,'TYPE'] = 'EQUITY'
    dffinal.loc[:,'CATEGORY'] = 'ORD'
    dffinal.loc[:,'SETTLEMENT PERIOD'] = 'T+2'
    dffinal.loc[:,'CURRENCY'] = 'VES'
    dffinal.loc[:,'EXCHANGE'] = 'CCS'
    dffinal.loc[:,'ROUND LOT SIZE'] = '1'
    #dffinal.loc[:,'PILC'] =
    #dffinal.loc[:,'ISIN'] =

    try:
        dfFinal = pd.concat([dfFinal, dfAux])
    except:
        pass
    dfFinal.drop_duplicates(inplace=True)

    dfFinal = pd.DataFrame()

    dfFinal.to_csv(localSaveFile+'/ADD_Equities_NDA_'+dt+'.csv',sep=',', index=False, encoding='utf-8-sig')

    return localSaveFile+'/ADD_Equities_NDA_'+dt+'.csv'

# ===============================================================

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