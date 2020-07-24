
# It Receive downloaded file then apply due rules to make 
# the final files and record at DATA folder.

import pandas as pd
import xlrd
pd.options.mode.chained_assignment = None
import os
from datetime import datetime

pathRoot = os.getcwd().replace('\\','/')

def bulkCorp(date, pathRoot, df):
    bulk = pd.DataFrame()

    bulk['SYMBOL'] = df['RIC']
    bulk['DSPLY_NAME'] = df['INSTRUMENTO\n']
    bulk['RIC'] = df['RIC']
    bulk['OFFCL_CODE'] = df['ISIN\n']
    bulk['EX_SYMBOL'] = bulk['DSPLY_NAME']
    bulk['MATUR_DATE'] = df['FECHA DE VENCIMIENTO\n'].map(lambda x: datetime.strftime(datetime.strptime(str(x),'%d/%m/%Y'), '%d-%b-%Y').upper())
    bulk['COUPN_RATE'] = df['TASA %\n']
    bulk['ISSUE_DATE'] = df['FECHA DEL CIERRE DEL PRECIO ANTERIOR\n'].map(lambda x: datetime.strftime(datetime.strptime(str(x), '%Y-%m-%d %H:%M:%S'), '%d-%b-%Y').upper())
    bulk['BOND_TYPE'] = '#IGNORE#'
    bulk['EUROCLR_NO'] = '0'
    bulk['CEDEL_NO'] = '0'
    bulk['ISSUE_PRC'] = '0'
    #bulk['CALL_DATE'] = 
    #bulk['RATING_ID'] = 
    #bulk['BCKGRNDPAG'] = 
    #bulk['OFFC_CODE2'] = 
    #bulk['TDN_SYMBOL'] = 
    #bulk['DDS_SYMBOL'] = 
    #bulk['COMMNT'] = 
    bulk['EXL_NAME'] = df['EXL']

    bulk.to_csv(pathRoot + '/PANA_CORP_'+date+'.txt', sep='\t', index=False)

    return pathRoot + '/PANA_CORP_'+date+'.txt'

def bulkStock(date, pathRoot, df):
    bulk = pd.DataFrame()

    bulk['SYMBOL'] = df['RIC']
    bulk['DSPLY_NAME'] = df['EMISOR\n']
    bulk['RIC'] = df['RIC']
    bulk['OFFCL_CODE'] = df['ISIN\n']
    bulk['EX_SYMBOL'] = df['EMISOR\n']
    #bulk['BCKGRNDPAG'] =
    #bulk['TDN_SYMBOL'] =
    #bulk['DDS_SYMBOL'] =
    #bulk['COMMNT'] =
    bulk['EXL_NAME'] = df['EXL']

    bulk.to_csv(pathRoot + '/PANA_STOCK_'+date+'.txt', sep='\t', index=False)

    return pathRoot + '/PANA_STOCK_'+date+'.txt'

def bulkDebt(date, pathRoot, df):
    bulk = pd.DataFrame()
    
    bulk['SYMBOL'] = df['RIC']
    bulk['DSPLY_NAME'] = df['INSTRUMENTO\n']
    bulk['RIC'] = df['RIC']
    bulk['OFFCL_CODE'] = df['ISIN\n']
    bulk['EX_SYMBOL'] = bulk['DSPLY_NAME']
    bulk['MATUR_DATE'] = df['FECHA DE VENCIMIENTO\n'].map(lambda x: datetime.strftime(datetime.strptime(str(x),'%d/%m/%Y'), '%d-%b-%Y').upper())
    bulk['COUPN_RATE'] = df['TASA %\n']
    bulk['ISSUE_DATE'] = df['FECHA DEL CIERRE DEL PRECIO ANTERIOR\n'].map(lambda x: datetime.strftime(datetime.strptime(str(x),'%Y-%m-%d %H:%M:%S'), '%d-%b-%Y').upper())
    bulk['BOND_TYPE'] = '#IGNORE#'
    bulk['EUROCLR_NO'] = '0'
    bulk['CEDEL_NO'] = '0'
    bulk['ISSUE_PRC'] = '0'
    #bulk['CALL_DATE'] = 
    #bulk['RATING_ID'] = 
    #bulk['BCKGRNDPAG'] = 
    #bulk['OFFC_CODE2'] = 
    #bulk['TDN_SYMBOL'] = 
    #bulk['DDS_SYMBOL'] = 
    #bulk['COMMNT'] = 
    bulk['EXL_NAME'] = df['EXL']

    bulk.to_csv(pathRoot + '/PANA_DEBT_'+date+'.txt', sep='\t', index=False)

    return pathRoot + '/PANA_DEBT_'+date+'.txt'

def createBulk(date, pathRoot, pathFileXLS):
    wb = xlrd.open_workbook(pathRoot+ '/preciocierresTR.xls', logfile=open(os.devnull, 'w'))
    df = pd.read_excel(wb, engine='xlrd', sheet_name='preciocierresTR')
    # df = pd.read_excel(pathRoot+ '/preciocierresTR.xls', sheet_name='preciocierresTR')

    df.dropna(1,how='all', inplace=True)
    df.dropna(0,how='all', inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.drop([0], inplace=True)
    df.reset_index(drop=True, inplace=True)
    # print(df.columns)
    pathBulks = []
    debt = pd.DataFrame()
    debt = df[df['RUEDA QUE MARCA PRECIO DE CIERRE\n'] == 'GOB_MM_DEUDA']
    debt['RIC'] = debt['ISIN\n'].map(lambda x: x+'=PN')
    debt['EXL'] = 'PAN_DEBT'
    debt.reset_index(drop=True, inplace=True)
    pathBulks.append(bulkDebt(date, pathRoot, debt))
    print('\nDebentures\n\n',(debt if debt.empty == False else "- No update\n"))

    corp = pd.DataFrame()
    corp = df[(df['RUEDA QUE MARCA PRECIO DE CIERRE\n'] == 'SEC_DEUDA') | (df['RUEDA QUE MARCA PRECIO DE CIERRE\n'] == 'PRIM_DEUDA')]
    corp['RIC'] = corp['ISIN\n'].map(lambda x: x+'=PN')
    corp['EXL'] = 'PAN_CORP'
    corp.reset_index(drop=True, inplace=True)
    pathBulks.append(bulkCorp(date, pathRoot, corp))
    print('\nCorporate:\n\n',(corp if corp.empty == False else "- No update\n"))

    stock = pd.DataFrame()
    stock = df[(df['RUEDA QUE MARCA PRECIO DE CIERRE\n'] != 'SEC_DEUDA') & (df['RUEDA QUE MARCA PRECIO DE CIERRE\n'] != 'PRIM_DEUDA') & (df['RUEDA QUE MARCA PRECIO DE CIERRE\n'] != 'GOB_MM_DEUDA')]
    stock['RIC'] = stock['INSTRUMENTO\n'].map(lambda x: x+'.PN')
    stock['EXL'] = 'PAN_EQUITY'
    stock.reset_index(drop=True, inplace=True)
    pathBulks.append(bulkStock(date, pathRoot, stock))
    print('\nStock:\n\n',(stock if stock.empty == False else "- No update\n"))
    
    return pathBulks