# It receive DataFrames then apply due rules to make 
# the final files and record at DATA folder.

from Config import createDir
import pandas as pd
from datetime import datetime
import os, time, sys

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
