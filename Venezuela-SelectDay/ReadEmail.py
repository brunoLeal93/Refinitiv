
# It read and extract information from body emails.

from GoogleAPI import getEmail
from datetime import datetime, timedelta
import os
from Config import sentmail
from pprint import pprint

def getNew(dt):
    emails = getEmail(dt)

    emissor = []
    ISIN = []
    SYMBOL = []
    FechaInicioSesion = []
    FechaVencimientoSerie = []
    typeInst = []
    result= []

    for eml in emails:

        for i in range(len(eml['msg'])-1):
            # print(i, eml['msg'][i])
            if 'Mercado de Valores de Renta Fija' in eml['msg'][i]:
                typeInst.append('Bond')
            
            elif 'Mercado de Valores de Renta Variable' in eml['msg'][i]:
                typeInst.append('Equity')
            
            elif eml['msg'][i].strip() == 'Emisor':
                emissor.append(eml['msg'][i+2])

            elif eml['msg'][i].strip() == 'Códigos  ISIN':
                x = i
                while (eml['msg'][x].strip() != 'Código del Activo') and (eml['msg'][x].strip() != 'Código de Símbolo en SIBE, Mercado Primario'):
                    x+=1

                y = i
                while y+1 != x:
                    if len(eml['msg'][y+1]) > 12:
                        ISIN.append(eml['msg'][y+1][-12:])
                    else:
                        if eml['msg'][y+1] != '':
                            ISIN.append(eml['msg'][y+1])
                    y+=1

            elif eml['msg'][i].strip() == 'Código de Símbolo en SIBE, Mercado Primario':
                x = i
                while eml['msg'][x].strip() != 'Fecha de Inicio de Sesión Especial':
                    x+=1

                y = i
                while y+1 != x:
                    if len(eml['msg'][y+1]) > 6:
                        SYMBOL.append(eml['msg'][y+1][-6:])
                    else:
                        if eml['msg'][y+1] != '':
                            SYMBOL.append(eml['msg'][y+1])
                    y+=1

            elif eml['msg'][i].strip() == 'Fecha de Inicio de Sesión Especial':
                x = i
                while (eml['msg'][x].strip().upper() != 'FECHA DE VENCIMIENTO DE LA SERIE') and (eml['msg'][x].strip().upper() != 'FECHA DE VENCIMIENTO DE LA SERIE.') and (eml['msg'][x].strip().upper() != 'FECHA DE VENCIMIENTO DE LAS SERIES') and (eml['msg'][x].strip().upper() != 'FECHA DE VENCIMIENTO DE LAS SERIES.'):
                    x+=1
            
                y = i
                while y+1 != x:
                    try:
                        aux = datetime.strptime(eml['msg'][y+1], '%d/%m/%Y')
                        FechaInicioSesion.append(eml['msg'][y+1])
                        y+=1
                    except:
                        y+=1

            elif (eml['msg'][i].strip().upper() == 'FECHA DE VENCIMIENTO DE LA SERIE') or (eml['msg'][i].strip().upper() == 'FECHA DE VENCIMIENTO DE LA SERIE.') or (eml['msg'][i].strip().upper() == 'FECHA DE VENCIMIENTO DE LAS SERIES') or (eml['msg'][i].strip().upper() == 'FECHA DE VENCIMIENTO DE LAS SERIES.'):
                x = i
                while eml['msg'][x].strip() != 'Agentes de la Colocación':
                    x+=1

                y = i
                while y+1 != x:
                    try:
                        aux = datetime.strptime(eml['msg'][y+1], '%d/%m/%Y')
                        FechaVencimientoSerie.append(eml['msg'][y+1])
                        y+=1
                    except:
                        y+=1
                
                if len(FechaVencimientoSerie) == 0:
                    for dt in FechaInicioSesion:
                        date = datetime.strftime(datetime.strptime(dt, '%d/%m/%Y')+timedelta(366), '%d/%m/%Y')
                        FechaVencimientoSerie.append(date)

        if len(ISIN) != len(SYMBOL):
            SYMBOL.pop()
        
        for i in range(len(ISIN)-1):
            emissor.append(emissor[0])
            FechaInicioSesion.append(FechaInicioSesion[0])
            FechaVencimientoSerie.append(FechaVencimientoSerie[0])
            typeInst.append(typeInst[0])

        # print('emissor', len(emissor))
        # print('ISIN', len(ISIN))
        # print('SYMBOL', len(SYMBOL))
        # print('FechaInicioSesion', len(FechaInicioSesion))
        # print('FechaVencimientoSerie', len(FechaVencimientoSerie))
        # print('typeInst', str(len(typeInst))+'\n', typeInst)

        data = {
            'emissor': emissor,
            'ISIN': ISIN,
            'SYMBOL': SYMBOL,
            'FechaInicioSesion': FechaInicioSesion,
            'FechaVencimientoSerie': FechaVencimientoSerie,
            'typeInst': typeInst
        }
        
        result.append([data, eml['date']])
        ISIN = []
        emissor = []
        SYMBOL = []
        FechaInicioSesion = []
        FechaVencimientoSerie = []
        typeInst = []

    return result