from datetime import datetime
from datetime import timedelta
import openpyxl
import pandas as pd
import openpyxl as px
from pandas.tseries import holiday
import holidays
from pandas.tseries import holiday


arrayferiado = [
    '01-01',
    '02-02',
    '02-28',
    '03-01',
    '03-02',
    '04-14',
    '04-15',
    '04-21',
    '05-01',
    '06-16',
    '06-24',
    '06-29',
    '09-07',
    '10-12',
    '10-28',
    '10-30',
    '11-02',
    '11-15',
    '11-20',
    '11-30',
    '12-08',
    '12-24',
    '12-25',
    '12-31'


]

planilha = openpyxl.load_workbook("RELATORIO MÊS 07.xlsx")
aba=planilha["RELATORIO MÊS 07"]
feriados = holidays.Brazil()

df = pd.read_excel(
    io='RELATORIO MÊS 07.xlsx',
    sheet_name= 'RELATORIO MÊS 07',
    usecols='A:R'

)

quantidadelinhas=0


for row in aba.iter_rows(min_row=2, max_col=1):
    quantidadelinhas+=1

#print(df.iloc[:, 0])
#print(quantidadelinhas)

i=0
for i in range(quantidadelinhas):
    # print(df['DATA GERACAO'][i].weekday())
    hora=df['DATA GERACAO'][i].hour
    #print(str(df['DATA GERACAO'][i])[0:10])
    if df['DATA GERACAO'][i].weekday() ==3 and hora <= 12 and df['SERVICO TIPO'][i]== 'CORTE POR DEBITO':
        data = df['DATA GERACAO'][i] + timedelta(days=3)
        df['DATA GERACAO'][i] = data

    if df['DATA GERACAO'][i].weekday() == 4:
        data=df['DATA GERACAO'][i]+timedelta(days=3)
        df['DATA GERACAO'][i]=data

    if str(df['DATA GERACAO'][i])[0:10] in arrayferiado:
        data = df['DATA GERACAO'][i] + timedelta(days=1)
        print('gerado no feriado')

    if df['SERVICO TIPO'][i]=='VISITA DE COBRANCA':
        data = df['DATA GERACAO'][i] + timedelta(days=1)
        df['DATA GERACAO'][i] = data







df.to_excel ("DATAS RETIFICADAS MÊS 07.xlsx", index=False)

