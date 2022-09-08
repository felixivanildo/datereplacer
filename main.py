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

planilha = openpyxl.load_workbook("OBJETOTESTE.xlsx")
aba=planilha["OBJETO DE TESTE"]
feriados = holidays.Brazil()

df = pd.read_excel(
    io='OBJETOTESTE.xlsx',
    sheet_name= 'OBJETO DE TESTE',
    usecols='A:R'

)

quantidadelinhas=0


for row in aba.iter_rows(min_row=2, max_col=1):
    quantidadelinhas+=1

print(df.iloc[:, 0])
print(quantidadelinhas)


for i in range(quantidadelinhas):
    if datetime.strptime(str(df.iloc[i,0])[0:10],"%Y-%m-%d").weekday()==5:
        print('tem uma sexta aqui')
        df.replace(df[i][0], datetime.strptime(str(df.iloc[i,0]),"%Y-%m-%d %H:%M:%S") + timedelta(days=2))
    if datetime.strptime(str(df.iloc[i,0])[6:10],"%m-%d") in arrayferiado:
        print('gerado no feriado')


df.to_excel ("nemfodendo.xlsx", index=False)


