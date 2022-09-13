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



df = pd.read_excel(
    io='RELATORIO MÊS 07.xlsx',
    sheet_name= 'RELATORIO MÊS 07',
    usecols='A:R'

)

print((df['DATA GERACAO'][1]).hour)