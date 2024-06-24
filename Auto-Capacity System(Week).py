from pathlib import Path 
import pandas as pd
from openpyxl import load_workbook
import openpyxl as xl
import pyodbc 
import pandas as pd
import csv
import xlrd
import os
import re
import time
import datetime

#讀取CAPA_Report資料夾最新檔, 檔名前面為'Manpower(Weekly)_Based_on_MPS_*.xls', *為模糊搜尋日期(如20231122), 且資料夾中沒有_TEST字串的檔案
p=Path('N:\DataExchange\Production_Control\CAPA Report')
a=max([fn for fn in p.glob('Manpower(Weekly)_Based_on_MPS_*.xls') if not '_Test' in str(fn)], key=lambda f: f.stat().st_mtime)

#把檔名轉成字串
a1 = str(a)
print (a1)
#讀取檔名中的日期部分, 並且程式執行會於終端機顯示
Date1 = re.search('on_MPS_(.*?).xl', a1).group(1)
print ('FileDate:', Date1)
print('Data Uploading...')


data = pd.read_excel (a1, keep_default_na = False)  
#刪除用不到的column 
df = pd.DataFrame(data)
del df["Cat1"]; del df["Cat2"]; del df["Cat3"]; del df["Reference\nRsc"]; del df["Rsc Type"]; del df["Map Flag"]; del df["UPH"]; del df["Usage"]; del df["Reference\nItem"]; del df["Reference\nLevel"]
del df["ManHrs pcs\nERP Actual"]; del df["Main Process"]
#飾選cell中為MPS Qty出來
df = df[(df['Type'] == 'MPS Qty')]
#重新命名部分column名稱
df.rename(columns={'Item Desc':'ItemDesc','OSP Flag':'OSP_Flag','ManHrs pcs\nERP STD':'ManHrspcsERPSTD', 'ManHrs pcs\nMAP STD':'ManHrspcsMAPSTD'}, inplace = True)


#行列轉換: 如id_vars皆為函式功能
df = df.melt(id_vars=['Org','Item','ItemDesc','OSP_Flag','ManHrspcsERPSTD','ManHrspcsMAPSTD','Source','Type'],
             var_name='PLAN_WEEK',
             value_name = 'PlanQty'
             ).sort_values(by='Org')

#只讀取Column的PLAN_WEEK字串中的前8個數值
df['PLAN_WEEK'] = df['PLAN_WEEK'].str[:8]

#於excel第10列插入Date
df.insert(10, 'Date', Date1)
#獲取當下時間
t = time.localtime()
UpdatedTime = time.strftime("%Y/%m/%d, %H:%M:%S", t)
#於excel第11列插入UpdatedTime
df.insert(11, 'UpdatedTime', UpdatedTime)


#連接SQL SERVER
conn = pyodbc.connect(
    Trusted_Connection='no',
    DRIVER='{ODBC Driver 17 for SQL Server}',
    server='Server name',
    DATABASE='DB name',
    UID='User ID',
    PWD='Password')

cursor = conn.cursor()
#cursor.execute('DROP TABLE Hans.dbo.forecast_test')

#數據上傳資料庫
for row in df.itertuples():
    cursor.execute('''
                INSERT INTO Forecast_MPS_Week (Org,Item,ItemDesc,OSP_Flag,ManHrsPcsERPSTD,ManHrsPcsMAPSTD,Source,Type,PLAN_WEEK,PlanQty,Date,UpdatedTime)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
                ''',
                    row.Org, 
                    row.Item,
                    row.ItemDesc,
                    row.OSP_Flag,
                    row.ManHrspcsERPSTD,
                    row.ManHrspcsMAPSTD,
                    row.Source,
                    row.Type,
                    row.PLAN_WEEK,
                    row.PlanQty,
                    row.Date,
                    row.UpdatedTime
                    
                )   


conn.commit()

cursor.close
#上傳成功顯示Upload Successfully
print('Upload Successfully')