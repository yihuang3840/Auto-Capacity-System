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

#讀取CAPA_Report資料夾最新檔, 檔名前面為'Manpower(Monthly)_Based_on_MPS_*.xls', *為模糊搜尋日期(如20231122), 且資料夾中沒有_TEST字串的檔案
p=Path('N:\DataExchange\Production_Control\CAPA Report')
a=max([fn for fn in p.glob('Manpower(Monthly)_Based_on_MPS_*.xls') if not '_Test' in str(fn)], key=lambda f: f.stat().st_mtime)


#搜尋檔名中的日期(如20231122)並擷取
a1 = str(a)
print (a1)
Date1 = re.search('on_MPS_(.*?).xl', a1).group(1)
print ('FileDate:', Date1)
print('Data Uploading...')

#讀取最新檔的Excel, 參考來自15行
wb = pd.read_excel(a)

#轉excel格式
wb.to_excel('excel.xlsx', index=False)

#將excel load進並且更改column名稱, 因原column名稱有空格, sql server 匯入會檔
wb = load_workbook('excel.xlsx')
sheet = wb.active
sheet["C1"] = "ItemDesc"; sheet["G1"] = "ReferenceRsc"; sheet["H1"] = "RscType"; sheet["I1"]= "OSPFlag"
sheet["J1"] = "MapFlag"; sheet["M1"] = "ReferenceItem"; sheet["N1"] = "ReferenceLevel"; sheet["O1"] = "ManHrspcsERPActual"
sheet["P1"] = "ManHrspcsERPSTD"; sheet["Q1"] = "ManHrspcsMAPSTD"; sheet["R1"] = "MainProcess"

sheet["U1"] = "Month1" ; sheet["V1"] = "Month2" ; sheet["W1"] = "Month3" ; sheet["X1"] = "Month4"
sheet["Y1"] = "Month5" ; sheet["Z1"] = "Month6" ; sheet["AA1"] = "Month7" ; sheet["AB1"] = "Month8"
sheet["AC1"] = "Month9" ; sheet["AD1"] = "Month10" ; sheet["AE1"] = "Month11" ; sheet["AF1"] = "Month12"
#儲存至excel
wb.save(filename= 'D:/forecast_test.xlsx')

#讀取excel, 用pandas讀
data = pd.read_excel ('D:/forecast_test.xlsx', keep_default_na = False)   
df = pd.DataFrame(data)
#於excel第32列插入Date, 並將21行的擷取日期匯入
df.insert(32, 'Date', Date1)
#匯入上傳時的當下時間
t = time.localtime()
UpdatedTime = time.strftime("%Y/%m/%d, %H:%M:%S", t)
#於excel第33列插入UpdatedTime
df.insert(33, 'UpdatedTime', UpdatedTime)

#連接SQL SERVER
conn = pyodbc.connect(
    Trusted_Connection='no',
    DRIVER='{ODBC Driver 17 for SQL Server}',
    server='Server name',
    DATABASE='DB name',
    UID='User ID',
    PWD='Password')

cursor = conn.cursor()
#先刪除SQL SERVER Table 
cursor.execute('DROP TABLE dbo.forecast_test')


#後續create新table
cursor.execute('''
		CREATE TABLE forecast_test (
            Org nvarchar(200),
			Item nvarchar(200),
            ItemDesc nvarchar(200),
            Cat1 nvarchar(200),
            Cat2 nvarchar(200),
            Cat3 nvarchar(200),
            ReferenceRsc nvarchar(200),
            RscType nvarchar(200),
            OSPFlag nvarchar(200),
            MapFlag nvarchar(200),
            UPH nvarchar(200),
            Usage nvarchar(200),
            ReferenceItem nvarchar(200),
            ReferenceLevel nvarchar(200),
            ManHrspcsERPActual nvarchar(200),
            ManHrspcsERPSTD nvarchar(200),
            ManHrspcsMAPSTD nvarchar(200),
            MainProcess nvarchar(200),
            Source nvarchar(200),
            Type nvarchar(200),
            Month1 float(50),
            Month2 float(50),
            Month3 float(50),
            Month4 float(50),
            Month5 float(50),
            Month6 float(50),
            Month7 float(50),
            Month8 float(50),
            Month9 float(50),
            Month10 float(50),
            Month11 float(50),
            Month12 float(50),
            Date nvarchar(200),
            UpdatedTime nvarchar(200)
            )
            ''')


#資料匯入SQL SERVER
for row in df.itertuples():
    cursor.execute('''
                INSERT INTO forecast_test (Org,Item,ItemDesc,Cat1,Cat2,Cat3,ReferenceRsc,RscType,OSPFlag,MapFlag,UPH,Usage,ReferenceItem,ReferenceLevel,ManHrspcsERPActual,ManHrsPcsERPSTD,ManHrsPcsMAPSTD,MainProcess,Source,Type,Month1,Month2,Month3,Month4,Month5,Month6,Month7,Month8,Month9,Month10,Month11,Month12,Date,UpdatedTime)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''',
                    row.Org, 
                    row.Item,
                    row.ItemDesc,
                    row.Cat1,
                    row.Cat2,
                    row.Cat3,
                    row.ReferenceRsc,
                    row.RscType,
                    row.OSPFlag,
                    row.MapFlag,
                    row.UPH,
                    row.Usage,
                    row.ReferenceItem,
                    row.ReferenceLevel,
                    row.ManHrspcsERPActual,
                    row.ManHrspcsERPSTD,
                    row.ManHrspcsMAPSTD,
                    row.MainProcess,
                    row.Source,
                    row.Type,
                    row.Month1,
                    row.Month2,
                    row.Month3,
                    row.Month4,
                    row.Month5,
                    row.Month6,
                    row.Month7,
                    row.Month8,
                    row.Month9,
                    row.Month10,
                    row.Month11,
                    row.Month12,
                    row.Date,
                    row.UpdatedTime
                    
                )   


conn.commit()

cursor.close
#刪除留在D槽的檔案
os.remove("D:/forecast_test.xlsx")

#上傳成功, 回傳Upload Successfully
print('Upload Successfully')