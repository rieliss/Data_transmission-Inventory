import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Stock out QR Shopper BPK (from SQL Saver).xlsm', engine = 'openpyxl')
df = df.fillna(value=0)
print(df)

server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO stock_out (KEYID, DeviceID, WareHouse, IDNo, TRDate, Result, ErrCode, TagDataCode1, DensoPartNo1, CustPartNo1, Quantity1, TagSeqNo1, ScanDateTime1, Date1, Time1, TagDataCode2, DensoPartNo2, CustPartNo2, Quantity2, TagSeqNo2, ScanDateTime2, Date2, Time2, File_Name, EID) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", row.KEYID, row.DeviceID, row.WareHouse, row.IDNo, row.TRDate, row.Result, row.ErrCode, row.TagDataCode1, row.DensoPartNo1, row.CustPartNo1, row.Quantity1, row.TagSeqNo1, row.ScanDateTime1, row.Date1, row.Time1, row.TagDataCode2, row.DensoPartNo2, row.CustPartNo2, row.Quantity2, row.TagSeqNo2, row.ScanDateTime2, row.Date2, row.Time2, row.File_Name, row.EID)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE stock_out SET DeviceID = ?, WareHouse = ?, IDNo = ?, TRDate = ?, Result = ?, ErrCode = ?, TagDataCode1 = ?, DensoPartNo1 = ?, CustPartNo1 = ?, Quantity1 = ?, TagSeqNo1 = ?, ScanDateTime1 = ?, Date1 = ?, Time1 = ?, TagDataCode2 = ?, DensoPartNo2 = ?, CustPartNo2 = ?, Quantity2 = ?, TagSeqNo2 = ?, ScanDateTime2 = ?, Date2 = ?, Time2 = ?, File_Name = ?, EID = ? WHERE KEYID = ?;", row.DeviceID, row.WareHouse, row.IDNo, row.TRDate, row.Result, row.ErrCode, row.TagDataCode1, row.DensoPartNo1, row.CustPartNo1, row.Quantity1, row.TagSeqNo1, row.ScanDateTime1, row.Date1, row.Time1, row.TagDataCode2, row.DensoPartNo2, row.CustPartNo2, row.Quantity2, row.TagSeqNo2, row.ScanDateTime2, row.Date2, row.Time2, row.File_Name, row.EID, row.KEYID)
          print("Update")

print(df)
cnxn.commit()
cursor.close()