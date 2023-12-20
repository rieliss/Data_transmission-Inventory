import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Stock in (from CIGMA ODBC).xlsm', engine = 'openpyxl')
df = df.fillna(value=0)

server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO stock_in (KEYID,PRTNO,QTYCD,TRQTY,TRNDT,PRCT6,PDTZ6,DEVID,TCODE,TRNNO,TAGCD,INSNO,HOUSE,SUCCD,SLPNO,NGDCD,REASN,PRICE,CURCY,PURUM,UMCNV,ITCLS,CLSNO,CPRNO,DUEDT,ORQTY,TERNO,SRADR,LINID,DATE,TextDate,QTY,SUFFIX,TYPECODE,TYPE,Time,Shift,TextTime,Nosuffix,formatted_date,EID) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", row.KEYID, row.PRTNO, row.QTYCD, row.TRQTY, row.TRNDT, row.PRCT6, row.PDTZ6, row.DEVID, row.TCODE, row.TRNNO, row.TAGCD, row.INSNO, row.HOUSE, row.SUCCD, row.SLPNO, row.NGDCD, row.REASN, row.PRICE, row.CURCY, row.PURUM, row.UMCNV, row.ITCLS, row.CLSNO, row.CPRNO, row.DUEDT, row.ORQTY, row.TERNO, row.SRADR, row.LINID, row.DATE, row.TextDate, row.QTY, row.SUFFIX, row.TYPECODE, row.TYPE, row.Time, row.Shift, row.TextTime, row.Nosuffix, row.formatted_date, row.EID)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE stock_in SET PRTNO = ?, QTYCD = ?, TRQTY = ?, TRNDT = ?, PRCT6 = ?, PDTZ6 = ?, DEVID = ?, TCODE = ?, TRNNO = ?, TAGCD = ?, INSNO = ?, HOUSE = ?, SUCCD = ?, SLPNO = ?, NGDCD = ?, REASN = ?, PRICE = ?, CURCY = ?, PURUM = ?, UMCNV = ?, ITCLS = ?, CLSNO = ?, CPRNO = ?, DUEDT = ?, ORQTY = ?, TERNO = ?, SRADR = ?, LINID = ?, DATE = ?, TextDate = ?, QTY = ?, SUFFIX = ?, TYPECODE = ?, TYPE = ?, Time = ?, Shift = ?, TextTime = ?, Nosuffix = ?, formatted_date = ?, EID = ? WHERE KEYID = ?;", row.PRTNO, row.QTYCD, row.TRQTY, row.TRNDT, row.PRCT6, row.PDTZ6, row.DEVID, row.TCODE, row.TRNNO, row.TAGCD, row.INSNO, row.HOUSE, row.SUCCD, row.SLPNO, row.NGDCD, row.REASN, row.PRICE, row.CURCY, row.PURUM, row.UMCNV, row.ITCLS, row.CLSNO, row.CPRNO, row.DUEDT, row.ORQTY, row.TERNO, row.SRADR, row.LINID, row.DATE, row.TextDate, row.QTY, row.SUFFIX, row.TYPECODE, row.TYPE, row.Time, row.Shift, row.TextTime, row.Nosuffix, row.formatted_date, row.EID, row.KEYID)
          print("Update")

print(df)
cnxn.commit()
cursor.close()