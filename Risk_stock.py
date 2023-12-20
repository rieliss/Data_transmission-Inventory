import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Risk stock.xlsm', engine = 'openpyxl')
df = df.fillna(value=0)


server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO risk_stock (KEYID, ITNBR, ITDSC, HOUSE, BEGINQTY, MOHTQ, PROCS, NOSUFFIX, WH) values(?,?,?,?,?,?,?,?,?)", row.KEYID, row.ITNBR, row.ITDSC, row.HOUSE, row.BEGINQTY, row.MOHTQ, row.PROCS, row.NOSUFFIX, row.WH)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE risk_stock SET ITNBR = ?, ITDSC = ?, HOUSE = ?, BEGINQTY = ?, MOHTQ = ?, PROCS = ?, NOSUFFIX = ?, WH = ? WHERE KEYID = ?;", row.KEYID, row.ITNBR, row.ITDSC, row.HOUSE, row.BEGINQTY, row.MOHTQ, row.PROCS, row.NOSUFFIX, row.WH)
          print("Update")

print(df)
cnxn.commit()
cursor.close()