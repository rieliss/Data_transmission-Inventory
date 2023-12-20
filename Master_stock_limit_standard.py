import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Master_Stock limit standard.xlsm', engine = 'openpyxl')
df = df.fillna(value=0)


server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO master_stock_limit_standard (MSTST, PRTNO, PROCS, Max_DOH, Min_DOH) values(?,?,?,?,?)", row.MSTST, row.PRTNO, row.PROCS, row.Max_DOH, row.Min_DOH)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE master_stock_limit_standard SET PRTNO = ?, PROCS = ?, Max_DOH = ?, Min_DOH = ? WHERE MSTST = ?;", row.PRTNO, row.PROCS, row.Max_DOH, row.Min_DOH, row.MSTST)
          print("Update")

print(df)
cnxn.commit()
cursor.close()