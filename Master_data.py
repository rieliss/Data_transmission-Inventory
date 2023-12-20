import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Master data.xlsm', engine = 'openpyxl')
df = df.fillna(value=0)


server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO master_data (ITNBR, ITDSC, ENGNO, ITTYP, ITCLS, PLANN, MULQY, LOTSZ, PROCS, UNTSZ, FRQTY, Suffix, MSTMD) values(?,?,?,?,?,?,?,?,?,?,?,?,?)", row.ITNBR, row.ITDSC, row.ENGNO, row.ITTYP, row.ITCLS, row.PLANN, row.MULQY, row.LOTSZ, row.PROCS, row.UNTSZ, row.FRQTY, row.Suffix, row.MSTMD)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE master_data SET ITNBR = ?, ITDSC = ?, ENGNO = ?, ITTYP = ?, ITCLS = ?, PLANN = ?, MULQY = ?, LOTSZ = ?, PROCS = ?, UNTSZ = ?, FRQTY = ?, Suffix = ? WHERE MSTMD = ?;", row.ITNBR, row.ITDSC, row.ENGNO, row.ITTYP, row.ITCLS, row.PLANN, row.MULQY, row.LOTSZ, row.PROCS, row.UNTSZ, row.FRQTY, row.Suffix, row.MSTMD)
          print("Update")

print(df)
cnxn.commit()
cursor.close()