import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Master_Inventory cost.xlsm', engine = 'openpyxl')
df = df.fillna(value=0)


server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO master_inventory_cost (Run_No, ITNBR, ITDSC, ITCLS, MNFCS, ITTYP) values(?,?,?,?,?,?)", row.Run_No, row.ITNBR, row.ITDSC, row.ITCLS, row.MNFCS, row.ITTYP)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE master_inventory_cost SET ITNBR = ?, ITDSC = ?, ITCLS = ?, MNFCS = ?, ITTYP = ? WHERE Run_No = ?;", row.ITNBR, row.ITDSC, row.ITCLS, row.MNFCS, row.ITTYP, row.Run_No)
          print("Update")

print(df)
cnxn.commit()
cursor.close()