import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Master_Begin_Stock.xlsx', engine = 'openpyxl')
df = df.fillna(value=0)

server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO master_begin (part_no, part_name, model, suffix, in_QTY, Out_QTY, stock_qty, BEGINQTY, Beginstk) values(?,?,?,?,?,?,?,?,?)", row.part_no, row.part_name, row.model, row.suffix, row.in_QTY, row.Out_QTY, row.stock_qty, row.BEGINQTY, row.Beginstk)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE master_begin SET part_name = ?, model = ?, suffix = ?, in_QTY = ?, Out_QTY = ?, stock_qty = ?, BEGINQTY = ?, Beginstk = ? WHERE part_no = ?;", row.part_name, row.model, row.suffix, row.in_QTY, row.Out_QTY, row.stock_qty, row.BEGINQTY, row.Beginstk, row.part_no)
          print("Update")

print(df)
cnxn.commit()
cursor.close()