import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Master_Customer name.xlsm', engine = 'openpyxl')
df = df.fillna(value=0)


server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     print("INSERT INTO master_customer_name (Suffix,Type,CustomerName) values(?,?,?)", row.Suffix, row.Type, row.CustomerName)
     print("UPDATE master_customer_name SET Type = ?, CustomerName = ? WHERE Suffix = ?;", row.Type, row.CustomerName, row.Suffix)
     try:
          cursor.execute("INSERT INTO master_customer_name (Suffix,Type,CustomerName) values(?,?,?)", row.Suffix, row.Type, row.CustomerName)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE master_customer_name SET Type = ?, CustomerName = ? WHERE Suffix = ?;", row.Type, row.CustomerName, row.Suffix)
          print("Update")

print(df)
cnxn.commit()
cursor.close()