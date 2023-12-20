import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Subassy requirement.xlsm', engine = 'openpyxl')
df = df.fillna(value=0)


server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO subassy_requirement (KEYID, CPANO, ITDSC, BPRTNO, QTYPR, PRODUC, YYMD, EDATM, EDATO, QTY, Column3, Column1, Column2, Type) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row.KEYID, row.CPANO, row.ITDSC, row.BPRTNO, row.QTYPR, row.PRODUC, row.YYMD, row.EDATM, row.EDATO, row.QTY, row.Column3, row.Column1, row.Column2, row.Type)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE subassy_requirement SET CPANO = ?, ITDSC = ?, BPRTNO = ?, QTYPR = ?, PRODUC = ?, YYMD = ?, EDATM = ?, EDATO = ?, QTY = ?, Column3 = ?, Column1 = ?, Column2 = ?, Type = ? WHERE KEYID = ?;", row.CPANO, row.ITDSC, row.BPRTNO, row.QTYPR, row.PRODUC, row.YYMD, row.EDATM, row.EDATO, row.QTY, row.Column3, row.Column1, row.Column2, row.Type, row.KEYID)
          print("Update")

print(df)
cnxn.commit()
cursor.close()