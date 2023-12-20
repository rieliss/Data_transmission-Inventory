import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Master_Process.xlsm', engine = 'openpyxl')
df = df.fillna(value=0)


server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO master_process (MSTPC, PROCS, DEVID1, DEVID2, DEVID3, DEVID32, DEVID4, DEVID5, Process_Name, Process_type, Status, CT, OA, WKTIME, WKSHIFT, CAP_No_OT, CAP_OT_25, CAP_OT_50, Risk_Stock_Target, Risk_Stock_Max, Risk_Stock_Min) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row.MSTPC, row.PROCS, row.DEVID1, row.DEVID2, row.DEVID3, row.DEVID32, row.DEVID4, row.DEVID5, row.Process_Name, row.Process_type, row.Status, row.CT, row.OA, row.WKTIME, row.WKSHIFT, row.CAP_No_OT, row.CAP_OT_25, row.CAP_OT_50, row.Risk_Stock_Target, row.Risk_Stock_Max, row.Risk_Stock_Min)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE master_process SET PROCS = ?, DEVID1 = ?, DEVID2 = ?, DEVID3 = ?, DEVID32 = ?, DEVID4 = ?, DEVID5 = ?, Process_Name = ?, Process_type = ?, Status = ?, CT = ?, OA = ?, WKTIME = ?, WKSHIFT = ?, CAP_No_OT = ?, CAP_OT_25 = ?, CAP_OT_50 = ?, Risk_Stock_Target = ?, Risk_Stock_Max = ?, Risk_Stock_Min = ? WHERE MSTPC = ?;", row.PROCS, row.DEVID1, row.DEVID2, row.DEVID3, row.DEVID32, row.DEVID4, row.DEVID5, row.Process_Name, row.Process_type, row.Status, row.CT, row.OA, row.WKTIME, row.WKSHIFT, row.CAP_No_OT, row.CAP_OT_25, row.CAP_OT_50, row.Risk_Stock_Target, row.Risk_Stock_Max, row.Risk_Stock_Min, row.MSTPC)
          print("Update")

print(df)
cnxn.commit()
cursor.close()