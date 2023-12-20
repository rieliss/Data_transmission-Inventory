import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Notify/TPM_Notify.xlsx', engine = 'openpyxl')
df = df.fillna(value=0)


server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()

for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO tpm_notify (ID, Information, Class, MC_No, Event_Time, textDate, textTime, Recieve_s_code, Contact_to, Tel_to, Manager_Name, Manager_email, Level, Line_Name, F_Location, Incharge, Response_time, Rating) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row.ID, row.Information, row.Class, row.MC_No, row.Event_Time, row.textDate, row.textTime, row.Recieve_s_code, row.Contact_to, row.Tel_to, row.Manager_Name, row.Manager_email, row.Level, row.Line_Name, row.F_Location, row.Incharge, row.Response_time, row.Rating)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE tpm_notify SET Information = ?, Class = ?, MC_No = ?, Event_Time = ?, textDate = ?, textTime = ?, Recieve_s_code = ?, Contact_to = ?, Tel_to = ?, Manager_Name = ?, Manager_email = ?, Level = ?, Line_Name = ?, F_Location = ?, Incharge = ?, Response_time = ?, Rating = ? WHERE ID = ?;", row.Information, row.Class, row.MC_No, row.Event_Time, row.textDate, row.textTime, row.Recieve_s_code, row.Contact_to, row.Tel_to, row.Manager_Name, row.Manager_email, row.Level, row.Line_Name, row.F_Location, row.Incharge, row.Response_time, row.Rating, row.ID)
          print("Update")

print(df)
cnxn.commit()
cursor.close()