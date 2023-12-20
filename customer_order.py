import pyodbc
import pandas as pd

df = pd.read_excel(f'//172.23.3.44/public/Inventory/Customer order.xlsm', engine = 'openpyxl')
df = df.fillna(value=0)


server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)

cursor = cnxn.cursor()



for index, row in df.iterrows():
     try:
          cursor.execute("INSERT INTO customer_order (KEYID,CUSNO,SHPNO,CORNO,OGQTY,ORQTY,ALQTY,SHDQY,ODRFL,DUEDT,CPRTN,CUSPO,LOTNO,LMNDT,CUSNO2,SHPNO3,SHPCT,SHPAA,SHPCL,SHPTM,CARNM,LDUPD,SHPNT,SHPAA4,ITDSC,INTYP,ITCLS,PROCS,M,Suffix1,Suffix2,TYPE,SHPDT,PsRTNO,PNO) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",row.KEYID,row.CUSNO,row.SHPNO,row.CORNO,row.OGQTY,row.ORQTY,row.ALQTY,row.SHDQY,row.ODRFL,row.DUEDT,row.CPRTN,row.CUSPO,row.LOTNO,row.LMNDT,row.CUSNO2,row.SHPNO3,row.SHPCT,row.SHPAA,row.SHPCL,row.SHPTM,row.CARNM,row.LDUPD,row.SHPNT,row.SHPAA4,row.ITDSC,row.INTYP,row.ITCLS,row.PROCS,row.M,row.Suffix1,row.Suffix2,row.TYPE,row.SHPDT,row.PRTNO,row.PNO)
          print("Insert")
     except:
          pass
          cursor.execute("UPDATE master_customer_name SET CUSNO = ?,SHPNO = ?,CORNO = ?,OGQTY = ?,ORQTY = ?,ALQTY = ?,SHDQY = ?,ODRFL = ?,DUEDT = ?,CPRTN = ?,CUSPO = ?,LOTNO = ?,LMNDT = ?,CUSNO2 = ?,SHPNO3 = ?,SHPCT = ?,SHPAA = ?,SHPCL = ?,SHPTM = ?,CARNM = ?,LDUPD = ?,SHPNT = ?,SHPAA4 = ?,ITDSC = ?,INTYP = ?,ITCLS = ?,PROCS = ?,M = ?,Suffix1 = ?,Suffix2 = ?,TYPE = ?,SHPDT = ?,PRTNO = ?, PNO = ? WHERE KEYID = ?;",row.KEYID,row.CUSNO,row.SHPNO,row.CORNO,row.OGQTY,row.ORQTY,row.ALQTY,row.SHDQY,row.ODRFL,row.DUEDT,row.CPRTN,row.CUSPO,row.LOTNO,row.LMNDT,row.CUSNO2,row.SHPNO3,row.SHPCT,row.SHPAA,row.SHPCL,row.SHPTM,row.CARNM,row.LDUPD,row.SHPNT,row.SHPAA4,row.ITDSC,row.INTYP,row.ITCLS,row.PROCS,row.M,row.Suffix1,row.Suffix2,row.TYPE,row.SHPDT,row.PRTNO,row.PNO)
          print("Update")
     else:
          print("Failed!")

print(df)
cnxn.commit()
cursor.close()