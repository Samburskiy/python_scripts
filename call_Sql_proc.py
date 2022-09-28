import pymssql as ms
import openpyxl as op



conn = ms.connect()
conn.autocommit(True)
cursor = conn.cursor()

cursor.callproc(name)
print('success')
