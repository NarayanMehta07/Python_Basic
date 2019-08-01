import cx_Oracle
import xlrd
import csv
from xlsxwriter.workbook import Workbook
from twilio.rest import Client
import subprocess

ip = '10.160.1.22'
port=1521
SID='etlife1'
dsn_tns= cx_Oracle.makedsn(ip, port, SID)
db = cx_Oracle.connect('ibexi', 'ibexi', dsn_tns)
print("Connection Established")
cur = db.cursor()
#connection
statement = """SELECT REPORT_NAME FROM MIS_MAILING_LIST"""
cur.execute(str(statement))
res = cur.fetchall()
for x in res:
  print(x)

print("Record retrive succesfully")

print ("All Code are running successfully")
cur.close()
db.close()
