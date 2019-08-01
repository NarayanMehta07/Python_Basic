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
statement = """DELETE FROM EXCEL_DUMP"""
cur.execute(str(statement))
print ("delete EXCEL_DUMP successfully")
#res = cur.fetchall()
#print (res)
#print("Record retrive succesfully")
book = xlrd.open_workbook("BatchIDOutput.xlsx")
sheet = book.sheet_by_name("Sheet1")
# Create the INSERT INTO sql query
query = """INSERT INTO EXCEL_DUMP (policy_id, cvg_num, val, col) VALUES """
print ("total rows in excel file",sheet.nrows)
# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
for r in range(1, sheet.nrows):
		POL_ID		= sheet.cell(r,0).value
		cvg_num	= sheet.cell(r,1).value
		val			= sheet.cell(r,2).value
		col		= sheet.cell(r,3).value
		values = (POL_ID, cvg_num, val, col)
		i= (str(query) + str(values))
		cur.execute(str(query) + str(values))



# Commit the transaction
db.commit()
print ("Data Insert")

# Print results
print ("")
print ("All Done!")
print ("")
columns = str(sheet.ncols)
rows = str(sheet.nrows)
print ("total column =",str(sheet.ncols),"Total rows=",
       str(sheet.nrows),"\nI just imported total ",+r ,"to Oracle Database!")

#Call Procedure to update Old table



#myvar = cur.var(cx_Oracle.NUMBER)
#cur.callproc('prc_updating_valuation_rec')
#i ='myproc'
#cur.callproc(i)


print ("All Code are running successfully")
cur.close()
db.close()
        
