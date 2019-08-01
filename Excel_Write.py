import pandas as pd
import json
import requests

# Create a Pandas dataframe from the data.
#df = pd.DataFrame({'Email_To': [10, 20, 30, 20, 15, 30, 45]})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('D:\\Email_CC.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
#df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
#writer.save()


#Invalid Email_CC

Invalid_mail = requests.get("http://10.160.1.217:9090/MISTicketGenerator/validateMISEmailId/email_cc")
aa = {}
aa=json.loads(Invalid_mail.text)
bb = []
bb=aa['Invalid EmailId']
print("Invalid Email_CC\n")
for i in bb:
	print(i)
	i.to_excel(writer, sheet_name='Sheet1')

	
print("\n")
writer.save()
