import json
import requests
import pandas as pd


#Invalid Email_CC
Invalid_mail = requests.get("http://10.160.1.217:9090/MISTicketGenerator/validateMISEmailId/email_cc")
aa = {}
aa=json.loads(Invalid_mail.text)
bb = []
bb=aa['Invalid EmailId']
df = pd.DataFrame(bb)
df = pd.DataFrame({'Email_CC': bb})
writer = pd.ExcelWriter('D:\\Emails_Error.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Email_Ids',index=False)


print("Invalid Email_CC File Generated Successfully\n")


#Invalid Email_to
Invalid_mail = requests.get("http://10.160.1.217:9090/MISTicketGenerator/validateMISEmailId/email_to")
aa = {}
aa=json.loads(Invalid_mail.text)
bb = []
bb=aa['Invalid EmailId']
df1 = pd.DataFrame(bb)
df1 = pd.DataFrame({'Email_TO': bb})
df1.to_excel(writer, sheet_name='Email_Ids',index=False, startcol=1)


print("Invalid Email_TO File Generated Successfully\n")

#Invalid Email_to
Invalid_mail = requests.get("http://10.160.1.217:9090/MISTicketGenerator/validateMISEmailId/email_bcc")
aa = {}
aa=json.loads(Invalid_mail.text)
bb = []
bb=aa['Invalid EmailId']
df2 = pd.DataFrame(bb)
df2 = pd.DataFrame({'Email_BCC': bb})
df2.to_excel(writer, sheet_name='Email_Ids',index=False,startcol=2)
writer.save()

print("Invalid Email_BCC File Generated Successfully\n")





