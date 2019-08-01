import json
import requests
#Invalid Email_CC
filename = "D:\\Email_CC.xls"
file = open(filename, "w")
file.write('Email_CC')
file.write("\n")
Invalid_mail = requests.get("http://10.160.1.217:9090/MISTicketGenerator/validateMISEmailId/email_cc")
aa = {}
aa=json.loads(Invalid_mail.text)
bb = []
bb=aa['Invalid EmailId']
print("Invalid Email_CC\n")
for i in bb:
	print(i)
	file.write(i)
	file.write("\n")
	
print("\n")
file.close()

#Invalid Email_TO
filename = "D:\\Email_TO.xls"
file = open(filename, "w")
Invalid_mail = requests.get("http://10.160.1.217:9090/MISTicketGenerator/validateMISEmailId/email_to")
aa = {}
aa=json.loads(Invalid_mail.text)
bb = []
bb=aa['Invalid EmailId']
print("Invalid Email_TO\n")
for i in bb:
	print(i)
	file.write(i)
	file.write("\n")
 
	
print("\n")


#Invalid Email_BCC
filename = "D:\\Email_BCC.xls"
file = open(filename, "w")
Invalid_mail = requests.get("http://10.160.1.217:9090/MISTicketGenerator/validateMISEmailId/email_bcc")
aa = {}
aa=json.loads(Invalid_mail.text)
bb = []
bb=aa['Invalid EmailId']
print("Invalid Email_BCC\n")
for i in bb:
	print(i)
	file.write(i)
	file.write("\n")
	
file.close()
