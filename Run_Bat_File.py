import subprocess
filepath="C:\\Users\\narayanmehta\\Desktop\\VALUATION\\Valuation_Bat.bat"
p = subprocess.Popen(filepath, shell=True, stdout = subprocess.PIPE)
stdout, stderr = p.communicate()
print (p.returncode) # is 0 if success
