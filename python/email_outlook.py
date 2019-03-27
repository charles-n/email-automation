### python on windows machine
#	requires python, powershell, outlook installed on your computer


import subprocess 
import sys
import functools

def reduce_concat(x, sep=""):
    return functools.reduce(lambda x, y: str(x) + sep + str(y), x)

def paste(*lists, sep=" ", collapse=None):
    result = map(lambda x: reduce_concat(x, sep=sep), zip(*lists))
    if collapse is not None:
        return reduce_concat(result, sep=collapse)
    return list(result)


### email input
ps1 = [
	"$Outlook = New-Object -ComObject Outlook.Application ", 
	"$Mail = $Outlook.CreateItem(0) ", 

	### list email addresses, you can get clever by using an R vector of emails or for loops
	"$Send_list = 'charles.nguyen@hilton.com', 'charles.nguyen@hilton.com' ", 
	"$Send_list | % {$Mail.Recipients.add($_)} ", 

	### email subject line
	"$Mail.Subject = 'This is the Subject Line' ", 

	### email body text
	"$Mail.Body = echo 'This is the body of the email' ",

	### for file attachment
	# "$Mail.Attachments.Add('//dcscfile01/HotOps/DATA SCIENCE/projects/edison/inventory_loads/output/inventory_bot_los_ota.txt')",
	"$Mail.Send()",
	]

cmd = paste(ps1, collapse=' \n ')

### send email
returned_value = subprocess.Popen(["powershell.exe", cmd], stdout=sys.stdout) 