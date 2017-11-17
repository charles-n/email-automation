### Requires three windows applications
    # Microsoft Outlook 2016
    # Windows PowerShell Version 5.1.14393.1770  
    # R version 3.4.2

ps1 <- c(
    "powershell ",
    "$Outlook = New-Object -ComObject Outlook.Application ", 
    "$Mail = $Outlook.CreateItem(0) ", 
    
    ### list email addresses, you can get clever by using an R vector of emails or for loops
    "$Send_list = 'charles.nguyen@hilton.com','charles_nguyen@neimanmarcus.com' ", 
    "$Send_list | % {$Mail.Recipients.add($_)} ", 
    
    ### email subject line
    "$Mail.Subject = 'Model Forecast Process Failed/Successful' ", 
    
    ### email body text
    "$Mail.Body = echo 'Please view log **file path**' ", 
    
    ### first attachment
    # "$Mail.Attachments.Add('D:/Pictures/SAS_Foundation_Software.PNG')",
    
    ### second attachment
    # "$Mail.Attachments.Add('D:/Pictures/SAS_Installation_Data_File.PNG')",

    "$Mail.Send()"
  )

### send email
system(paste(ps1, collapse='\n '))
