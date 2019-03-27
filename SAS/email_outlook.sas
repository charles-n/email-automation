data _null_;
	file "D:\email_notif.ps1";
		put '[String[]]$Send_list = "charles.nguyen@hilton.com","charles.nguyen@mavs.uta.edu" ';

		put '$Outlook = New-Object -ComObject Outlook.Application ';
		put '$Mail = $Outlook.CreateItem(0) ';
		put '$Send_list | % {$Mail.Recipients.add($_)} ';
		put '$Mail.Subject = "Live Leads Scoring Process Complete" ';
		put "$Mail.Body = echo 'Export complete -- number of rows: &cnt.'; ";

		/*put '$Mail.Attachments.Add("D:\Pictures\SAS_Foundation_Software.PNG")';*/
		/*put '$Mail.Attachments.Add("D:\Pictures\SAS_Installation_Data_File.PNG")';*/

		put '$Mail.Send()';
run;

data _null_;
	call system("powershell D:\email_notif.ps1 > D:\email_notif.log  2>&1");
run;
