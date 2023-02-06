strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
' C drive is specified at Device ID
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_LogicalDisk where DeviceID='c:'",,48) 
' Preset value for sending email using AUTH LOGIN
cdoBasic = 1

For Each objItem in colItems
   if len(objItem.VolumeName) > 0 then
      freeSpace = FormatNumber((CDbl(objItem.FreeSpace)/1024/1024/1024))
      diskSize = FormatNumber((CDbl(objItem.Size)/1024/1024/1024))
      usedSpace = diskSize - freeSpace
      percentage = FormatNumber((CDbl(objItem.FreeSpace)/1024/1024/1024)) / _
         FormatNumber((CDbl(objItem.Size)/1024/1024/1024)) * 100
      diskDescription = "-----------------------------------" & vbCrLf _
           & "VolumeName: " & vbTab & objItem.VolumeName  & vbCrLf _
           & "-----------------------------------" & vbCrLf _
           & "Free Space: " & vbTab & freeSpace & " GB" & vbCrLf _
           & "Total Size: " & vbTab & diskSize & " GB" & vbCrLf _
           & "Occupied Space: " & usedSpace & " GB" & vbTab
      Wscript.Echo diskDescription
	      if percentage < 20 then
		         Set objEmail = CreateObject("CDO.Message")
         		objEmail.From = "youremailaddress@yourdomain.com"
		         objEmail.To = "yourreceipient@domain.com"
         		objEmail.Subject = "FREE SPACE ON C DRIVE IS LOWER THAN " & _
            percentage & "%" 
      		   objEmail.Textbody = diskDescription
      		   objEmail.Configuration.Fields.Item _
             			("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
         'Type of authentication, NONE, Basic (Base64 encoded), NTLM
         		objEmail.Configuration.Fields.Item _
			            ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
		         'Your UserID on the SMTP server
         		objEmail.Configuration.Fields.Item _
            			("http://schemas.microsoft.com/cdo/configuration/sendusername") = "youremail@yourdomain.com" 
		         'Your password on the SMTP server
         		objEmail.Configuration.Fields.Item _
            			("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "yourpassword"
		         objEmail.Configuration.Fields.Item _
            			("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "yourSMTPmailserver" 
		         objEmail.Configuration.Fields.Item _
            			("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		         'Use SSL for the connection (False or True)
         		objEmail.Configuration.Fields.Item _
            			("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
         'The connection will timeout in 60 seconds
		         objEmail.Configuration.Fields.Item _
            			("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
         		objEmail.Configuration.Fields.Update
         		objEmail.Send
      		end if
   end if
Next
