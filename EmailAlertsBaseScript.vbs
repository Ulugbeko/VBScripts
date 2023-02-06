
'   * Declaring variables for actual date , most recent file date and for path of the folder:

Dim fNewest, newestFIleDate, dateVar , actualDate 
Const folderPath
folderPath = "C:\Tester"


' Getting actual date in output format: month/day/year
    dateVar=now
    actualDate = (Month(dateVar) & "/" & Day(dateVar)  & "/" & Year(dateVar))
        'output format: yyyymmddHHnn
        ' wscript.echo ((year(dt)*100 + month(dt))*100 + day(dt))*10000 + hour(dt)*100 + minute(dt)


' Finding and getting most recent updated text files time in output format: yyyymmddHHnn
    set oFolder=createobject("scripting.filesystemobject").getfolder(folderPath)
    For Each aFile In oFolder.Files
        If fNewest = "" Then
            Set fNewest = aFile
        Else
            If fNewest.DateCreated < aFile.DateCreated Then
                Set fNewest = aFile
            End If
        End If
    Next
newestFIleDate = FormatDateTime(fNewest.DateLastModified,2)

'  * To determine the most recent file in the specified folder:
dim sMostRecent, dMostRecent
MostRecent("D:\Temp")
WScript.Echo sMostRecent, dMostRecent
Sub MostRecent (sFolder)

Set oFSO = CreateObject("Scripting.FileSystemObject")

dMostRecent = 0
sMostRecent = ""
For Each oFile In oFSO.GetFolder(sFolder).Files
dFileDate = oFile.DateLastModified
If dFileDate > dMostRecent Then
dMostRecent = dFileDate
sMostRecent = oFile.Path
End If
Next
End Sub

' With If statement validating capture process of load and sending email alert

If (newestFIleDate = actualDate) Then
WScript.Echo " the date is matching "
Else
WScript.Echo "the date is not matching"
End If


const cdoBasic=1
schema = "http://schemas.microsoft.com/cdo/configuration/"
Set objEmail = CreateObject("CDO.Message")
With objEmail
.From = "Ja...@company.com"
.To = "J...@company.com"
.Subject = "Test Mail"
.Textbody = "The quick brown fox " & Chr(10) & "jumps over the lazy dog"
.AddAttachment "d:\Testfile.txt"
With .Configuration.Fields
.Item (schema & "sendusing") = 2
.Item (schema & "smtpserver") = "mail.company.com"
.Item (schema & "smtpserverport") = 25
.Item (schema & "smtpauthenticate") = cdoBasic
.Item (schema & "sendusername") = "Ja...@company.com"
.Item (schema & "smtpaccountname") = "Ja...@company.com"
.Item (schema & "sendpassword") = "SomePassword"
End With
.Configuration.Fields.Update
.Send
End With


WScript.Echo fNewest.Name
WScript.Echo newestFIleDate
WScript.Echo actualDate
