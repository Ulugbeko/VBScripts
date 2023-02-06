On Error Resume Next

'Dim some stuffs
Dim objFso, objFolder, objFile, strFolderPath, intDaysOlderThan
Set objFso = CreateObject("Scripting.FileSystemObject")

'Change these variables
strFolderPath = "C:\Temp"
intDaysOlderThan = 28

'Do the magic
Call DeleteFiles(strFolderPath, intDaysOlderThan)


'Procedures
Sub DeleteFiles(path, days)
  Set objFolder = objFso.GetFolder(path)
  For Each objFile In objFolder.Files
    If objFile.DateLastModified < (Now() - days) Then
      objFile.Delete(True)
    End If
  Next
End Sub
