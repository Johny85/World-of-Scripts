Option Explicit
On Error Resume Next

Dim objFso, objFile, DateC, stc
DateC = Date()
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFile = objFso.GetFile("C:\Campaign_Mails\RBNA_Res.csv")
WScript.Echo Left(objFile.DateLastModified,10)
stc = StrComp(DateC,Left(objFile.DateLastModified,10))
WScript.Echo (stc)

Set objFso = Nothing
Set objFile = Nothing
