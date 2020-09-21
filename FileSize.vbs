Option Explicit
On Error Resume Next

Dim TSO
Set TSO = CreateObject("Scripting.FileSystemObject")
WScript.Echo (TSO.GetFile("C:\Campaign_Mails\RBNA_Res.csv").Size/1024)

Set TSO =  Nothing