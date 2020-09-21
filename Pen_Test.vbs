Option Explicit
On Error Resume Next

Dim FSO, RSO, FLD, fil, WS, RS, SLine
Dim strFolder
Const ForReading = 1, ForWriting = 2, ForAppending = 8 


strFolder = "C:\Users\PR172959\Downloads\Region Eligible User\Unprocess"
Set FSO = CreateObject("Scripting.FileSystemObject")
Set RSO = CreateObject("Scripting.FileSystemObject")
set FLD = FSO.GetFolder(strFolder)

For Each fil In FLD.Files
Set WS = FSO.CreateTextFile(fil.Path+"_modify.csv", ForWriting)
Set RS = RSO.OpenTextFile(fil.Path, ForReading)

Do Until RS.AtEndOfStream
SLine = RS.ReadLine
SLine = Replace(SLine,chr(34),"")
'WScript.Echo SLine & vbCrLf
WS.Write (SLine) & vbCrLf


Loop
WS.Close
RS.Close
Set WS = Nothing
Set RS = Nothing	
Next

	
Set FLD = Nothing
Set FSO = Nothing
