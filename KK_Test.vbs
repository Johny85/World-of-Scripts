Option Explicit
On Error Resume Next

Dim RSO, WSO, RS, WS, strFolder, FLD, fil, SLine

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

strFolder = "C:\Users\PR172959\Desktop\IMO_Workshop"
Set FLD = RSO.GetFolder(strFolder)
For Each fil In FLD.Files    
Set RS = RSO.OpenTextFile(fil.Path, ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Desktop\IMO_Workshop\"&fil.Name&"Updated.txt", ForWriting)

Do Until RS.AtEndOfStream
SLine = RS.ReadLine
WScript.Echo (SLine)
WS.Write Trim(SLine)
Loop
Next

RS.Close
WS.Close

Set SLine = Nothing
Set WSO = Nothing
Set RSO = Nothing
Set WS = Nothing
Set RS = Nothing
Set fil = Nothing
Set FLD = Nothing
