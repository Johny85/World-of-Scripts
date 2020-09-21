Option Explicit
'On Error Resume Next

Dim RSO, WSO, RS, WSP, SLine, arr, FLD, strFolder, Fil, FSO
Const ForReading = 1, ForWriting = 2, ForAppending = 8

strFolder = "C:\Users\PR172959\Documents\Pritimay\Recharge"
Set FSO = CreateObject("Scripting.FileSystemObject")
Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

Set WSP = WSO.CreateTextFile("C:\MConnect\Reports\Recharge_DB.csv", ForWriting)
set FLD = FSO.GetFolder(strFolder)
			
For Each Fil In FLD.Files
Set RS = RSO.OpenTextFile(fil.Path, ForReading)
If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
SLine = RS.ReadLine
SLine = Replace(SLine,chr(34),"")
'WScript.Echo (SLine)
arr = split(SLine,"|")

WSP.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) & "|"& arr(4) &"|"& arr(5) &"|"& arr(6) & vbCrLf
			
Loop
Next
RS.Close

'Clean up
Set arr = Nothing
Set SLine = Nothing
Set RS = Nothing
Set WSP = Nothing
Set RSO = Nothing
Set WSO = Nothing
Set FLD = Nothing
Set FSO = Nothing
Set strFolder = Nothing



strFolder = "C:\Users\PR172959\Documents\Pritimay\BillPay"
Set FSO = CreateObject("Scripting.FileSystemObject")
Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")


Set WSP = WSO.CreateTextFile("C:\MConnect\Reports\Bill_Payment_DB.csv", ForWriting)
set FLD = FSO.GetFolder(strFolder)

For Each Fil In FLD.Files
Set RS = RSO.OpenTextFile(fil.Path, ForReading)
If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
SLine = RS.ReadLine
SLine = Replace(SLine,chr(34),"")
arr = split(SLine,"|")

WSP.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) & "|"& arr(4) &"|"& arr(5) & vbCrLf
			
Loop
Next
RS.Close

'Clean up
Set arr = Nothing
Set SLine = Nothing
Set RS = Nothing
Set WSP = Nothing
Set RSO = Nothing
Set WSO = Nothing
Set FLD = Nothing
Set FSO = Nothing
Set strFolder = Nothing
