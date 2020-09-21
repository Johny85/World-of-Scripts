Option Explicit

Dim RSO, WSO, WSP, RS, SLine, arr

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

Set WSP = WSO.CreateTextFile("C:\MConnect\Reports\Recharge_DB.csv", ForWriting)
Set RS = RSO.OpenTextFile("C:\MConnect\Reports\Recharge.txt", ForReading)
			
			If not RS.AtEndOfStream Then RS.Skipline
			Do Until RS.AtEndOfStream
			SLine = RS.ReadLine
			SLine = Replace(SLine,chr(34),"")
			'WScript.Echo (SLine)
	   		arr = split(SLine,"|")

WSP.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) & "|"& arr(4) &"|"& arr(5) &"|"& arr(6) & vbCrLf
			
Loop

	RS.Close
	WSP.Close
	
'Clean up
Set arr = Nothing
Set SLine = Nothing
Set RS = Nothing
Set WSP = Nothing
Set RSO = Nothing
Set WSO = Nothing

WScript.Echo ("Recharge File processing completed")

Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

Set WSP = WSO.CreateTextFile("C:\MConnect\Reports\Bill_Payment_DB.csv", ForWriting)
Set RS = RSO.OpenTextFile("C:\MConnect\Reports\Bill Payment.txt", ForReading)
			
			If not RS.AtEndOfStream Then RS.Skipline
			Do Until RS.AtEndOfStream
			SLine = RS.ReadLine
			SLine = Replace(SLine,chr(34),"")
	   		arr = split(SLine,"|")

WSP.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) & "|"& arr(4) &"|"& arr(5) & vbCrLf

Loop

	RS.Close
	WSP.Close
	
'Clean up
Set arr = Nothing
Set SLine = Nothing
Set RS = Nothing
Set WSP = Nothing
Set RSO = Nothing
Set WSO = Nothing
