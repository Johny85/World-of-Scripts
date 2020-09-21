Dim RSO, RS, WS, WSO
Dim arr, SLine

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

Set WS = WSO.CreateTextFile("C:\MConnect\Reports\Bill_Payment_Ready.txt", ForWriting)
Set RS = RSO.OpenTextFile("C:\MConnect\Reports\Bil_Payment.txt", ForReading)
			
			If not RS.AtEndOfStream Then RS.Skipline
			Do Until RS.AtEndOfStream
			SLine = RS.ReadLine
			SLine = Replace(SLine,chr(34),"")
			'WScript.Echo (SLine)
	   		arr = split(SLine,"|")

'WScript.Echo arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &vbCrLf

WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|XXXXXXXXXXXXXX|"& arr(4) &"|"& arr(5) & vbCrLf	
			
Loop

		'Close the file
		RS.Close
	WS.Close
	
'Clean up
Set arr = Nothing
Set SLine = Nothing
Set RS = Nothing
Set WS = Nothing
Set RSO = Nothing
Set WSO = Nothing