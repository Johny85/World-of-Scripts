Option Explicit
'On Error Resume Next

Dim ToRead, RSO, WSO, WSP, RS, SLine, arr

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

ToRead = inputbox ("Please Enter Region Name")

Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")
Set WSP = WSO.CreateTextFile("C:\Users\PR172959\Downloads\"&ToRead&".csv", ForWriting)
Set RS = RSO.OpenTextFile("C:\Users\PR172959\Downloads\Region Wise Data\"&ToRead&".csv", ForReading)


If not RS.AtEndOfStream Then RS.Skipline
			Do Until RS.AtEndOfStream
			SLine = RS.ReadLine
			SLine = Replace(SLine,chr(34),"")
			'WScript.Echo (SLine)
	   		arr = split(SLine,"|")

WSP.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(7)&"|"& arr(8)&"|"& arr(10)&"|"& arr(15)&"|"& arr(19) & vbCrLf

Loop


	RS.Close
	WS.Close
	WSP.Close
	
	
Set SLine = Nothing
Set RS = Nothing
Set WSP = Nothing
Set RSO = Nothing
Set WSO = Nothing
