Option Explicit
On Error Resume Next

Dim RSO, WSO, WS, RS, Line, Srchs

Const ForReading = 1, ForWriting = 2, ForAppending = 8 

	Set RSO = CreateObject("Scripting.FileSystemObject")
	'Set WSO = CreateObject("Scripting.FileSystemObject")

	'Set WS = WSO.CreateTextFile("E:\DashBoard\eVB\Write"&DateC&".txt", ForWriting)
	Set RS = RSO.OpenTextFile("E:\DashBoard\eVB\Read"&DateC&".txt", ForReading)
			
	Do Until RS.AtEndOfStream

		Line = RS.ReadLine
		Srchs = Left(Line,30)
			
			If Trim(Srchs) = "Non-Financial Transactions" Then
			RS.Skipline
			WScript.Echo(RS.Readline)
			RS.Skipline
			WScript.Echo(RS.Readline)
			RS.Skipline
			WScript.Echo(RS.Readline)
			RS.Skipline
			WScript.Echo(RS.Readline)
			RS.Skipline
			WScript.Echo(RS.Readline)
			RS.Skipline
			WScript.Echo(RS.Readline)
			RS.Skipline
			WScript.Echo(RS.Readline)
			RS.Skipline
			WScript.Echo(RS.Readline)
			RS.Skipline
			WScript.Echo(RS.Readline)
			RS.Skipline
			WScript.Echo(RS.Readline)
			RS.Skipline
			WScript.Echo(RS.Readline)
			End If
	Loop 

	RS.Close
	'WS.Close
	
'Clean up
'Set WS = Nothing
Set RS = Nothing
'Set WSO = Nothing
Set RSO = Nothing
Set Line = Nothing
Set Srchs = Nothing