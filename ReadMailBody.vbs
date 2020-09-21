Option Explicit
On Error Resume Next

Dim WSO, RSO, WS, RS, ObjOutlook, SSession, Item1, Inbox, OSub, TDate, DateC, SLine, arr, xlApp, TS, TotalDown

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

DateC = Date()-1
TotalDown = 0

Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")
Set Inbox = SSession.GetDefaultFolder(6).Folders("Service Providers")

Set WSO = CreateObject("Scripting.FileSystemObject")
WScript.Echo("Text File Created")
For Each Item1 in Inbox.Items

OSub = Item1.Subject
OSub = Trim(Replace(OSub," ",""))
TDate = Left(Item1.ReceivedTime,10)

	If Left(UCase(OSub),9)&TDate = "BOB(IMPS)"&DateC Then
	Set WS = WSO.CreateTextFile("E:\Mconnect Plus\IMPS_Downtime\"&DateC&".txt", ForWriting)
	WS.Write(Item1.body)
	ElseIf Left(UCase(OSub),12)&TDate = "RE:BOB(IMPS)"&DateC Then
	Set WS = WSO.CreateTextFile("E:\Mconnect Plus\IMPS_Downtime\"&DateC&".txt", ForWriting)
	WS.Write(Item1.body)	
	End If
	Set WS = Nothing
Next
Set OSub = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing
Set WSO = Nothing
Set WS = Nothing
WScript.Echo("Mail Read Success")

Set xlApp = CreateObject("Excel.Application")
xlApp.Workbooks.Open("E:\Mconnect Plus\MIS\"&DateC&".xlsx")
Set RSO = CreateObject("Scripting.FileSystemObject")
Set TS = xlApp.ActiveWorkbook.Worksheets("Sample")
Set FSO = CreateObject("Scripting.FileSystemObject")

WScript.Echo("Variable Initialized")

If FSO.FileExists("E:\Mconnect Plus\IMPS_Downtime\"&DateC&".txt") Then
Set RS = RSO.OpenTextFile("E:\Mconnect Plus\IMPS_Downtime\"&DateC&".txt", ForReading)
			Do Until RS.AtEndOfStream
			SLine = RS.ReadLine
			If Mid(SLine,4,4) = "Mins" OR Mid(SLine,5,4) = "Mins" Then
			TotalDown = TotalDown + Left(SLine,3)*1
			End If
			Loop
			WScript.Echo(TotalDown)
			TS.Cells(37,9).Value = "Decline Duration time "&TotalDown&" Minutes"
Else
TS.Cells(37,9).Value = "Decline Duration time 00 Minutes"
End If
WScript.Echo("Downtime Updated")

Set RSO = Nothing
Set RS = Nothing
Set SLine = Nothing

Set TS = Nothing
xlApp.ActiveWorkbook.Save
xlApp.ActiveWorkbook.Close
xlApp.Application.Quit

Set xlApp = Nothing
