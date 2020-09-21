Option Explicit
On Error Resume Next

Dim DateC, DateF, objOutlook, SSession, Item1, Atchm, Inbox, OFrm, OSub, TDate, WSO, WS, RSO, RS, Line

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

DateC = Date()-1
DateF = Date()

Set WSO = CreateObject("Scripting.FileSystemObject")
Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")
Set Inbox = SSession.GetDefaultFolder(6).Folders("BOB Internal Mail")


For Each Item1 in Inbox.Items
OFrm = Item1.Sender.GetExchangeUser().PrimarySmtpAddress
OSub = Item1.Subject
TDate = Left(Item1.ReceivedTime,10)

	If OFrm&TDate&UCase(Left(OSub, 23))  = "mobilebanking.evb@bankofbaroda.com"&DateF&"MOBILE BANKING - REPORT" Then
	Set WS = WSO.CreateTextFile("C:\Users\PR172959\Desktop\DR\eVB_Temp.txt", ForWriting)
	WS.Write(Item1.body)
	End If

Next
Set OSub = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing
Set WSO = Nothing
Set WS = Nothing
WScript.Echo("Mail Read Success")


Set WSO  = CreateObject("Scripting.FileSystemObject")
Set RSO  = CreateObject("Scripting.FileSystemObject")
Set RS  = RSO.OpenTextFile("C:\Users\PR172959\Desktop\DR\eVB_Temp.txt", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Desktop\DR\eVB.txt", ForWriting)

			If not RS.AtEndOfStream Then RS.Skipline
			Do Until RS.AtEndOfStream
			Line = RS.ReadLine
			If Len(Line)>0 Then
			WS.Write (Line) & vbCrLf
			End If
			Loop

	RS.Close
	WS.Close
	
Set WSO = Nothing
Set RSO = Nothing
Set RS = Nothing
Set WS = Nothing
Set Line = Nothing
