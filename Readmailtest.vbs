Option Explicit
On Error Resume Next

Dim WSO, WS, RS, ObjOutlook, SSession, Item1, Inbox, OSub, OSender, DateC, TDate, Rlin, RLine

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

DateC = Date()

Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")
Set Inbox = SSession.GetDefaultFolder(6).Folders("Amalgamation")


For Each Item1 in Inbox.Items

OSub = Item1.Subject
OSender = Item1.SenderEmailAddress
OSub = Trim(Replace(OSub," ",""))
TDate = Left(Item1.ReceivedTime,10)

If OSender = "mobilebanking@VIJAYABANK.co.in" AND ((UCase(OSub)&TDate = "MOBILEBANKING-REPORT"&DateC) OR (Right(UCase(OSub),6)&TDate = "REPORT"&DateC)) Then
WScript.Echo(Item1.Body)
Set WSO = CreateObject("Scripting.FileSystemObject")
Set WS = WSO.CreateTextFile("E:\DashBoard\eVB\Read_Data.txt", ForWriting, True)
WS.Write(Item1.Body)
WS.Close
End If
Next

Set OSub = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing
Set OSender = Nothing