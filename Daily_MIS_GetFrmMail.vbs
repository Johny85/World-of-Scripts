Option Explicit
On Error Resume Next

Const DeleteReadOnly = True
'Dim ObjOutlook, SSession, Inbox, OSub, Item1, IntC, Atchm, OFrm, ODate
'Dim RSO, RS, WS, WSO, WSP, DateF, DateC
'Dim arr, SLine

Dim DateF, DateC, ObjOutlook, SSession, Item1, Atchm, Inbox, OSub, OFrm, ODate, IntC

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

DateF = Date()-1
DateC = Date()


Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")
Set Inbox = SSession.GetDefaultFolder(6).Folders("BOB Internal Mail")

For Each Item1 in Inbox.Items

OSub = Item1.Subject
OFrm = Item1.Sender.GetExchangeUser().PrimarySmtpAddress
ODate = Left(Item1.ReceivedTime,10)

If OFrm&ODate  = "mconnect.ito@bankofbaroda.com"&DateC Then
IntC = Item1.Attachments.Count
If IntC > 0 Then
For Each Atchm In Item1.Attachments
If Right(Atchm.FileName,12) = Replace(DateF,"-","")&".xls" OR Right(Atchm.FileName,13) = Replace(DateF,"-","")&".xlsx" Then
Atchm.SaveAsFile "E:\Mconnect Plus\Dashboard\MIS.xls"
Set Atchm = Nothing
End If
Next
End If
End If
Next

Set OSub = Nothing
Set OFrm = Nothing
Set ODate = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing
Set DateF = Nothing
Set DateC = Nothing