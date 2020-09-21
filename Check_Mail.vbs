Option Explicit
On Error Resume Next

Const DeleteReadOnly = True
Dim ObjOutlook, SSession, Inbox, OSub, Item1, IntC, Atchm, OFrm, ODate
Dim RSO, RS, WS, WSO, WSP, DateF, DateC
Dim arr, SLine
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
Set Inbox = SSession.GetDefaultFolder(6).Folders("Amalgamation")

For Each Item1 in Inbox.Items

OSub = Item1.Subject
OFrm = Item1.SenderEmailAddress
ODate = Left(Item1.ReceivedTime,10)

If OFrm&ODate  = "denaiconnect@denabank.co.in"&DateC Then

IntC = Item1.Attachments.Count
If IntC > 0 Then
For Each Atchm In Item1.Attachments
If Mid(Atchm.FileName,8,15) = "Data_"&DateF Then
Atchm.SaveAsFile "E:\Mconnect Plus\Dashboard\eDB.xlsx"
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



Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")
Set Inbox = SSession.GetDefaultFolder(6).Folders("Amalgamation")

For Each Item1 in Inbox.Items

OSub = Item1.Subject
OFrm = Item1.SenderEmailAddress
ODate = Left(Item1.ReceivedTime,10)


If OFrm&ODate = "mobilebanking@VIJAYABANK.co.in"&DateC Then

IntC = Item1.Attachments.Count
If IntC > 0 Then
For Each Atchm In Item1.Attachments
If Left(Atchm.FileName,10) = "Activation" Then
'WScript.Echo (Atchm.FileName)
Atchm.SaveAsFile "E:\Mconnect Plus\Dashboard\eVB.xls"
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


