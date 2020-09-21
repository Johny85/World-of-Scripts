Option Explicit
On Error Resume Next

Const DeleteReadOnly = True
Dim objShl, DateF, objFSO
Dim ObjOutlook, SSession, Inbox, OSub, Item1, IntC, Atchm

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8


DateF = Date()-1
'WScript.Echo (DateF)

Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = ObjOutlook.MailItem
Set Atchm = ObjOutlook.Attachment
Set Inbox = SSession.GetDefaultFolder(6).Folders("BOB Internal Mail")

For Each Item1 in Inbox.Items

'If Item1.UnRead=True Then
OSub = Item1.Subject

	If Left(OSub,7) = "Reports" AND Right(OSub,11) = (DateF)&" " Then
	Wscript.Echo (OSub)
	'IntC = Item1.Attachments.Count
		'If IntC > 0 Then
			'For Each Atchm In Item1.Attachments 
			'Atchm.SaveAsFile "C:\MConnect\" & Atchm.FileName
			'Set Atchm = Nothing
			'Next
		'End If
		'Item1.UnRead = False
	'End If
End If

Next
Set OSub = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing

