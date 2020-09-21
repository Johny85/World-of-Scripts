Option Explicit
'On Error Resume Next

Dim ObjOutlook, SSession, Inbox, OSub, IntC
Dim DateF
Dim Item1, Atchm
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")

DateF = Date()-1


Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Inbox = SSession.GetDefaultFolder(6).Folders("BOB Internal Mail")

For Each Item1 in Inbox.Items

OSub = Item1.Subject

	If Left(OSub,7) = "Reports" AND Right(OSub,11) = " "&(DateF) Then
	IntC = Item1.Attachments.Count
			If IntC > 0 Then
			For Each Atchm In Item1.Attachments
			If Left(Atchm.FileName,7) = "Reports" Then
			Atchm.SaveAsFile "C:\MConnect\" & Atchm.FileName
			Set Atchm = Nothing
			End If
			Next
		End If
		End If
Next
Set Item1 = Nothing
Set Atchm = Nothingk
Set OSub = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing


' 