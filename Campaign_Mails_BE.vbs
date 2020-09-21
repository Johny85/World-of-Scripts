Option Explicit
On Error Resume Next

Dim objFSO, WS, WSO, DateF, ObjOutlook, SSession, Item1, ATchm, Inbox, OSub
Dim DB, IntC, objShl

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

WScript.Echo ("Module Starts..")
 
DateF = Date()-1 
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists("C:\Campaign_Mails\Registered.zip") Then
objFSO.DeleteFile("C:\Campaign_Mails\Registered.zip")
End If
objFSO.DeleteFile("C:\Campaign_Mails\RAW\*")

Set objFSO = Nothing

WScript.Echo ("Previous Files Deleted")

Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")
Set Inbox = SSession.GetDefaultFolder(6).Folders("BOB Internal Mail")

WScript.Echo ("Checking Mailbox for Data Files")

For Each Item1 in Inbox.Items

OSub = Item1.Subject
OSub = Trim(Replace(OSub," ",""))

	If UCase(OSub) = "REGISTEREDONLYUSERS|"&DateF Then
	
	IntC = Item1.Attachments.Count
		If IntC > 0 Then
		For Each Atchm In Item1.Attachments
		Atchm.SaveAsFile "C:\Campaign_Mails\Registered.zip"
		Set Atchm = Nothing
		Next
		End If
		End If

Next
Set Atchm = Nothing
Set OSub = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing

WScript.Echo ("File Downloaded from Mailbox")

Set objShl = WScript.CreateObject ("WScript.shell")
objShl.run """C:\Program Files\WinRAR\WinRAR.exe"" X ""C:\Campaign_Mails\Registered.zip"" ""C:\Campaign_Mails\RAW"""
Set objShl = Nothing
