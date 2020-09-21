Option Explicit
On Error Resume Next

Dim WSO, WS, ObjOutlook, SSession, Item1, Inbox, OSub, OSender, DateC, TDate

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

DateC = Date()

'DateC = left(DateC,2)&"."&mid(DateC, 4,2)&"."&right(DateC,4)
WScript.Echo(DateC)

Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")
Set Inbox = SSession.GetDefaultFolder(6).Folders("Amalgamation")
Set WSO = CreateObject("Scripting.FileSystemObject")
For Each Item1 in Inbox.Items

OSub = Item1.Subject
OSender = Item1.SenderEmailAddress
'WScript.Echo(OSender)
OSub = Trim(Replace(OSub," ",""))
TDate = Left(Item1.ReceivedTime,10)

    If OSender = "mobilebanking@VIJAYABANK.co.in" AND UCase(OSub) = "MOBILEBANKING-REPORT" Then
    Set WS = WSO.CreateTextFile("E:\DashBoard\eVB\Read"&DateC&".txt", ForWriting, True)
    WS.Write(Item1.Body)
    'WScript.Echo(Item1.Body)
    End If
    WS.Close
    Set WS = Nothing
Next
Set WSO = Nothing
Set OSub = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing
Set DateC = Nothing
Set OSender = Nothing
Set TDate = Nothing
WScript.Echo("Mail Read Success")






Dim strFolder, RSO, FLD, Fil, RS, Line, ACount

	strFolder = "E:\DashBoard\eVB"

	'Create the filesystem object
	Set RSO = CreateObject("Scripting.FileSystemObject")
	'Set WSO = CreateObject("Scripting.FileSystemObject")

	'Set WS = WSO.CreateTextFile("E:\Project UPI Recon\Excel Files\MText.txt", ForWriting)
	Set FLD = RSO.GetFolder(strFolder)
	
	'loop through the folder and get the files

	For Each Fil In FLD.Files

			'Open the file to read

			Set RS = RSO.OpenTextFile(fil.Path, ForReading)
			'If not RS.AtEndOfStream Then RS.Skipline
			For ACount = 1 to 31
			RS.SkipLine
			Next
			Do Until RS.AtEndOfStream
			Line = RS.ReadLine
			WScript.Echo(Line)
			If Trim(Line) = "Non-Financial Transactions" Then
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			RS.Skipline
			WScript.Echo(RS.ReadLine)
			End If
			Loop 

		'Close the file
		RS.Close
		Next

'Clean up
Set TS = Nothing
Set FLD = Nothing
Set RSO = Nothing
Set RS = Nothing
Set strFolder = Nothing
Set Fil = Nothing
Set Line = Nothing