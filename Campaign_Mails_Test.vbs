Option Explicit
On Error Resume Next

Dim RS, FLD, strFolder, RSO, objFSO, WS, WSO, TSO, TS, DateF, DateC, STC
Dim DB, ORS, CMD, fil, NLine, arr, strvar, SLine
Dim ObjOutlook, SSession, Item1, Atchm, Inbox, OSub, IntC, objShl

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

WScript.Echo ("Module Starts")
 
DateF = Date()-2 
 
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


Set DB = CreateObject("ADODB.Connection")
Set CMD = CreateObject("ADODB.Command")
		
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Campaign_Mails\Campaign.accdb;"

WScript.Echo ("DB Connection Success")
	
With CMD
.ActiveConnection = DB
.CommandText = "Delete from RBNA"
End With
CMD.Execute
Set CMD = Nothing
DB.Close

WScript.Echo ("DB Data Cleared")
	
	strFolder = "C:\Campaign_Mails\RAW"

	Set RSO = CreateObject("Scripting.FileSystemObject")
	Set ORS = CreateObject("ADODB.Recordset")
	ORS.Open "RBNA", DB, adOpenStatic, adLockOptimistic
	
	Set FLD = RSO.GetFolder(strFolder)
	For Each fil In FLD.Files

	Set RS = RSO.OpenTextFile(fil.Path, ForReading)
	If not RS.AtEndOfStream Then RS.Skipline					
	
	Do Until RS.AtEndOfStream
	NLine = RS.ReadLine

	NLine = Replace(NLine,chr(34),"")
	arr = Split(NLine, vbTab)
	
	WScript.Echo (arr(0) & "|" & arr(1) & "|" & arr(2) & "|" & arr(3) & "|" & arr(4) & "|" & arr(5) & vbCrLf)	
	
	ORS.AddNew
		ORS("SOL_ID") = arr(0)
		ORS("CUSTOMER_ID") = arr(1)
		ORS("CUSTOMER_NAME") = arr(2)
		ORS("MOBILE_NUMBER") = arr(3)
		ORS("REGISTRATION_DATE") = arr(4)
		ORS("REGISTRATION_CHANNEL") = arr(5)
	ORS.Update 

	Loop
	Next
	
	ORS.Close
WScript.Echo ("Data Upload to DB")
	
Set ORS = Nothing
Set RS = Nothing
Set arr = Nothing
Set FSO = Nothing
Set RSO = Nothing
Set NLine = Nothing
Set FLD = Nothing



Set WSO = CreateObject("Scripting.FileSystemObject")
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForWriting)
Set RS = DB.Execute("SELECT BMaster.ZONE_NAME, BMaster.REGION_NAME, RBNA.SOL_ID, RBNA.CUSTOMER_ID, RBNA.CUSTOMER_NAME, RBNA.MOBILE_NUMBER, RBNA.REGISTRATION_DATE, RBNA.REGISTRATION_CHANNEL, BMaster.BRANCH_NAME FROM BMaster, RBNA WHERE BMaster.BANK_NAME = 'BANK OF BARODA' AND BMaster.SOL_ID = RBNA.SOL_ID ORDER BY BMaster.ZONE_NAME, BMaster.REGION_NAME")

DO WHILE NOT RS.EOF
strVar = RS.Fields(0) & "|" & RS.Fields(1) & "|" & RS.Fields(2) & "|" & RS.Fields(3) & "|" & RS.Fields(4) & "|" & RS.Fields(5) & "|" & RS.Fields(6) & "|" & RS.Fields(7) & "|" & RS.Fields(8) & vbCrLf
WS.Write (strvar)
RS.MoveNext
Loop

RS.Close
DB.Close

Set RS = Nothing
Set strVar = Nothing
Set WS = Nothing
Set WSO = Nothing
Set DB = Nothing


DateC =Date()
Set TSO = CreateObject("Scripting.FileSystemObject")
Set TS = TSO.GetFile("C:\Campaign_Mails\RBNA_Res.csv")
strVar = StrComp(DateC,Left(TS.DateLastModified,10))

If strVar = 0 Then

WScript.Echo ("Sending Mails to Respective Zones")

Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

Set ObjOutlook = CreateObject("Outlook.Application")
Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "AHMEDABAD" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "Mobility"
			
			.To = "mobility@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing

Else
WScript.Echo ("Date not Matching")
WScript.Sleep (2000)
WScript.Quit
End If

