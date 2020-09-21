Option Explicit
On Error Resume Next

Dim RSO, WSO, TSO, TS, ObjOutlook, RS, WS, SLine, arr, SSession
Dim DB, ORS, CMD, strFolder, FLD, fil, Line, strVar

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

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
	Line = RS.ReadLine

	Line = Replace(Line,chr(34),"")
	arr = Split(Line, vbTab)
		
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
Set Line = Nothing
Set FLD = Nothing
Set WSO = Nothing



Set WSO = CreateObject("Scripting.FileSystemObject")
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForWriting)
Set RS = DB.Execute("SELECT BMaster.ZONE_NAME, BMaster.REGION_NAME, RBNA.SOL_ID, RBNA.CUSTOMER_ID, RBNA.CUSTOMER_NAME, RBNA.MOBILE_NUMBER, RBNA.REGISTRATION_DATE, RBNA.REGISTRATION_CHANNEL, BMaster.BRANCH_NAME FROM BMaster, RBNA WHERE BMaster.BANK_NAME = 'BANK OF BARODA' AND BMaster.SOL_ID = RBNA.SOL_ID ORDER BY BMaster.ZONE_NAME, BMaster.REGION_NAME")

DO WHILE NOT RS.EOF
strVar = RS.Fields(0) & "|" & RS.Fields(1) & "|" & RS.Fields(2) & "|" & RS.Fields(3) & "|" & RS.Fields(4) & "|" & RS.Fields(5) & "|" & RS.Fields(6) & "|" & RS.Fields(7) & "|" & RS.Fields(8) & vbCrLf
WS.Write (strVar)
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
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.AHMEDABAD@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
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



Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "BARODA" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.BARODA@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Baroda Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing




Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "BENGALURU" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.BENGALURU@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Bengaluru Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "BHOPAL" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.BHOPAL@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Bhopal Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,10) = "CHANDIGARH" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "zdm.zochd@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Chandigarh Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,7) = "CHENNAI" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.CHENNAI@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for CHENNAI ZONE"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "ERNAKULAM" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "zdm.ekmz@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ernakulam Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "HYDERABAD" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.Hyderabad@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Hyderabad Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing




Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "JAIPUR" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.JAIPUR@bankofbaroda.com;IT.RZ@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Jaipur Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing




Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,7) = "KOLKATA" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.KOLKATA@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Kolkata Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing




Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,7) = "LUCKNOW" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.LUCKNOW@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Lucknow Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "MANGALURU" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "it.zomglr@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Mangaluru Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "MEERUT" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.BAREILLY@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "MUMBAI" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.MUMBAI@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Mumbai Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "NEW DELHI" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.NEWDELHI@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Delhi Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,5) = "PATNA" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.PATNA@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Patna Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,4) = "PUNE" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "ZDM.PUNE@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Pune Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing






Set RS = RSO.OpenTextFile("C:\Campaign_Mails\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Campaign_Mails\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|SOL ID|CUSTOMER ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "RAJKOT" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,"|")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.SendUsingAccount = "mobility@bankofbaroda.com"
			
			.To = "zdm.zorajkot@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Rajkot Zone"
			.Attachments.Add "C:\Campaign_Mails\Data.csv"
			.Attachments.Add "E:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set RSO = Nothing
Set SLine = Nothing
Set WSO = Nothing
Set WS = Nothing

Else
WScript.Echo ("Error in Module!!! Exiting Now...")
WScript.Sleep (2000)
WScript.Quit

End If