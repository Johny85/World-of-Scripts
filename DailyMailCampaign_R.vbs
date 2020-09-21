Option Explicit
'On Error Resume Next

Dim DateF, RS, WS, RSO, WSO, SLine, arr, ObjOutlook, SSession

Const ForReading = 1, ForWriting = 2, ForAppending = 8 
DateF = Date()-1

Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

Set ObjOutlook = CreateObject("Outlook.Application")
Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "AHMEDABAD" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop
	
	Set RS = Nothing
	Set SLine = Nothing
	Set WS = Nothing
	Set arr = Nothing
			

Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,16,15) = "AHMEDABAD - I R" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "IT.Ahmedabadcity@bankofbaroda.com"
			.Cc = "ZDM.AHMEDABAD@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing

Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,16,16) = "AHMEDABAD - II R" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "it.ahmedabad2@bankofbaroda.com"
			.Cc = "ZDM.AHMEDABAD@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,16,7) = "ANAND R" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "IT.Anand@bankofbaroda.com"
			.Cc = "ZDM.AHMEDABAD@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,16,7) = "BANAS K" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ecm.banaskantha@bankofbaroda.com"
			.Cc = "ZDM.AHMEDABAD@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,16,11) = "GANDHINAGAR" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "IT.GNGR@bankofbaroda.com"
			.Cc = "ZDM.AHMEDABAD@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,16,7) = "KHEDA R" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "IT.Kheda@bankofbaroda.com"
			.Cc = "ZDM.AHMEDABAD@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,16,7) = "MEHSANA" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "IT.Mehsana@bankofbaroda.com"
			.Cc = "ZDM.AHMEDABAD@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,16,7) = "SABARKA" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "bobsabarkantha@bankofbaroda.com"
			.Cc = "ZDM.AHMEDABAD@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing




Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "BARODA" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

	Set RS = Nothing
	Set SLine = Nothing
	Set WS = Nothing
	Set arr = Nothing
			

Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,13,11) = "BARODA CITY" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "IT.BARODA@bankofbaroda.com"
			.Cc = "ZDM.BARODA@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,13,11) = "BARODA DIST" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "it.barodadistrict@bankofbaroda.com"
			.Cc = "ZDM.BARODA@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,13,11) = "BARODA DIST" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "it.barodadistrict@bankofbaroda.com"
			.Cc = "ZDM.BARODA@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,13,11) = "BHARUCH REG" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.BARODA@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking for Bharuch Region as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data_Region.csv", ForWriting)
WS.Write "ZONE NAME|REGION NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & 

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Mid(SLine,13,11) = "BARODA DIST" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) &"|"& arr(8) & vbCrLf
	End If
	Loop

			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "it.barodadistrict@bankofbaroda.com"
			.Cc = "ZDM.BARODA@bankofbaroda.com; KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing
Set arr = Nothing




































Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "BENGALURU" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.BENGALURU@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Bengaluru Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "BHOPAL" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.BHOPAL@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Bhopal Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,10) = "CHANDIGARH" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "zdm.zochd@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Chandigarh Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,7) = "CHENNAI" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.CHENNAI@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for CHENNAI ZONE"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "ERNAKULAM" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "zdm.ekmz@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Ernakulam Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "HYDERABAD" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.Hyderabad@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Hyderabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing




Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "JAIPUR" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.JAIPUR@bankofbaroda.com;IT.RZ@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Jaipur Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing




Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,7) = "KOLKATA" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.KOLKATA@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Kolkata Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing




Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,7) = "LUCKNOW" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.LUCKNOW@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Lucknow Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "MANGALURU" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "it.zomglr@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Mangaluru Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "MEERUT" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.BAREILLY@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "MUMBAI" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.MUMBAI@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Mumbai Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,9) = "NEW DELHI" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.NEWDELHI@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Delhi Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,5) = "PATNA" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.PATNA@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Patna Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing





Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,4) = "PUNE" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.PUNE@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Pune Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set SLine = Nothing
Set WS = Nothing






Set RS = RSO.OpenTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\RBNA_Res.csv", ForReading)
Set WS = WSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv", ForWriting)
WS.Write "ZONE NAME|CUSTOMER ID|SOL ID|CUSTOMER NAME|MOBILE NUMBER|REGISTRATION DATE|REGISTRATION CHANNEL|REGION NAME|BRANCH_NAME" & vbCrLf

	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	SLine = RS.ReadLine
	If Left(SLine,6) = "RAJKOT" Then
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
	WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &"|"& arr(6) &"|"& arr(7) & vbCrLf
	End If
	Loop
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "zdm.zorajkot@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "Registered but not active report of Mobile Banking for Rajkot Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Campaign\Data.csv"
			.Attachments.Add "D:\CSVtoExcel_tutorial.docx"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached registered but not activated data of mobile banking as on " & DateF& "." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing
Set RS = Nothing
Set RSO = Nothing
Set SLine = Nothing
Set WSO = Nothing
Set WS = Nothing