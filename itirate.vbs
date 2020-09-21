Option Explicit

Dim Reinitiate, Rex, DateF
Dim OptionH

Do While Reinitiate = 00

OptionH = InputBox ("Please enter respective number of one of the below option: " & vbCrLf & vbCrLf & "1. Initiate Recharge Recon" & vbCrLf & vbCrLf & "2. Initiate BillPay Recon" & vbCrLf & vbCrLf & "3. Recharge TTUM Upload" & vbCrLf & vbCrLf & "4. BillPay TTUM Upload" & vbCrLf & vbCrLf & "5. Recharge TTUM Generate" & vbCrLf & vbCrLf & "6. BillPay TTUM Generate" & vbCrLf & vbCrLf & "7. Recharge Tally" & vbCrLf & vbCrLf & "8. BillPay Tally" & vbCrLf & vbCrLf & "9. General Inquiry" & vbCrLf & vbCrLf &  "0. Quit","HomeScreen")

If OptionH = "1" Then


On Error Resume Next

Dim FSO, TS, FLD, strFolder, Fil
Dim intRow, ExcelObject1, ROWC, DB, ORS, CMD
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 
 

Set FSO = CreateObject("Scripting.FileSystemObject")
strFolder = "C:\Users\PR172959\Documents\Testing\Recharge"

		Set DB = WScript.CreateObject("ADODB.Connection")
		Set ORS = CreateObject("ADODB.Recordset")
		Set CMD = CreateObject("ADODB.Command")
		
	DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\PR172959\Documents\Pritimay\Database.accdb;"
	
	With CMD
	.ActiveConnection = DB
	.CommandText = "Delete from Recharge_Status_Temp"
	End With
	CMD.Execute
	WScript.Echo "BillPay_Status_Temp Data Cleared"

	With CMD
	.ActiveConnection = DB
	.CommandText = "Delete from REC_Sum"
	End With
	CMD.Execute
	WScript.Echo "REC_Sum Data Cleared"
	Set CMD = Nothing
	
	
Set ExcelObject1 = CreateObject("Excel.Application")
	set FLD = FSO.GetFolder(strFolder)
	
intRow = 2
ORS.Open "Recharge_Status_Temp", DB, adOpenStatic, adLockOptimistic

	For Each Fil In FLD.Files
	Call ExcelObject1.Workbooks.Open(fil.Path, ForReading)
	Set TS = ExcelObject1.ActiveWorkbook.Worksheets(1)
	ROWC = TS.UsedRange.Rows.Count
				
Do Until intRow > ROWC

If TS.Cells(intRow, 7).Text <> "" Then
	ORS.AddNew
	ORS("Cust_Id") = TS.Cells(intRow, 1).Text
	ORS("Reference") = TS.Cells(intRow, 2).Text
	ORS("Amount") = TS.Cells(intRow, 3).Value
	ORS("RRN") = TS.Cells(intRow, 5).Text
	ORS("Status") = TS.Cells(intRow, 7).Text
	ORS.Update 
End If
intRow = intRow +1		
Loop
Next
ORS.Close
WScript.Echo ("File Read Success")


	Set CMD = CreateObject("ADODB.Command")
	With CMD
	.ActiveConnection = DB
	.CommandText = "INSERT INTO REC_Sum SELECT Recharge_Status.RRN AS RRN, Recharge_Status.Amount AS Amount, Recharge_Status_Temp.Status AS Status FROM Recharge_Status INNER JOIN Recharge_Status_Temp ON (Left(Recharge_Status.RRN,9) = Left(Recharge_Status_Temp.RRN,9)) AND (Recharge_Status.Amount = Recharge_Status_Temp.Amount) AND (Recharge_Status.Cust_Id = Recharge_Status_Temp.Cust_Id);"
	End With
	CMD.Execute
	WScript.Echo "Status Uploaded in Temporary Table"
	Set CMD = Nothing
	
	Dim RS, strVar
	strVar = "**********************************************************************************" & vbCrLf & vbCrLf
	Set RS = DB.Execute("select Status, Count(*), Sum(Amount) from REC_Sum Group By Status Order by count(*)")
	DO WHILE NOT RS.EOF
	strVar = strVar & RS.Fields(0) & " --- " & RS.Fields(1)& " --- " & RS.Fields(2) &vbCrLf & vbCrLf
	RS.MoveNext
	Loop
	MsgBox (strVar)
	Set RS = Nothing
	Set strVar = Nothing
	
	

ExcelObject1.Quit
Set FLD = Nothing
Set strFolder = Nothing
Set Fil = Nothing
Set ORS = Nothing
Set TS = Nothing
Set WS = Nothing
Set FSO = Nothing
Set ExcelObject1 = Nothing
Set intRow = Nothing
Set ROWC = Nothing
DB.Close
Set DB = Nothing





ElseIf OptionH = "2" Then


Set FSO = CreateObject("Scripting.FileSystemObject")
strFolder = "C:\Users\PR172959\Documents\Testing\BillPay"

		Set DB = WScript.CreateObject("ADODB.Connection")
		Set ORS = CreateObject("ADODB.Recordset")
		Set CMD = CreateObject("ADODB.Command")
		
	DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\PR172959\Documents\Pritimay\Database.accdb;"
	
	With CMD
	.ActiveConnection = DB
	.CommandText = "Delete from BillPay_Status_Temp"
	End With
	CMD.Execute
	WScript.Echo "BillPay_Status_Temp Data Cleared"

	With CMD
	.ActiveConnection = DB
	.CommandText = "Delete from BP_Sum"
	End With
	CMD.Execute
	WScript.Echo "BP_Sum Data Cleared"
	Set CMD = Nothing
	
	
Set ExcelObject1 = CreateObject("Excel.Application")
	set FLD = FSO.GetFolder(strFolder)
	
intRow = 2
ORS.Open "BillPay_Status_Temp", DB, adOpenStatic, adLockOptimistic

	For Each Fil In FLD.Files
	Call ExcelObject1.Workbooks.Open(fil.Path, ForReading)
	Set TS = ExcelObject1.ActiveWorkbook.Worksheets(1)
	ROWC = TS.UsedRange.Rows.Count
				
Do Until intRow > ROWC

If Left(TS.Cells(intRow, 7).Text, 03) = "BB0" Then

	ORS.AddNew
	ORS("Cust_Id") = TS.Cells(intRow, 1).Text
	ORS("Reference") = TS.Cells(intRow, 2).Text
	ORS("Amount") = TS.Cells(intRow, 3).Value
	ORS("RRN") = TS.Cells(intRow, 5).Text
	ORS("Status") = TS.Cells(intRow, 8)
	ORS.Update 

ElseIf TS.Cells(intRow, 1).Text <> "" Then

	ORS.AddNew
	ORS("Cust_Id") = TS.Cells(intRow, 1).Text
	ORS("Reference") = TS.Cells(intRow, 2).Text
	ORS("Amount") = TS.Cells(intRow, 3).Value
	ORS("RRN") = TS.Cells(intRow, 5).Text
	ORS("Status") = TS.Cells(intRow, 7).Text
	ORS.Update 

End If
intRow = intRow +1

Loop
Next
ORS.Close
WScript.Echo ("File Read Success")

	Set CMD = CreateObject("ADODB.Command")
	With CMD
	.ActiveConnection = DB
	.CommandText = "INSERT INTO BP_Sum SELECT BillPay_Status.RRN AS RRN, BillPay_Status.Amount AS Amount, BillPay_Status_Temp.Status AS Status FROM BillPay_Status INNER JOIN BillPay_Status_Temp ON (Left(BillPay_Status.RRN,9) = Left(BillPay_Status_Temp.RRN,9)) AND (BillPay_Status.Amount = BillPay_Status_Temp.Amount) AND (BillPay_Status.Cust_Id = Billpay_Status_Temp.Cust_Id);"
	End With
	CMD.Execute
	WScript.Echo "Status Uploaded in Temporary Table"
	Set CMD = Nothing
	
	'Dim RS, strVar
	strVar = "**********************************************************************************" & vbCrLf & vbCrLf
	Set RS = DB.Execute("select Status, Count(*), Sum(Amount) from BP_Sum Group By Status Order by count(*)")
	DO WHILE NOT RS.EOF
	strVar = strVar & RS.Fields(0) & " --- " & RS.Fields(1)& " --- " & RS.Fields(2) &vbCrLf & vbCrLf
	RS.MoveNext
	Loop
	MsgBox (strVar)
	Set RS = Nothing
	Set strVar = Nothing


ExcelObject1.Quit
Set ORS = Nothing
Set TS = Nothing
Set FSO = Nothing
Set ExcelObject1 = Nothing
Set intRow = Nothing
Set ROWC = Nothing
DB.Close
Set DB = Nothing










ElseIf OptionH = "0" Then
WScript.Quit
End If








Reinitiate = inputbox ("Please enter 00 to Go To Home Screen")
Loop