Option Explicit
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
Set ORS = Nothing
Set TS = Nothing
Set WS = Nothing
Set FSO = Nothing
Set ExcelObject1 = Nothing
Set intRow = Nothing
Set ROWC = Nothing
DB.Close
Set DB = Nothing




