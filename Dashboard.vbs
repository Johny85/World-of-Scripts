Option Explicit
On Error Resume Next

Const DeleteReadOnly = True
Dim DateF, EObj, TS, I, DB, CMD, RSO, RS, ORS, SLine
Dim RegC, ActC, FTC, FTV, xlApp, arr, TxnCnt, TxnAmt


Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3

DateF = Date() - 1
WScript.Echo (DateF)
FTV = 0

Set EObj = CreateObject("Excel.Application")
EObj.Workbooks.Open("C:\MConnect\Reports\User Registration.xlsx")
Set TS = EObj.ActiveWorkbook.Worksheets(1)
RegC = (TS.UsedRange.Rows.Count)-1
WScript.Echo ("Total BoB Registeration: "&RegC)


	TS.Close
	EObj.ActiveWorkbook.Save
	EObj.ActiveWorkbook.Close
	EObj.Application.Quit

'Clean up
'Set DateF = Nothing
Set EObj = Nothing
Set TS = Nothing

Set EObj = CreateObject("Excel.Application")
EObj.Workbooks.Open("C:\MConnect\Reports\User Activation.xlsx")
Set TS = EObj.ActiveWorkbook.Worksheets(1)
ActC = (TS.UsedRange.Rows.Count)-1
WScript.Echo ("Total BoB Activation: "&ActC)

	TS.Close
	EObj.ActiveWorkbook.Save
	EObj.ActiveWorkbook.Close
	EObj.Application.Quit

'Clean up
'Set DateF = Nothing
Set EObj = Nothing
Set TS = Nothing





Set xlApp = CreateObject("Excel.Application")
xlApp.Workbooks.Open("C:\MConnect\Reports\Successful transactions.xlsx")
xlApp.Workbooks(1).SaveAs "E:\DashBoard\SUCC_"&DateF&".csv",6

xlApp.DisplayAlerts = False
xlApp.ActiveWorkbook.Save = False
xlApp.ActiveWorkbook.Close
xlApp.Application.Quit


Set DB = CreateObject("ADODB.Connection")
Set CMD = CreateObject("ADODB.Command")

DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\DashBoard\Dashboard.accdb;"

'WScript.Echo ("DB Connection Success")
	
With CMD
.ActiveConnection = DB
.CommandText = "Delete from Succ_Txn_Temp"
End With
CMD.Execute
Set CMD = Nothing
'WScript.Echo ("Previous DB Data Deleted")








Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "Succ_Txn_Temp", DB, adOpenStatic, adLockOptimistic

Set RSO = CreateObject("Scripting.FileSystemObject")
Set RS = RSO.OpenTextFile("E:\DashBoard\SUCC_"&DateF&".csv")
	
	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	
	SLine = RS.ReadLine
	'WScript.Echo(SLine)
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
		
	ORS.AddNew
	ORS("SERVICE_CODE") = arr(0)
	ORS("CUSTOMER_ID") = arr(1)
	ORS("REQUEST_DATE") = Left(arr(2),10)
	ORS("ACCOUNT_NUMBER") = arr(3)
	ORS("AMOUNT") = arr(4)
	ORS("REFERENCE_NUMBER") = arr(5)
	ORS("STATUS") = arr(6)
	ORS.Update
	
	Loop

RS.Close
ORS.Close
DB.Close
'WScript.Echo ("SUCCESSFUL TRANSACTION FILE UPLOADED TO DB")


Set RS = Nothing
Set ORS = Nothing
Set arr = Nothing
Set SLine = Nothing
Set RSO = Nothing


DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\DashBoard\Dashboard.accdb;"
Set RS = DB.Execute("SELECT Count(*), Sum(AMOUNT) FROM Succ_Txn_Temp")

'TxnCnt = RS.Fields(0)
'TxnAmt = RS.Fields(1)

WScript.Echo("Total Successful Transaction Count: "& Round(RS.Fields(0)/100000,2) &" Lakhs")
WScript.Echo("Total Successful Transaction Amount: "& Round(RS.Fields(1)/10000000,2)&" Crores")

DB.Close

Set RS = Nothing
Set DB = Nothing




