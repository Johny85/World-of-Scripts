Option Explicit
'On Error Resume Next

'Dim FSO, FLD, FIL, TS
'Dim strFolder, WS, intRow, Txn_Type, ExcelObject1, File_Type, File_Date, Description, Count, Debit_Amount, Credit_Amount
Dim FSO, TS, WS, ExcelObject1, RowC 
Const ForReading = 1, ForWriting = 2, ForAppending = 8 


	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set WS = FSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\NTSL.txt", ForWriting)
	Set ExcelObject1 = CreateObject("Excel.Application")

	
Call ExcelObject1.Workbooks.Open("C:\Users\PR172959\Documents\Pritimay\Recharge.xlsx", ForReading)
Set TS = ExcelObject1.ActiveWorkbook.Worksheets(1)
RowC = TS.UsedRange.Rows.Count
WScript.Echo (RowC)

Set FSO = Nothing
Set WS = Nothing
Set TS = Nothing
Set ExcelObject1 = Nothing
Set RowC = Nothing