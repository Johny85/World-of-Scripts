Option Explicit
On Error Resume Next

Dim FSO, TS
Dim WS, intRow, USER, ExcelObject1, BOB_ID, AMT, ACCT, RRN, BANK, STATUS, ROWC, BBPS
Const ForReading = 1, ForWriting = 2, ForAppending = 8 

	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set WS = FSO.CreateTextFile("C:\Users\PR172959\Documents\Pritimay\Recharge_Status.csv", ForWriting)
	Set ExcelObject1 = CreateObject("Excel.Application")

	
	'Open the file to read
		Call ExcelObject1.Workbooks.Open("C:\Users\PR172959\Documents\Pritimay\Recharge.xlsx", ForReading)
		Set TS = ExcelObject1.ActiveWorkbook.Worksheets(1)

intRow = 2
ROWC = TS.UsedRange.Rows.Count
				
Do Until intRow > ROWC

			USER = TS.Cells(intRow, 1).Text
			BOB_ID = TS.Cells(intRow, 2).Text
			AMT = TS.Cells(intRow, 3).Value
			ACCT = TS.Cells(intRow, 4).Text
			RRN = TS.Cells(intRow, 5).Text
			BANK = TS.Cells(intRow, 6).Text
			STATUS = TS.Cells(intRow, 7).Text
			'WSCript.Echo 
		
If STATUS <> "" Then
WS.Write (USER) & "|" & (BOB_ID) & "|" & (AMT) & "|" & (ACCT) & "|" & (RRN) & "|" & (BANK) & "|" & (STATUS) & vbCrLf
End If
			
				intRow = intRow +1		
    
Loop 

ExcelObject1.Quit
Set TS = Nothing
Set WS = Nothing
Set FSO = Nothing
Set ExcelObject1 = Nothing
Set intRow = Nothing
Set ROWC = Nothing
Set USER = Nothing
Set BOB_ID = Nothing
Set AMT = Nothing
Set ACCT = Nothing
Set RRN = Nothing
Set BANK = Nothing
Set STATUS = Nothing

