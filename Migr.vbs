Option Explicit
On Error Resume Next

Dim filsys, DB, RS, xlApp, TS, i, FieldSOL, FieldSTAT, FieldCNT

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 12
Const adLockOptimistic = 12 

DateC = Date()-1

Set filsys = CreateObject("Scripting.FileSystemObject")
filsys.CopyFile "E:\Mconnect Plus\Migrated.xlsx", "E:\Mconnect Plus\Mig\"&DateC&".xlsx"
Set filsys = Nothing
WScript.Echo ("SAMPLE FILE COPIED")

Set xlApp = CreateObject("Excel.Application")
xlApp.Workbooks.Open("E:\Mconnect Plus\Migrated.xlsx")
Set TS = xlApp.ActiveWorkbook.Worksheets("Main")
i = 0

Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\DashBoard\Dashboard.accdb;"
WScript.Echo ("DB Connection Success")
Set RS = DB.Execute("SELECT SOL, Status, Count(*) FROM Migrated WHERE SOL in ('7001','71279','71294','71295','74012','7410','7411','7428','7469','6705','6709') and Date = '"&DateC&"' GROUP BY SOL, Status ORDER BY SOL")
WScript.Echo ("Query Execution Success")

DO WHILE NOT RS.EOF
FieldSOL = RS.Fields(0)
FieldSTAT = RS.Fields(1)
FieldCNT = RS.Fields(2)
WScript.Echo (RS.Fields(0) & "|" & RS.Fields(1) & "|" & RS.Fields(2))
If FieldSOL = "6705" Then
	If FieldSTAT = "A" Then
	WScript.Echo ("Active Users")
	TS.Cells(12,16).Value = ""&FieldCNT&""
	TS.Cells(12,5).Value = TS.Cells(12,5).Value + TS.Cells(12,16).Value
	TS.Cells(12,16).Value = "0"
	Else
	WScript.Echo ("Registered Users")
	TS.Cells(12,16).Value = ""&FieldCNT&""
	TS.Cells(12,4).Value = TS.Cells(12,4).Value + TS.Cells(12,16).Value
	TS.Cells(12,16).Value = "0"
	End If
ElseIf FieldSOL = "6709" Then
	If FieldSTAT = "A" Then
	TS.Cells(4,16).Value = ""&FieldCNT&""
	TS.Cells(4,5).Value = TS.Cells(4,5).Value + TS.Cells(4,16).Value
	TS.Cells(4,16).Value = "0"
	Else
	TS.Cells(4,16).Value = ""&FieldCNT&""
	TS.Cells(4,4).Value = TS.Cells(4,4).Value + TS.Cells(4,16).Value
	TS.Cells(4,16).Value = "0"
	End If
ElseIf FieldSOL = "7001" Then
	If FieldSTAT = "A" Then
	TS.Cells(5,16).Value = ""&FieldCNT&""
	TS.Cells(5,5).Value = TS.Cells(5,5).Value + TS.Cells(5,16).Value
	TS.Cells(5,16).Value = "0"
	Else
	TS.Cells(5,16).Value = ""&FieldCNT&""
	TS.Cells(5,4).Value = TS.Cells(5,4).Value + TS.Cells(5,16).Value
	TS.Cells(5,16).Value = "0"
	End If
ElseIf FieldSOL = "71279" Then
	If FieldSTAT = "A" Then
	TS.Cells(6,16).Value = ""&FieldCNT&""
	TS.Cells(6,5).Value = TS.Cells(6,5).Value + TS.Cells(6,16).Value
	TS.Cells(6,16).Value = "0"
	Else
	TS.Cells(6,16).Value = ""&FieldCNT&""
	TS.Cells(6,4).Value = TS.Cells(6,4).Value + TS.Cells(6,16).Value
	TS.Cells(6,16).Value = "0"
	End If
ElseIf FieldSOL = "71294" Then
	If FieldSTAT = "A" Then
	TS.Cells(7,16).Value = ""&FieldCNT&""
	TS.Cells(7,5).Value = TS.Cells(7,5).Value + TS.Cells(7,16).Value
	TS.Cells(7,16).Value = "0"
	Else
	TS.Cells(7,16).Value = ""&FieldCNT&""
	TS.Cells(7,4).Value = TS.Cells(7,4).Value + TS.Cells(7,16).Value
	TS.Cells(7,16).Value = "0"
	End If
ElseIf FieldSOL = "71295" Then
	If FieldSTAT = "A" Then
	TS.Cells(8,16).Value = ""&FieldCNT&""
	TS.Cells(8,5).Value = TS.Cells(8,5).Value + TS.Cells(8,16).Value
	TS.Cells(8,16).Value = "0"
	Else
	TS.Cells(8,16).Value = ""&FieldCNT&""
	TS.Cells(8,4).Value = TS.Cells(8,4).Value + TS.Cells(8,16).Value
	TS.Cells(8,16).Value = "0"
	End If
ElseIf FieldSOL = "74012" Then
	If FieldSTAT = "A" Then
	TS.Cells(9,16).Value = ""&FieldCNT&""
	TS.Cells(9,5).Value = TS.Cells(9,5).Value + TS.Cells(9,16).Value
	TS.Cells(9,16).Value = "0"
	Else
	TS.Cells(9,16).Value = ""&FieldCNT&""
	TS.Cells(9,4).Value = TS.Cells(9,4).Value + TS.Cells(9,16).Value
	TS.Cells(9,16).Value = "0"
	End If
ElseIf FieldSOL = "7410" Then
	If FieldSTAT = "A" Then
	TS.Cells(10,16).Value = ""&FieldCNT&""
	TS.Cells(10,5).Value = TS.Cells(10,5).Value + TS.Cells(10,16).Value
	TS.Cells(10,16).Value = "0"
	Else
	TS.Cells(10,16).Value = ""&FieldCNT&""
	TS.Cells(10,4).Value = TS.Cells(10,4).Value + TS.Cells(10,16).Value
	TS.Cells(10,16).Value = "0"
	End If
ElseIf FieldSOL = "7411" Then
	If FieldSTAT = "A" Then
	TS.Cells(11,16).Value = ""&FieldCNT&""
	TS.Cells(11,5).Value = TS.Cells(11,5).Value + TS.Cells(11,16).Value
	TS.Cells(11,16).Value = "0"
	Else
	TS.Cells(11,16).Value = ""&FieldCNT&""
	TS.Cells(11,4).Value = TS.Cells(11,4).Value + TS.Cells(11,16).Value
	TS.Cells(11,16).Value = "0"
	End If
ElseIf FieldSOL = "7428" Then
	If FieldSTAT = "A" Then
	TS.Cells(12,16).Value = ""&FieldCNT&""
	TS.Cells(12,5).Value = TS.Cells(12,5).Value + TS.Cells(12,16).Value
	TS.Cells(12,16).Value = "0"
	Else
	TS.Cells(12,16).Value = ""&FieldCNT&""
	TS.Cells(12,4).Value = TS.Cells(12,4).Value + TS.Cells(12,16).Value
	TS.Cells(12,16).Value = "0"
	End If
ElseIf FieldSOL = "7469" Then
	If FieldSTAT = "A" Then
	TS.Cells(13,16).Value = ""&FieldCNT&""
	TS.Cells(13,5).Value = TS.Cells(13,5).Value + TS.Cells(13,16).Value
	TS.Cells(13,16).Value = "0"
	Else
	TS.Cells(13,16).Value = ""&FieldCNT&""
	TS.Cells(13,4).Value = TS.Cells(13,4).Value + TS.Cells(13,16).Value
	TS.Cells(13,16).Value = "0"	
	End If
End If
RS.MoveNext
Loop

RS.Close
DB.Close

xlApp.ActiveWorkbook.Save
xlApp.ActiveWorkbook.Close
xlApp.Application.Quit
		
		
Set xlApp = Nothing
Set TS = Nothing
Set FieldSOL = Nothing
Set FieldSTAT = Nothing
Set FieldCNT = Nothing

Set RS = Nothing
Set DB = Nothing
