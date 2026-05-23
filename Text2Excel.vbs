Option Explicit
On Error Resume Next

Dim objFSO, objFile, strLine, strPath, Sum, Def, Cnt, objExcel, objWorkbook, objSheet, row


strPath = "D:\Debit Card Projects\Nirbhay\INC_CRADJ_NFS_21-MAY-2026_E01_126_0.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False


Set objWorkbook = objExcel.Workbooks.Add()
Set objSheet = objWorkbook.Sheets(1)

row = 1

If objFSO.FileExists(strPath) Then
    Set objFile = objFSO.OpenTextFile(strPath, 1)
	Sum = 0
	Def = 0
	Cnt = 0

    Do While Not objFile.AtEndOfStream
        strLine = objFile.ReadLine
        'WScript.Echo (Left(strLine,14)&"|" & Mid(strLine,17,7)& "|" & Mid(strLine,28,1) & "|" & (Mid(strLine,29,17))*1 & "|" & Mid(strLine,46,27) & "|" & Mid(strLine,81,12) & "|" & Mid(strLine,179,10) & vbCrlf)
		objSheet.Cells.NumberFormat = "@"
		objSheet.Cells(row, 1).Value = Left(strLine,14)
		objSheet.Cells(row, 2).Value = Mid(strLine,17,7)
		objSheet.Cells(row, 3).Value = Mid(strLine,28,1)
		objSheet.Cells(row, 4).Value = Mid(strLine,29,17)*1
		objSheet.Cells(row, 5).Value = Mid(strLine,46,27)
		objSheet.Cells(row, 6).Value = Mid(strLine,81,12)
		objSheet.Cells(row, 7).Value = Mid(strLine,179,10)
		

		Cnt = Cnt + 1
		If Mid(strLine,28,1) = "C" Then
		Sum = (Sum + (Mid(strLine,29,17))*1)
		ElseIf Mid(strLine,28,1) = "D" Then
		Def = (Def + (Mid(strLine,29,17))*1)
		Else
		WScript.Echo "No Transaction Type Found"
		End If

		row = row + 1
    Loop

    objFile.Close

	objWorkbook.SaveAs "D:\Debit Card Projects\Nirbhay\Output.xlsx"
	objWorkbook.Close True
	objExcel.Quit
Else
    WScript.Echo "File not found: " & strPath
End If

WScript.Echo ("Total Count in File: " & FormatNumber(Round(Cnt,2)))
WScript.Echo ("Total Credit Amount: " & FormatNumber(Round(Sum,2)))
WScript.Echo ("Total Debit Amount: " & FormatNumber(Round(Def,2)))

Set objFile = Nothing
Set objFSO = Nothing
Set strLine = Nothing
Set strPath = Nothing
Set Sum = Nothing
Set Def = Nothing
Set Cnt = Nothing
Set objExcel = Nothing
Set objWorkbook = Nothing
Set objSheet = Nothing
Set row = Nothing