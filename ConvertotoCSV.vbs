Option Explicit
On Error Resume Next

Dim objBook, objExcel, objSheet, RSO, strFolder, FLD, fil

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 
Const xlCSV = 6

Set objExcel = CreateObject("Excel.Application")
Set RSO = CreateObject("Scripting.FileSystemObject")

strFolder = "C:\Users\PR172959\AppData\Local\Temp\ACT"
Set FLD = RSO.GetFolder(strFolder)

For Each fil In FLD.Files
Set objBook = objExcel.Workbooks.Open(fil.Path)

objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objSheet = objBook.Worksheets(1)
objSheet.SaveAs "C:\Users\PR172959\AppData\Local\Temp\ACT\"&fil.Name&".csv", xlCSV
Next

objExcel.ActiveWorkbook.Save = False
objExcel.ActiveWorkbook.Close = True
objExcel.Quit

Set objBook = Nothing
Set objExcel = Nothing
Set objSheet = Nothing
Set RSO = Nothing
