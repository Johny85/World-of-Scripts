Option Explicit
On Error Resume Next

Dim RSO, RS, strFolder, FLD, fil, Line, arr, xlApp, TS, i, j

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

Set RSO = CreateObject("Scripting.FileSystemObject")

strFolder = "C:\TK\SPGRS"

Set xlApp = CreateObject("Excel.Application")
xlApp.Workbooks.Open("C:\TK\Output.xlsx")
Set TS = xlApp.ActiveWorkbook.Worksheets("Sample")

Set FLD = RSO.GetFolder(strFolder)
i=0
j=0

For Each fil In FLD.Files    
Set RS = RSO.OpenTextFile(fil.Path, ForReading)

If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,chr(34),"")
arr = split(Line,",")
i = i+1
j = j+1
TS.Cells(i,j).Value = arr(0)
j = j+1
TS.Cells(i,j).Value = arr(1)
j = j+1
TS.Cells(i,j).Value = arr(2)
j = j+1
TS.Cells(i,j).Value = arr(3)
j = j+1
TS.Cells(i,j).Value = arr(4)
j = j+1
TS.Cells(i,j).Value = arr(5)
j = j+1
TS.Cells(i,j).Value = arr(6)
j = j+1
TS.Cells(i,j).Value = arr(7)
j = j+1
TS.Cells(i,j).Value = arr(8)
j = j+1
TS.Cells(i,j).Value = arr(9)
j = j+1
TS.Cells(i,j).Value = arr(10)
j = j+1
TS.Cells(i,j).Value = arr(11)
j = j+1
TS.Cells(i,j).Value = arr(12)
j = j+1
TS.Cells(i,j).Value = arr(13)
j = j+1
TS.Cells(i,j).Value = arr(14)
j = j+1
TS.Cells(i,j).Value = arr(15)
j = j+1
TS.Cells(i,j).Value = arr(16)
j = j+1
TS.Cells(i,j).Value = arr(17)
j = j+1
TS.Cells(i,j).Value = arr(18)
j = j+1
TS.Cells(i,j).Value = arr(19)
j = j+1
TS.Cells(i,j).Value = arr(20)
j = j+1
TS.Cells(i,j).Value = arr(21)
j = j+1
TS.Cells(i,j).Value = arr(22)
j = j+1
TS.Cells(i,j).Value = arr(23)

Loop

Next


xlApp.ActiveWorkbook.Save
xlApp.ActiveWorkbook.Close
xlApp.Application.Quit

Set xlApp = Nothing
Set xlBook = Nothing
Set xlSheet = Nothing
Set WSO = Nothing
Set RSO = Nothing
Set WS = Nothing
Set RS = Nothing
