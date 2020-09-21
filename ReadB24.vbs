Option Explicit
On Error Resume Next

Dim RSO, FLD, fil, RS, MyExcelFilePath, ExcelObject, SheetObject, Line
Const ForReading = 1, ForWriting = 2, ForAppending = 8 


	'Change as needed
	strFolder = "D:\ScantoPay\B24"

	'Create the filesystem object
	Set RSO = CreateObject("Scripting.FileSystemObject")

	MyExcelFilePath = "D:\ScantoPay\B24.xlsx"
	Set ExcelObject = CreateObject("Excel.Application")
	ExcelObject.WorkBooks.Open MyExcelFilePath
	Set SheetObject = ExcelObject.ActiveWorkbook.Worksheets(1)
	SheetObject.UsedRange.ClearContents
	
	
	
	Set FLD = RSO.GetFolder(strFolder)
	
	'loop through the folder and get the files
	X = 1
	Call ExcelObject.Workbooks.Open(fil.Path, ForWriting)

	For Each Fil In FLD.Files

	Set RS = RSO.OpenTextFile(fil.Path, ForReading)
			
	Do Until RS.AtEndOfStream

		' Read a line from the text file into a string called Line
		Line = RS.ReadLine
		WScript.Echo (Line)

				If Mid(Line,187,23) = "M     VISA  PAYMENT  IN" Then
				
			Line = Replace(Line,"M     VISA  PAYMENT  IN","      MVISAPAYMENT   IN")
			
		SheetObject.Cells(x, 1).Value = Trim(Mid(Line,0,42)) 
		SheetObject.Cells(x, 2).Value = Trim(Mid(Line,43,60))
		SheetObject.Cells(x, 3).Value = Trim(Mid(Line,61,63))
		SheetObject.Cells(x, 4).Value = Trim(Mid(Line,64,68)) 
		SheetObject.Cells(x, 5).Value = Trim(Mid(Line,69,76))
		SheetObject.Cells(x, 6).Value = Trim(Mid(Line,77,90))
		SheetObject.Cells(x, 7).Value = Trim(Mid(Line,91,107)) 
		SheetObject.Cells(x, 8).Value = Trim(Mid(Line,108,118)) 
		SheetObject.Cells(x, 9).Value = Trim(Mid(Line,119,121)) 
		SheetObject.Cells(x, 10).Value = Trim(Mid(Line,122,130)) 
		SheetObject.Cells(x, 11).Value = Trim(Mid(Line,131,132)) 
		SheetObject.Cells(x, 12).Value = Trim(Mid(Line,133,138)) 
		SheetObject.Cells(x, 13).Value = Trim(Mid(Line,139,144)) 
		SheetObject.Cells(x, 14).Value = Trim(Mid(Line,145,151)) 
		SheetObject.Cells(x, 15).Value = Trim(Mid(Line,152,156)) 
		SheetObject.Cells(x, 16).Value = Trim(Mid(Line,157,162)) 
		SheetObject.Cells(x, 17).Value = Trim(Mid(Line,163,165)) 
		SheetObject.Cells(x, 18).Value = Trim(Mid(Line,166,192)) 
		SheetObject.Cells(x, 19).Value = Trim(Mid(Line,193,207)) 
		SheetObject.Cells(x, 20).Value = Trim(Mid(Line,208,211)) 
		SheetObject.Cells(x, 21).Value = Trim(Mid(Line,212,218)) 
     		x = x + 1
		
		End If
		
		
		SheetObject.Cells(x, 1).Value = Trim(Mid(Line,0,42)) 
		SheetObject.Cells(x, 2).Value = Trim(Mid(Line,43,60))
		SheetObject.Cells(x, 3).Value = Trim(Mid(Line,61,63))
		SheetObject.Cells(x, 4).Value = Trim(Mid(Line,64,68)) 
		SheetObject.Cells(x, 5).Value = Trim(Mid(Line,69,76))
		SheetObject.Cells(x, 6).Value = Trim(Mid(Line,77,90))
		SheetObject.Cells(x, 7).Value = Trim(Mid(Line,91,107)) 
		SheetObject.Cells(x, 8).Value = Trim(Mid(Line,108,118)) 
		SheetObject.Cells(x, 9).Value = Trim(Mid(Line,119,121)) 
		SheetObject.Cells(x, 10).Value = Trim(Mid(Line,122,130)) 
		SheetObject.Cells(x, 11).Value = Trim(Mid(Line,131,132)) 
		SheetObject.Cells(x, 12).Value = Trim(Mid(Line,133,138)) 
		SheetObject.Cells(x, 13).Value = Trim(Mid(Line,139,144)) 
		SheetObject.Cells(x, 14).Value = Trim(Mid(Line,145,151)) 
		SheetObject.Cells(x, 15).Value = Trim(Mid(Line,152,156)) 
		SheetObject.Cells(x, 16).Value = Trim(Mid(Line,157,162)) 
		SheetObject.Cells(x, 17).Value = Trim(Mid(Line,163,165)) 
		SheetObject.Cells(x, 18).Value = Trim(Mid(Line,166,192)) 
		SheetObject.Cells(x, 19).Value = Trim(Mid(Line,193,207)) 
		SheetObject.Cells(x, 20).Value = Trim(Mid(Line,208,211)) 
		SheetObject.Cells(x, 21).Value = Trim(Mid(Line,212,218)) 
     		x = x + 1		
		
Loop 

		'Close the file
		RS.Close
		
Next

SheetObject.Saved = True
SheetObject.Close
ExcelObject.Quit
	
'Clean up
Set RS = Nothing
Set FLD = Nothing
Set RSO = Nothing
Set ExcelObject = Nothing
Set SheetObject = Nothing