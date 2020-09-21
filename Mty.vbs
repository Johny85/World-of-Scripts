Option Explicit
On Error Resume Next

Dim FSO, FLD, FIL, TS
Dim strFolder, WS, ExcelObject1, File_Name
Const ForReading = 1, ForWriting = 2, ForAppending = 8 


	'Change as needed
	strFolder = "E:\Meity\MEITY"
	'Create the filesystem object
	Set FSO = CreateObject("Scripting.FileSystemObject")
	'Set WS = FSO.CreateTextFile("E:\Project UPI Recon\Excel Files\CRET_NPCI.txt", ForWriting)
	Set ExcelObject1 = CreateObject("Excel.Application")

	
	set FLD = FSO.GetFolder(strFolder)

	'loop through the folder and get the files
For Each Fil In FLD.Files

		'Open the file to read
Call ExcelObject1.Workbooks.Open(fil.Path, ForReading)
Set TS = ExcelObject1.ActiveWorkbook.Worksheets(1)

File_Name = fil.Name					
WScript.Echo (File_Name)&"|"&(TS.UsedRange.Rows.Count)

		'Close the file
		TS.Close
		Set TS = Nothing
	
Next

	
'Clean up
ExcelObject1.ActiveWorkBook.Close
Set TS = Nothing
Set WS = Nothing
Set FLD = Nothing
Set FSO = Nothing
