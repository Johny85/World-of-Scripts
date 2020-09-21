Option Explicit
On Error Resume Next

Dim FSO, FLD, FIL, TS
Dim strFolder, WS, intRow, Txn_Type, ExcelObject1, File_Type, File_Date, Description, Reference, Debit_Amount, Credit_Amount
Const ForReading = 1, ForWriting = 2, ForAppending = 8 


	'Change as needed
	strFolder = "E:\Project UPI Recon\CRET Files"
	'Create the filesystem object
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set WS = FSO.CreateTextFile("E:\Project UPI Recon\Excel Files\CRET_NPCI.txt", ForWriting)
	Set ExcelObject1 = CreateObject("Excel.Application")

	
	set FLD = FSO.GetFolder(strFolder)

	'loop through the folder and get the files
For Each Fil In FLD.Files

		'Open the file to read
		Call ExcelObject1.Workbooks.Open(fil.Path, ForReading)
		Set TS = ExcelObject1.ActiveWorkbook.Worksheets(1)
		intRow = 1
		File_Date = Mid(Fil.Name, 12, 6)
		
				Select Case Mid(File_Date,3,4)
				Case "0219"
				File_Date = Left(File_Date,2)&"-Feb-2019"
				Case "0319"
				File_Date = Left(File_Date,2)&"-Mar-2019"
				Case "0419"
				File_Date = Left(File_Date,2)&"-Apr-2019"
				Case "0519"
				File_Date = Left(File_Date,2)&"-May-2019"
				Case "0619"
				File_Date = Left(File_Date,2)&"-Jun-2019"
				Case "0719"
				File_Date = Left(File_Date,2)&"-Jul-2019"
				Case "0819"
				File_Date = Left(File_Date,2)&"-Aug-2019"
				Case "0919"
				File_Date = Left(File_Date,2)&"-Sep-2019"
				Case "1019"
				File_Date = Left(File_Date,2)&"-Oct-2019"
				Case "1119"
				File_Date = Left(File_Date,2)&"-Nov-2019"
				Case "1219"
				File_Date = Left(File_Date,2)&"-Dec-2019"
				End Select
		
		File_Type = Mid(Fil.Name, 19, 2)
		
				
Do Until TS.Cells(intRow,2).Value = "Ref. No / RRN / date"
intRow = intRow +1		
Loop
intRow = intRow +1
Do Until TS.Cells(intRow,1).Value = "Adjustment Sub Total"



		' Read a line from the text file into a string called Line
		If  TS.Cells(intRow, 1).Value = ""  OR TS.Cells(intRow, 2).Value = "" Then
			intRow = intRow +1
		Else 
			Description = TS.Cells(intRow, 1).Value
			Reference = TS.Cells(intRow, 2).Value
			Debit_Amount = TS.Cells(intRow, 3).Value
			Credit_Amount = TS.Cells(intRow, 4).Value
			'WSCript.Echo 
		
If Right(Reference, 03) = "/U2" OR  Right(Reference, 03) = "/U3" OR  Right(Reference, 03) = "/UC" OR  Right(Reference, 03) = "/UU" Then
WS.Write (File_Date) & "|" & (File_Type) & "|" & (Reference) & "|" & (Description) & "|" & (Debit_Amount*1) & "|" & (Credit_Amount*1) & "|" & (Left(Right(Reference, 26), 12)) & "|" & (Left(Right(Reference, 13), 10)) & vbCrLf
Else
WS.Write (File_Date) & "|" & (File_Type) & "|" & (Reference) & "|" & (Description) & "|" & (Debit_Amount*1) & "|" & (Credit_Amount*1) & "|" & (Left(Right(Reference, 29), 12)) & "|" & (Left(Right(Reference, 16), 10)) & vbCrLf
End If
			
				intRow = intRow +1		
			
	          End If
	     
	    
Loop 

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
