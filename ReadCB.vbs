Option Explicit
On Error Resume Next

Dim RSO, FLD, FIL, RS, WS, WSO
Dim strFolder, Line, RRN, NDate, Amount, Tran_Id
Const ForReading = 1, ForWriting = 2, ForAppending = 8 


	'Change as needed
	strFolder = "E:\Recon\CBS"

	'Create the filesystem object
	Set RSO = CreateObject("Scripting.FileSystemObject")
	Set WSO = CreateObject("Scripting.FileSystemObject")

	Set WS = WSO.CreateTextFile("E:\Recon\CBS.txt", ForWriting)
	Set FLD = RSO.GetFolder(strFolder)
	
	'loop through the folder and get the files

	For Each Fil In FLD.Files

		'Open the file to read
			Set RS = RSO.OpenTextFile(fil.Path, ForReading)
			
	Do Until RS.AtEndOfStream

		' Read a line from the text file into a string called Line
		Line = RS.ReadLine

				If Mid(Line,39,5) = "MBK/9" Then
				
				NDate = Mid(Line,2,10)
				Tran_Id = Mid(Line,14,8)
				RRN = Mid(Line,43,12)
				Amount = Mid(Line,90,15)
				End If
				
		WS.Write (NDate) & "|" & (Tran_Id) & "|" & (RRN) & "|" & (Amount*1) & vbCrLf
Loop 

		'Close the file
		RS.Close
		
Next
WS.Close
	
'Clean up
Set RS = Nothing
Set FLD = Nothing
Set RSO = Nothing
Set WSO = Nothing