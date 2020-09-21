Option Explicit
On Error Resume Next

Dim RSO, FLD, Fil, RS, WS, WSO
Dim strFolder, Line,arr
Const ForReading = 1, ForWriting = 2, ForAppending = 8 


	'Change as needed
	strFolder = "D:\Meity\Daily"

	'Create the filesystem object
	Set RSO = CreateObject("Scripting.FileSystemObject")
	Set WSO = CreateObject("Scripting.FileSystemObject")

	Set WS = WSO.CreateTextFile("D:\Meity\Incorrect.csv", ForWriting)
	Set FLD = RSO.GetFolder(strFolder)
	
	'loop through the folder and get the files

	For Each Fil In FLD.Files

Set RS = RSO.OpenTextFile(fil.Path, ForReading)

If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
		' Read a line from the text file into a string called Line
		Line = RS.ReadLine
		Line = Replace(Line,chr(34),"")
		arr = split(Line,",")

	WS.Write (Fil.Name) & "|" & arr(0) & "|" & arr(1) & "|" & arr(2) & "|" & arr(3) & "|" & arr(4) & vbCrLf
	
	
Loop 

		'Close the file
		RS.Close
		
Next
WS.Close

'Clean up
Set RS = Nothing
Set WS = Nothing
Set TS = Nothing
Set FLD = Nothing
Set RSO = Nothing
Set WSO = Nothing
Set strFolder = Nothing





	strFolder = "C:\Users\PR172959\Documents\Meity Data Apr Onwards\MEITY_HIST"

	'Create the filesystem object
	Set RSO = CreateObject("Scripting.FileSystemObject")
	Set WSO = CreateObject("Scripting.FileSystemObject")

	Set WS = WSO.CreateTextFile("D:\Meity\correct.csv", ForWriting)
	Set FLD = RSO.GetFolder(strFolder)
	
	'loop through the folder and get the files

	For Each Fil In FLD.Files

Set RS = RSO.OpenTextFile(fil.Path, ForReading)

If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
		' Read a line from the text file into a string called Line
		Line = RS.ReadLine
		Line = Replace(Line,chr(34),"")
		arr = split(Line,",")

	WS.Write (Fil.Name) & "|" & arr(0) & "|" & arr(1) & "|" & arr(2) & "|" & arr(3) & "|" & arr(4) & vbCrLf
	
	
Loop 

		'Close the file
		RS.Close
		
Next
WS.Close

'Clean up
Set RS = Nothing
Set WS = Nothing
Set TS = Nothing
Set FLD = Nothing
Set RSO = Nothing
Set WSO = Nothing
Set strFolder = Nothing