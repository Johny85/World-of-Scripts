Option Explicit
On Error Resume Next

Dim RSO, FLD, FIL, RS, WS, WSO
Dim strFolder, Line1, Line2, Line3, Line
'Dim Var1, Var2, Var3, Var4, Var5, Var6, Var7, Var8, Var9
Dim arr
Const ForReading = 1, ForWriting = 2, ForAppending = 8 


	'Change as needed
	strFolder = "D:\ScantoPay\Visa"

	'Create the filesystem object
	Set RSO = CreateObject("Scripting.FileSystemObject")
	Set WSO = CreateObject("Scripting.FileSystemObject")

	Set WS = WSO.CreateTextFile("D:\ScantoPay\VISA.csv", ForWriting)
	Set FLD = RSO.GetFolder(strFolder)
	
	'loop through the folder and get the files

	For Each Fil In FLD.Files

		'Open the file to read
			Set RS = RSO.OpenTextFile(fil.Path, ForReading)
			
	Do Until RS.AtEndOfStream

		' Read a line from the text file into a string called Line
	Line = RS.ReadLine
	
	If Mid(Line,24,3) = "AUG" OR Mid(Line,37,5) = "CA ID" OR Mid(Line,22,15) = "FEE JURIS: VISA" Then
	
	'Line1 = RS.SkipLine & RS.SkipLine & RS.ReadLine
	'Line2 = RS.SkipLine & RS.SkipLine & RS.SkipLine & RS.SkipLine & RS.ReadLine
	'Line3 = RS.SkipLine & RS.SkipLine & RS.SkipLine & RS.SkipLine & RS.SkipLine & RS.ReadLine
	'Var1 = Mid(Line1,22,5)
	'Var2 = Mid(Line1,37,16)
	'Var3 = Mid(Line1,57,12)
	'Var4 = Mid(Line1,70,6)
	'Var5 = Mid(Line1,134,13)*1
	'Var6 = Mid(Line1,147,2)
	'Var7 = Mid(Line2,80,25)
	'Var8 = Mid(Line3,134,13)*1
	'Var9 = Mid(Line3,147,2)
	
	'arr = Split(Line, vbCrLf)
	'WS.Write arr(3) & "|" & arr(4) & "|" & arr(6) & vbCrLf
	'WS.Write (Var1) & "|" & (Var2) & "|" & (Var3) & "|" & (Var4) & "|" & (Var5) & "|" & (Var6) & "|" & Trim(Var7) & "|" & (Var8) & "|" & (Var9) & vbCrLf
	
	
	WS.Write (Line) & vbCrLf
	End If
	Loop 

RS.Close
		
Next
WS.Close
	
'Clean up
Set RS = Nothing
Set FLD = Nothing
Set RSO = Nothing
Set WSO = Nothing