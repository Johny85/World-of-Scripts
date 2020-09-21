Option Explicit
On Error Resume Next

Dim RSO, FLD, Fil, RS, WS, WSO
Dim strFolder, Line
Dim Var1, Var2, Var3, Var4, Var5, Var6, Var7, Var8, Var9, Var10, Var11, Var12, Var13, Var14, Var15, Var16, Var17, Var19, Var20, Var21
Const ForReading = 1, ForWriting = 2, ForAppending = 8 


	'Change as needed
	strFolder = "D:\ScantoPay\B24"

	'Create the filesystem object
	Set RSO = CreateObject("Scripting.FileSystemObject")
	Set WSO = CreateObject("Scripting.FileSystemObject")

	Set WS = WSO.CreateTextFile("D:\ScantoPay\Base24.csv", ForWriting)
	Set FLD = RSO.GetFolder(strFolder)
	
	'loop through the folder and get the files

	For Each Fil In FLD.Files

Set RS = RSO.OpenTextFile(fil.Path, ForReading)

If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
		' Read a line from the text file into a string called Line
		Line = RS.ReadLine
		Line = Replace(Line,chr(34),"")
		

If Mid(Line,187,23) = "M     VISA  PAYMENT  IN" Then
				
Line = Replace(Line,"M     VISA  PAYMENT  IN","      MVISAPAYMENT   IN")

End If

Var1 = Trim(Mid(Line,1,42)) 
Var2 = Trim(Mid(Line,43,6))
Var3 = Trim(Mid(Line,61,2))
Var4 = Trim(Mid(Line,64,4)) 
Var5 = Trim(Mid(Line,69,7))
Var6 = Trim(Mid(Line,77,13))
Var7 = Trim(Mid(Line,91,16)) 
Var8 = Trim(Mid(Line,108,10)) 
Var9 = Trim(Mid(Line,119,2)) 
Var10 = Trim(Mid(Line,122,8)) 
Var11 = Trim(Mid(Line,131,1)) 
Var12 = Trim(Mid(Line,133,5)) 
Var13 = Trim(Mid(Line,139,5)) 
Var14 = Trim(Mid(Line,145,6)) 
Var15 = Trim(Mid(Line,152,4)) 
Var16 = Trim(Mid(Line,157,5)) 
Var17 = Trim(Mid(Line,163,2))
Var19 = Trim(Left(Right(Line,25),13))
Var20 = Trim(Left(Right(Line,10),2)) 
Var21 = Trim(Right(Line,7)) 
			
WS.Write Var1 & "|" & Var2 & "|" & Var3 & "|" & Var4 & "|" & Var5 & "|" & Var6 & "|" & Var7 & "|" & Var8 & "|" & Var9 & "|" & Var10 & "|" & Var11 & "|" & Var12 & "|" & Var13 & "|" & Var14 & "|" & Var15 & "|" & Var16 & "|" & Var17 & "|" & Var20 & "|" & Var21 & vbCrLf	

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





strFolder = "D:\ScantoPay\Rupay\16092019"

Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")
Set WSP = WSO.CreateTextFile("D:\ScantoPay\Rupay.csv", ForWriting)

For Each Fil In FLD.Files

Set RS = RSO.OpenTextFile(fil.Path, ForReading)

If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,chr(34),"")
arr = split(Line,"|")

WSP.Write arr(0) & "|" & arr(1) &"|"& arr(2) & "|" & arr(3) & "|"& arr(4) & "|" & arr(5) & "|" & arr(6) & "|" & arr(7) & "|" & arr(8) & "|" & arr(0) & "|" & arr(1) & "|" & arr(2) & "|" & arr(3) & "|" & arr(4) & "|" & arr(5) & "|" & arr(6) & "|" & arr(7) & "|" & arr(8) & "|" & arr(0) & "|" & arr(1) & "|" & arr(2) & "|" & arr(3) & "|" & arr(4) & "|" & arr(5) & "|" & arr(6) & "|" & arr(7) & "|" & arr(8) & "|" & arr(0) & "|" & arr(1) & "|" & arr(2) & "|" & arr(3) & "|" & arr(4) & "|" & arr(5) & "|" & arr(6) & "|" & arr(7) & "|" & arr(8) & "|" & arr(0) & "|" & arr(1) & "|" & arr(2) & "|" & arr(3) & "|" & arr(4) & "|" & arr(5) & "|" & arr(6) & "|" & arr(7) & "|" & arr(8) & "|" & arr(0) & "|" & arr(1) & "|" & arr(2) & "|" & arr(3) & "|" & arr(4) & "|" & arr(5) & "|" & arr(6) & "|" & arr(7) & "|" & arr(8) & "|" & arr(0) & "|" & arr(1) & "|" & arr(2) & "|" & arr(3) & "|" & arr(4) & "|" & arr(5) & "|" & arr(6) & "|" & arr(7) & "|" & arr(8) & "|" & arr(0) & "|" & arr(1) & "|" & arr(2) & "|" & arr(3) & "|" & arr(4)& "|" & arr(5) & "|" & arr(6) & "|" & arr(7) & "|" & arr(8) & "|" & vbCrLf


& vbCrLf

Loop


	RS.Close
	WS.Close
	WSP.Close
	
	
Set SLine = Nothing
Set RS = Nothing
Set WSP = Nothing
Set RSO = Nothing
Set WSO = Nothing

