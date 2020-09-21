Option Explicit
'On Error Resume Next

Dim FSO, RSO, WSO, FLD, Fil, RS, WS
Dim strFolder, Line, NDate, Tran_Id, Amount, RRN, X,Y,Z

Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set FSO = CreateObject("Scripting.FileSystemObject")
Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")
strFolder = "E:\Recon\CBS"
Set FLD = RSO.GetFolder(strFolder)
For Each Fil In FLD.Files

If lcase(FSO.GetExtensionName(Fil.Name)) = "rpt" Then
Fil.name = "CBSRecord.txt"
Exit For
End If

Next


Set RS = RSO.OpenTextFile("E:\Recon\CBS\CBSRecord.txt", ForReading)
Set WS = WSO.CreateTextFile("E:\Recon\CBS.txt", ForWriting)

X=1


		' Read a line from the text file into a string called Line
		Line = RS.ReadLine
		
		Do While X<71
		RS.SkipLine
		X = X+1
		Loop
				
		
		Do Until RS.AtEndOfStream
		For Y=1 to 35
		WScript.Echo (RS.ReadLine)
		WS.Write (RS.ReadLine) & vbNewLine
		'RS.SkipLine
		Next
		
		For Z=1 to 12
		RS.SkipLine
		Next
		Loop
RS.Close
WS.Close		


Set strFolder = Nothing
Set FLD = Nothing
Set Fil = Nothing
Set RSO = Nothing
Set WSO = Nothing
Set RS = Nothing
Set WS = Nothing