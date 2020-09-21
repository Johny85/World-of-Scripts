Option Explicit
On Error Resume Next

Dim RSO, WSO, WS, strFolder, FLD, fil, RS, Line, DB, ORS, GL_Date, Val_Date, CMD, DateC, DateP, OPBal, CLBal, arr

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

'DateP = inputbox ("Enter Recon START Date in 'DD-MM-YYYY' format")
'DateP = (CDate(DateP))-1
'DateC = inputbox ("Enter Recon END Date in 'DD-MM-YYYY' format")

Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

strFolder = "C:\MCon_Recon\FinacleFile\RECHARGE"
Set FLD = RSO.GetFolder(strFolder)
For Each fil In FLD.Files

Set RS = RSO.OpenTextFile(fil.Path, ForReading)
Set WS = WSO.CreateTextFile("C:\MCon_Recon\FinacleFile\RECHARGE\FINACLE_REC.csv", ForWriting)

Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,",","")
Line = Replace(Line,"Cr","")

GL_Date = Trim(mid(Line,2,10))
Val_Date = Trim(mid(Line,38,11))

If (Right(GL_Date,4) = "2020" AND Right(Val_Date,4) = "2020") Then
WS.Write Trim(mid(Line,2,10)) & "|" & Trim(mid(Line,13,10)) & "|" & Trim(mid(Line,23,12)) & "|" & Trim(mid(Line,38,11)) & "|" & Trim(mid(Line,50,28)) & "|" & Trim(mid(Line,55,12)) & "|" & Trim(mid(Line,55,12))&Trim(mid(Line,81,15))&Trim(mid(Line,105,12)) & "|" & Trim(mid(Line,81,15)) & "|" & Trim(mid(Line,105,12)) & "|" & Trim(mid(Line,122,21)) & vbCrlf
End If
Loop

Next
RS.Close
Set RS = Nothing
Set Line = Nothing

Set RS = Nothing
Set ORS = Nothing
Set RSO = Nothing
Set WSO = Nothing
Set RS = Nothing
Set WS = Nothing
Set strFolder = Nothing
Set FLD = Nothing
Set fil = Nothing
Set Line = Nothing
Set GL_Date = Nothing
Set Val_Date = Nothing




