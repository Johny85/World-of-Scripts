Option Explicit
On Error Resume Next

Dim RSO, strFolder, FLD, fil, RS, Line, DB, ORS, GL_Date, Val_Date, CMD, OPBal, CLBal

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"

WScript.Echo ("FINACLE: Module Started")

Set RSO = CreateObject("Scripting.FileSystemObject")

WScript.Echo ("FINACLE: Previous Data Deleted")

DB.Execute("Delete from FINACLE_REC;")

DB.Execute("Alter Table FINACLE_REC Alter Column SERIAL_FIN COUNTER(1,1);")

Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "FINACLE_REC", DB, adOpenStatic, adLockOptimistic

Set RS = RSO.OpenTextFile(fil.Path, ForReading)

WScript.Echo ("FINACLE: Uploading CBS Statement to DB")

strFolder = "C:\MConnect\Reports"
Set FLD = RSO.GetFolder(strFolder)
For Each fil In FLD.Files

Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,",","")
Line = Replace(Line,"Cr","")

GL_Date = Trim(mid(Line,2,10))
Val_Date = Trim(mid(Line,38,11))



If (Right(GL_Date,4) = "2020" AND Right(Val_Date,4) = "2020") Then
'WSCript.Echo ("Recharge--"&Trim(mid(Line,55,12))&"--"&Trim(mid(Line,85,11))&"--"&Trim(mid(Line,105,12))&"--"&Trim(mid(Line,122,21))&vbCrlf)
ORS.AddNew
ORS("GL_DATE") = Trim(mid(Line,2,10))
ORS("TRAN_ID") = Trim(mid(Line,13,10))
ORS("REFER_NUM") = Trim(mid(Line,23,12))
ORS("VAL_DATE") = Trim(mid(Line,38,11))
ORS("PARTICULARS") = Trim(mid(Line,50,28))
ORS("RRN") = Trim(mid(Line,55,12))
ORS("UNIQUE_REF") = Trim(mid(Line,55,12))&Trim(mid(Line,81,15))&Trim(mid(Line,105,12))
ORS("DEBIT_AMT") = Trim(mid(Line,81,15))
ORS("CREDIT_AMT") = Trim(mid(Line,105,12))
ORS("BALANCE") = Trim(mid(Line,122,21))
ORS.Update
End If
Loop
Next
ORS.Close
RS.Close
Set RS = Nothing

WScript.Echo ("FINACLE: Data Upload Success")


DB.Close
Set DB = Nothing
Set RSO = Nothing
Set strFolder = Nothing
Set FLD = Nothing
Set fil = Nothing
Set Line = Nothing
Set ORS = Nothing
Set GL_Date = Nothing
Set Val_Date = Nothing
Set DateP = Nothing
Set DateC = Nothing

