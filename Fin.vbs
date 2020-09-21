Option Explicit
On Error Resume Next

Dim RSO, strFolder, FLD, fil, RS, Line, GL_Date, Val_Date, CMD, DateC
Dim DB, ORS, DateP, OPBal, CLBal, WSO, WS, arr

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

DateP = inputbox ("Enter Recon START Date in 'DD-MM-YYYY' format")
DateP = (CDate(DateP))-1
DateC = inputbox ("Enter Recon END Date in 'DD-MM-YYYY' format")

Set RSO = CreateObject("Scripting.FileSystemObject")
Set DB = CreateObject("ADODB.Connection")

strFolder = "C:\MCon_Recon\FinacleFile\RECHARGE"
Set FLD = RSO.GetFolder(strFolder)

DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"
DB.Execute("Delete from FINACLE_REC;")
DB.Execute("Alter Table FINACLE_REC Alter Column SERIAL_FIN COUNTER(1,1);")

For Each fil In FLD.Files    

Set RS = RSO.OpenTextFile(fil.Path, ForReading)
Set WSO = CreateObject("Scripting.FileSystemObject")

Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,",","")
Line = Replace(Line,"Cr","")

GL_Date = Trim(mid(Line,2,10))
Val_Date = Trim(mid(Line,38,11))

Set WS = WSO.CreateTextFile("C:\MCon_Recon\FinacleFile\RECHARGE\FINACLE_REC.csv", ForWriting)

If (Right(GL_Date,4) = "2020" AND Right(Val_Date,4) = "2020") Then
WS.Write Trim(mid(Line,2,10)) & "|" & Trim(mid(Line,13,10)) & "|" & Trim(mid(Line,23,12)) & "|" & Trim(mid(Line,38,11)) & "|" & Trim(mid(Line,50,28)) & "|" & Trim(mid(Line,55,12)) & "|" & Trim(mid(Line,55,12))&Trim(mid(Line,81,15))&Trim(mid(Line,105,12)) & "|" & Trim(mid(Line,81,15)) & "|" & Trim(mid(Line,105,12)) & "|" & Trim(mid(Line,122,21)) & vbCrLf
'WScript.Echo "Recharge--"&Trim(mid(Line,55,12))&"--"&Trim(mid(Line,85,11))&"--"&Trim(mid(Line,105,12))&"--"&Trim(mid(Line,122,21))&vbCrlf
End If
Loop
Next

RS.Close
WS.Close
Set RS = Nothing
Set WS = Nothing
Set strFolder = Nothing
Set FLD = Nothing
Set fil = Nothing


Set RS = RSO.OpenTextFile("C:\MCon_Recon\FinacleFile\RECHARGE\FINACLE_REC.csv", ForReading)
Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "FINACLE_REC", DB, adOpenStatic, adLockOptimistic
If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,chr(34),"")
arr = split(Line,"|")

ORS.AddNew
ORS("GL_DATE") = arr(0)
ORS("TRAN_ID") = arr(1)
ORS("REFER_NUM") = arr(2)
ORS("VAL_DATE") = arr(3)
ORS("PARTICULARS") = arr(4)
ORS("RRN") = arr(5)
ORS("UNIQUE_REF") = arr(6)
ORS("DEBIT_AMT") = arr(7)
ORS("CREDIT_AMT") = arr(8)
ORS("BALANCE") = arr(9)
ORS.Update
	
Loop

RS.Close
ORS.Close
Set RS = Nothing


'WScript.Echo ("DB Connection Success")
Set RS = DB.Execute("SELECT last(BALANCE) from FINACLE_REC where GL_DATE = '"& DateP &"';")
OPBal = RS.Fields(0)
RS.Close
Set RS = Nothing

Set RS = DB.Execute("SELECT last(BALANCE) from FINACLE_REC where GL_DATE = '"& DateC &"';")
CLBal = RS.Fields(0)
RS.Close
Set RS = Nothing
WScript.Echo ("Opening Balance (Recharge): "& OPBal)
WScript.Echo ("Closing Balance (Recharge): "& CLBal)

DB.Execute("UPDATE FINACLE_REC SET ENTERED_BY = 'SYSTEM' WHERE PARTICULARS Like '%MBK/%';")

DB.Execute("UPDATE FINACLE_REC SET ENTERED_BY = 'MANUAL' WHERE PARTICULARS Not Like '%MBK/%';")

DB.Execute("UPDATE FINACLE_REC INNER JOIN BILLDESK_REC ON FINACLE_REC.UNIQUE_REF = BILLDESK_REC.UNIQUE_REF SET FINACLE_REC.BILLDESK_STATUS = BILLDESK_REC.STATUS;")




Set RSO = Nothing

Set Line = Nothing
Set ORS = Nothing
Set GL_Date = Nothing
Set Val_Date = Nothing

''''''''''''''BILLPAY'''''''''''''''''''''''''''

Set RSO = CreateObject("Scripting.FileSystemObject")

strFolder = "C:\MCon_Recon\FinacleFile\BILLPAY"
Set FLD = RSO.GetFolder(strFolder)
For Each fil In FLD.Files

DB.Execute("Delete from FINACLE_BP;")

DB.Execute("Alter Table FINACLE_BP Alter Column SERIAL_FIN COUNTER(1,1);")

Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "FINACLE_BP", DB, adOpenStatic, adLockOptimistic

Set RS = RSO.OpenTextFile(fil.Path, ForReading)
    
Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,",","")
Line = Replace(Line,"Cr","")

GL_Date = Trim(mid(Line,2,10))
Val_Date = Trim(mid(Line,38,11))

If (Right(GL_Date,4) = "2020" AND Right(Val_Date,4) = "2020") Then
'WSCript.Echo ("BillPay--"&Trim(mid(Line,55,12))&"--"&Trim(mid(Line,81,15))&"--"&Trim(mid(Line,105,12))&"--"&Trim(mid(Line,122,21))&vbCrlf)
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


'DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"
'WScript.Echo ("DB Connection Success")
Set RS = DB.Execute("SELECT last(BALANCE) from FINACLE_REC where GL_DATE = '"& DateP &"';")
OPBal = RS.Fields(0)
RS.Close
Set RS = Nothing
WScript.Sleep 5000
Set RS = DB.Execute("SELECT last(BALANCE) from FINACLE_REC where GL_DATE = '"& DateC &"';")
CLBal = RS.Fields(0)
RS.Close
Set RS = Nothing
WScript.Echo ("Opening Balance (Recharge): "& OPBal)
WScript.Echo ("Closing Balance (Recharge): "& CLBal)

DB.Execute("UPDATE FINACLE_BP SET ENTERED_BY = 'SYSTEM' WHERE PARTICULARS Like '%MBK/%';")

DB.Execute("UPDATE FINACLE_BP SET ENTERED_BY = 'MANUAL' WHERE PARTICULARS Not Like '%MBK/%';")

DB.Execute("UPDATE FINACLE_BP INNER JOIN BILLDESK_BP ON FINACLE_BP.UNIQUE_REF = BILLDESK_BP.UNIQUE_REF SET FINACLE_BP.BILLDESK_STATUS = BILLDESK_BP.STATUS;")



DB.Close
Set RSO = Nothing
Set strFolder = Nothing
Set FLD = Nothing
Set fil = Nothing
Set Line = Nothing
Set RS = Nothing
Set ORS = Nothing
Set DB = Nothing
Set GL_Date = Nothing
Set Val_Date = Nothing