Option Explicit
On Error Resume Next

Dim RSO, RS, Line, arr, DB, ORS, strFolder, FLD, fil
Dim GL_Date, Val_Date, CMD, DateC, DateP, DatePr, OPBal, CLBal
Dim objBook, objExcel, objSheet, Diff, Diff2, Diff3, Result_Rec

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 
Const xlCSV = 6

DateP = inputbox ("Enter Recon START Date in 'DD-MM-YYYY' format")
DatePr = (CDate(DateP))-1
DateC = inputbox ("Enter Recon END Date in 'DD-MM-YYYY' format")

Set RSO = CreateObject("Scripting.FileSystemObject")
Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"

'DB.Execute("Delete from BILLDESK_REC_TEMP;")
'WScript.Sleep 10000
'DB.Execute("Alter Table BILLDESK_REC_TEMP Alter Column SERIAL_SUP COUNTER(1,1)")
'WScript.Sleep 5000
'DB.Execute("Delete from SUPPORT_REC_TEMP;")
'WScript.Sleep 10000
'DB.Execute("Alter Table SUPPORT_REC_TEMP Alter Column SERIAL_SUP COUNTER(1,1)")
'WScript.Sleep 5000
DB.Execute("Insert into BILLDESK_REC_TEMP Select * From BILLDESK_REC;")
'WScript.Sleep 10000
DB.Execute("Delete from BILLDESK_REC;")
'WScript.Sleep 10000
DB.Execute("Insert into SUPPORT_REC_TEMP Select * From SUPPORT_REC;")
'WScript.Sleep 10000
DB.Execute("Delete from SUPPORT_REC;")
'WScript.Sleep 10000
WScript.Echo ("SUPPORT: Previous Files Deleted")
Set RS = DB.Execute("Alter Table SUPPORT_REC Alter Column SERIAL_SUP COUNTER(1,1)")
Set RS = Nothing

strFolder = "C:\MCon_Recon\SupportFile\Recharge"
Set FLD = RSO.GetFolder(strFolder)
WScript.Echo ("SUPPORT: Ready to Insert New Records")

For Each fil In FLD.Files
Set RS = RSO.OpenTextFile(fil.Path, ForReading)
Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "SUPPORT_REC", DB, adOpenStatic, adLockOptimistic
If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,chr(34),"")
arr = split(Line,"|")

ORS.AddNew
ORS("CUST_ID") = arr(0)
ORS("REFERENCE_NUM") = arr(1)
ORS("AMOUNT") = arr(2)
ORS("ACCOUNT_NUM") = arr(3)
ORS("RRN") = arr(4)
ORS("UNIQUE_REF") = arr(4)&arr(2)
ORS("BANK") = arr(5)
ORS("PAYMENT_ID") = arr(6)
ORS.Update
	
Loop
Next
ORS.Close
RS.Close
WScript.Echo ("SUPPORT: Files Upload Success")


DB.Close
Set DB = Nothing
Set RSO = Nothing
Set RS = Nothing
Set Line = Nothing
Set ORS = Nothing
Set arr = Nothing
Set fil = Nothing
Set FLD = Nothing
Set strFolder = Nothing

'''''''''''''''''' SUPPORT FILE UPLOADED '''''''''''''''''''''''''''''''''

Set objExcel = CreateObject("Excel.Application")
Set RSO = CreateObject("Scripting.FileSystemObject")
WScript.Echo ("BILLDESK: Reading Billdesk File")
strFolder = "C:\MCon_Recon\BilldeskFile\Recharge"
Set FLD = RSO.GetFolder(strFolder)

Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"

DB.Execute("Delete from BILLDESK_REC;")
'WScript.Sleep 10000
DB.Execute("Alter Table BILLDESK_REC Alter Column SERIAL_FIN COUNTER(1,1);")

For Each fil In FLD.Files
Set objBook = objExcel.Workbooks.Open(fil.Path)

objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objSheet = objBook.Worksheets("Sheet1")
objSheet.SaveAs "C:\MCon_Recon\BilldeskFile\Recharge\Recharge.csv", xlCSV


WScript.Echo ("BILLDESK: Excel converted to CSV")

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit

Set objBook = Nothing
Set objExcel = Nothing
Set objSheet = Nothing


WScript.Echo ("BILLDESK: Reading CSV File")

Set RS = RSO.OpenTextFile("C:\MCon_Recon\BilldeskFile\Recharge\Recharge.csv", ForReading)
Set ORS = CreateObject("ADODB.Recordset")

WScript.Echo ("BILLDESK: Ready to Upload Data")

If not RS.AtEndOfStream Then RS.Skipline
ORS.Open "BILLDESK_REC", DB, adOpenStatic, adLockOptimistic
Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,chr(34),"")
arr = split(Line,",")

If len(Line)>0 Then
ORS.AddNew
ORS("CUST_ID") = arr(0)
ORS("REFER_NUM") = arr(1)
ORS("AMOUNT") = arr(2)
ORS("RRN") = arr(4)
ORS("UNIQUE_REF") = arr(4)&arr(2)
ORS("BANK") = arr(5)
ORS("STATUS") = arr(6)
ORS.Update
End If

Loop

RS.Close
ORS.Close
Set RS = Nothing
Set ORS = Nothing

Next



WScript.Echo ("BILLDESK: Data Upload Success")

Set RS = DB.Execute("SELECT STATUS, Count(CUST_ID), Sum(AMOUNT) FROM BILLDESK_REC GROUP BY STATUS;")
WScript.Echo ("PARTICULARS					            | Count|    AMOUNT" & vbCrlf)
WScript.Echo ("------------------------------------------------------------------------------" & vbCrlf)
DO WHILE NOT RS.EOF
Diff = 60-Len(RS.Fields(0))
Diff2 = 6-Len(RS.Fields(1))
Diff3 = 10-Len(RS.Fields(2))
WScript.Echo (RS.Fields(0)&Space(Diff) & "|" &Space(Diff2)& RS.Fields(1) & "|" &Space(Diff3)& RS.Fields(2) & vbCrlf)
Result_Rec = RS.Fields(0)&Space(Diff) & "|" &Space(Diff2)& RS.Fields(1) & "|" &Space(Diff3)& RS.Fields(2) & vbCrlf
RS.MoveNext
Loop
RS.Close

WScript.Echo ("BILLDESK: Summary Data Generated")

If RSO.FileExists("C:\MCon_Recon\BilldeskFile\Recharge\Recharge.csv") Then
RSO.DeleteFile("C:\MCon_Recon\BilldeskFile\Recharge\Recharge.csv")
End If

DB.Close
Set DB = Nothing
Set RSO = Nothing
Set ORS = Nothing
Set RS = Nothing
Set WS = Nothing
Set Line = Nothing
Set Result_Rec = Nothing
Set Diff = Nothing
Set Diff2 = Nothing
Set Diff3 = Nothing
Set Result_Rec = Nothing
Set fil = Nothing
Set FLD = Nothing
Set strFolder = Nothing


'''''''''''''''''' BILLDESK FILE UPLOADED '''''''''''''''''''''''''''''''''

Set RSO = CreateObject("Scripting.FileSystemObject")
WScript.Echo ("FINACLE: Deleting Previous Data")

Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"

DB.Execute("Delete from FINACLE_REC;")
DB.Execute("Alter Table FINACLE_REC Alter Column SERIAL_FIN COUNTER(1,1);")

strFolder = "C:\MCon_Recon\FinacleFile\RECHARGE"
Set FLD = RSO.GetFolder(strFolder)

For Each fil In FLD.Files
Set RS = RSO.OpenTextFile(fil.Path, ForReading)
Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "FINACLE_REC", DB, adOpenStatic, adLockOptimistic
    
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

ORS.Close
RS.Close
Set RS = Nothing
Set ORS = Nothing

Next

WScript.Echo ("FINACLE: Data Upload Success")

'WScript.Echo ("DB Connection Success")
Set RS = DB.Execute("SELECT last(BALANCE) from FINACLE_REC where GL_DATE = '"& DatePr &"';")
OPBal = RS.Fields(0)
RS.Close
Set RS = Nothing
WScript.Sleep 5000
Set RS = DB.Execute("SELECT last(BALANCE) from FINACLE_REC where GL_DATE = '"& DateC &"';")
CLBal = RS.Fields(0)
RS.Close
Set RS = Nothing

WScript.Sleep 5000
DB.Execute("UPDATE FINACLE_REC SET ENTERED_BY = 'SYSTEM' WHERE PARTICULARS Like '%MBK/%';")
'WScript.Sleep 5000
DB.Execute("UPDATE FINACLE_REC SET ENTERED_BY = 'MANUAL' WHERE PARTICULARS Not Like '%MBK/%';")
'WScript.Sleep 5000
DB.Execute("UPDATE FINACLE_REC INNER JOIN BILLDESK_REC ON FINACLE_REC.UNIQUE_REF = BILLDESK_REC.UNIQUE_REF SET FINACLE_REC.BILLDESK_STATUS = BILLDESK_REC.STATUS;")
'WScript.Sleep 5000
DB.Execute("UPDATE FINACLE_REC INNER JOIN SUPPORT_REC ON FINACLE_REC.UNIQUE_REF = SUPPORT_REC.UNIQUE_REF SET FINACLE_REC.SUPPORT_STATUS = 'SUCCESS';")
'WScript.Sleep 5000
DB.Execute("UPDATE FINACLE_REC INNER JOIN BILLDESK_REC_TEMP ON FINACLE_REC.UNIQUE_REF = BILLDESK_REC_TEMP.UNIQUE_REF SET FINACLE_REC.BILLDESK_STATUS = BILLDESK_REC_TEMP.STATUS WHERE GL_DATE = '"& DateP &"';")
'WScript.Sleep 5000
DB.Execute("UPDATE FINACLE_REC INNER JOIN SUPPORT_REC_TEMP ON FINACLE_REC.UNIQUE_REF = BILLDESK_REC_TEMP.UNIQUE_REF SET FINACLE_REC.SUPPORT_STATUS = 'SUCCESS' WHERE GL_DATE = '"& DateP &"';")
'WScript.Sleep 5000
Set RS = DB.Execute("Select BILLDESK_STATUS, Count(*), SUM(Val(CREDIT_AMT)-Val(DEBIT_AMT)) from FINACLE_REC where GL_DATE between '" &DateP& "' AND '" &DateC& "' and ENTERED_BY = 'SYSTEM' and BILLDESK_STATUS is NOT NULL Group By BILLDESK_STATUS;")
WScript.Echo ("PARTICULARS					            | Count|    AMOUNT|		   BALANCE|" & vbCrlf)
WScript.Echo ("------------------------------------------------------------------------------" & vbCrlf)
WScript.Echo ("Opening Balance (Recharge): 								"& OPBal & vbCrlf)
DO WHILE NOT RS.EOF
Diff = 60-Len(RS.Fields(0))
Diff2 = 6-Len(RS.Fields(1))
Diff3 = 10-Len(RS.Fields(2))
OPBal = OPBal + RS.Fields(2)
WScript.Echo (RS.Fields(0)&Space(Diff) & "|" &Space(Diff2)& RS.Fields(1) & "|" &Space(Diff3)& RS.Fields(2) & "|		" & (OPBal) & vbCrlf)
RS.MoveNext
Loop
RS.Close
Set RS = Nothing
Set Diff = Nothing
Set Diff2 = Nothing
Set Diff3 = Nothing

WScript.Sleep 5000


Set RS = DB.Execute("Select BILLDESK_STATUS, PARTICULARS, Val(CREDIT_AMT)-Val(DEBIT_AMT) from FINACLE_REC where GL_DATE between '" &DateP& "' AND '" &DateC& "' and ENTERED_BY = 'MANUAL';")
DO WHILE NOT RS.EOF
Diff = 60-Len(RS.Fields(0))
Diff2 = 6-Len(RS.Fields(1))
Diff3 = 10-Len(RS.Fields(2))
OPBal = OPBal + RS.Fields(2)
WScript.Echo ("MANUAL ENTRY"&Space(Diff) & "|" &Space(Diff2)& RS.Fields(1) & "|" &Space(Diff3)& RS.Fields(2) & "|		" & (OPBal) & vbCrlf)
RS.MoveNext
Loop
RS.Close
Set RS = Nothing
Set Diff = Nothing
Set Diff2 = Nothing
Set Diff3 = Nothing


WScript.Echo ("Closing Balance (Recharge): 								"& CLBal & vbCrlf)
WScript.Echo ("FINACLE: Summary Report Generated")
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
Set DatePr = Nothing
Set DateC = Nothing

