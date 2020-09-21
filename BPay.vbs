Option Explicit
On Error Resume Next

Dim objBook, objExcel, objSheet, RSOO, strFolder, FLD, fil, DB, ORS, BBPS_ID, RSO, CUST_ID, Line, arr, RS, Diff, Diff2, Diff3

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 
Const xlCSV = 6

Set objExcel = CreateObject("Excel.Application")
Set RSOO = CreateObject("Scripting.FileSystemObject")

strFolder = "C:\MCon_Recon\BilldeskFile\BillPay"
Set FLD = RSOO.GetFolder(strFolder)

For Each fil In FLD.Files
Set objBook = objExcel.Workbooks.Open(fil.Path)

objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objSheet = objBook.Worksheets("Sheet1")
objSheet.SaveAs "C:\MCon_Recon\BilldeskFile\BillPay\Billpay.csv", xlCSV
Next

objExcel.ActiveWorkbook.Save = False
objExcel.ActiveWorkbook.Close = True
objExcel.Quit

Set objBook = Nothing
Set objExcel = Nothing
Set objSheet = Nothing
Set RSOO = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set RSO = CreateObject("Scripting.FileSystemObject")

Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"

DB.Execute("Delete from BILLDESK_BP;")
WScript.Sleep 10000
DB.Execute("Alter Table BILLDESK_BP Alter Column SERIAL_FIN COUNTER(1,1);")

Set RS = DB.Execute("Alter Table BILLDESK_BP Alter Column SERIAL_FIN COUNTER(1,1)")
Set RS = Nothing

Set RS = RSO.OpenTextFile("C:\MCon_Recon\BilldeskFile\BillPay\Billpay.csv", ForReading)
Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "BILLDESK_BP", DB, adOpenStatic, adLockOptimistic

If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,chr(34),"")
arr = split(Line,",")
CUST_ID = arr(1)
BBPS_ID = arr(7)
If Cust_ID <> "" Then
If (Left(BBPS_ID,4) = "BBPS") Then
ORS.AddNew
ORS("CUST_ID") = arr(0)
ORS("REFER_NUM") = arr(1)
ORS("AMOUNT") = arr(2)
ORS("RRN") = arr(4)
ORS("BANK") = arr(5)
ORS("BBPS_ID") = arr(6)
ORS("STATUS") = arr(7)
ORS.Update
Else
ORS.AddNew
ORS("CUST_ID") = arr(0)
ORS("REFER_NUM") = arr(1)
ORS("AMOUNT") = arr(2)
ORS("RRN") = arr(4)
ORS("UNIQUE_REF") = arr(4)&arr(2)
ORS("BANK") = arr(5)
ORS("BBPS_ID") = arr(7)
ORS("STATUS") = arr(6)
ORS.Update

End If
End If
Loop

RS.Close
ORS.Close
Set RS = Nothing

Set RS = DB.Execute("SELECT STATUS, Count(CUST_ID), Sum(AMOUNT) FROM BILLDESK_BP GROUP BY STATUS;")
WScript.Echo ("PARTICULARS					            | Count|    AMOUNT" & vbCrlf)
WScript.Echo ("------------------------------------------------------------------------------" & vbCrlf)
DO WHILE NOT RS.EOF
Diff = 60-Len(RS.Fields(0))
Diff2 = 6-Len(RS.Fields(1))
Diff3 = 10-Len(RS.Fields(2))
WScript.Echo (RS.Fields(0)&Space(Diff) & "|" &Space(Diff2)& RS.Fields(1) & "|" &Space(Diff3)& RS.Fields(2) & vbCrlf)
RS.MoveNext
Loop
RS.Close
DB.Close

If RSO.FileExists("C:\MCon_Recon\BilldeskFile\BillPay\Billpay.csv") Then
RSO.DeleteFile("C:\MCon_Recon\BilldeskFile\BillPay\Billpay.csv")
End If

Set RSO = Nothing
Set ORS = Nothing
Set RS = Nothing
Set WS = Nothing
Set Line = Nothing
Set DB = Nothing

