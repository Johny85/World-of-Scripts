Option Explicit
On Error Resume Next

Dim RSO, DB, RS, ORS, Line, arr

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 


Set RSO = CreateObject("Scripting.FileSystemObject")
Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"

DB.Execute("Delete from BILLDESK_REC_TEMP;")
'WScript.Sleep 10000
DB.Execute("Alter Table BILLDESK_REC_TEMP Alter Column SERIAL_FIN COUNTER(1,1);")
'WScript.Sleep 5000
DB.Execute("Delete from SUPPORT_REC_TEMP;")
'WScript.Sleep 10000
DB.Execute("Alter Table SUPPORT_REC_TEMP Alter Column SERIAL_SUP COUNTER(1,1);")

WScript.Echo ("BILLDESK: Reading CSV File")

Set RS = RSO.OpenTextFile("C:\MCon_Recon\BilldeskFile\Prev.csv", ForReading)
Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "BILLDESK_REC_TEMP", DB, adOpenStatic, adLockOptimistic

WScript.Echo ("BILLDESK: Ready to Upload Data")

If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,chr(34),"")
arr = split(Line,",")

ORS.AddNew
ORS("CUST_ID") = arr(0)
ORS("REFER_NUM") = arr(1)
ORS("AMOUNT") = arr(2)
ORS("RRN") = arr(4)
ORS("UNIQUE_REF") = arr(4)&arr(2)
ORS("BANK") = arr(5)
ORS("STATUS") = arr(6)
ORS.Update

Loop

RS.Close
ORS.Close
Set RS = Nothing
Set ORS = Nothing
Set Line = Nothing
Set arr = Nothing


Set RS = RSO.OpenTextFile("C:\MCon_Recon\SupportFile\Recharge.txt", ForReading)
Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "SUPPORT_REC_TEMP", DB, adOpenStatic, adLockOptimistic

WScript.Echo ("SUPPORT: Ready to Upload Data")
If not RS.AtEndOfStream Then RS.Skipline
Do Until RS.AtEndOfStream
Line = RS.ReadLine
Line = Replace(Line,chr(34),"")
arr = split(Line,"|")

ORS.AddNew
ORS("CUST_ID") = arr(0)
ORS("REFER_NUM") = arr(1)
ORS("AMOUNT") = arr(2)
ORS("RRN") = arr(4)
ORS("UNIQUE_REF") = arr(4)&arr(2)
ORS("BANK") = arr(5)
ORS("STATUS") = arr(6)
ORS.Update

Loop

RS.Close
ORS.Close
Set RS = Nothing
Set ORS = Nothing



WScript.Echo ("BILLDESK: Data Upload Success")
DB.Close
Set DB = Nothing
Set Line = Nothing
Set arr = Nothing
