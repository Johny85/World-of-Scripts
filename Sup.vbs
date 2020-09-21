Option Explicit
On Error Resume Next

Dim RSO, RS, Line, arr, DB, ORS

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

Set RSO = CreateObject("Scripting.FileSystemObject")

Set RS = RSO.OpenTextFile("C:\MConnect\Reports\Recharge.txt", ForReading)

Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"
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
ORS("BANK") = arr(5)
ORS("PAYMENT_ID") = arr(6)
ORS.Update
	
Loop

ORS.Close
RS.Close
DB.Close

Set RSO = Nothing
Set RS = Nothing
Set Line = Nothing
Set ORS = Nothing
Set DB = Nothing
Set Line = Nothing
Set arr = Nothing



''''''''''''''''''''''''''''''BILLDESK''''''''''''''''''

Set RSO = CreateObject("Scripting.FileSystemObject")

Set RS = RSO.OpenTextFile("C:\MConnect\Reports\Bill Payment.txt", ForReading)

Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"
Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "SUPPORT_BP", DB, adOpenStatic, adLockOptimistic

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
ORS("BANK") = arr(5)
ORS("PAYMENT_ID") = arr(6)
ORS.Update
	
Loop

ORS.Close
RS.Close
DB.Close

Set RSO = Nothing
Set RS = Nothing
Set Line = Nothing
Set ORS = Nothing
Set DB = Nothing
Set Line = Nothing
Set arr = Nothing