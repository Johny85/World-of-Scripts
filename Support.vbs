Option Explicit
On Error Resume Next

Dim RSO, RS, Line, arr, DB, ORS, strFolder, FLD, fil

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

Set RSO = CreateObject("Scripting.FileSystemObject")
Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"
Set RS = DB.Execute("Alter Table SUPPORT_REC Alter Column SERIAL_SUP COUNTER(1,1)")
Set RS = Nothing

strFolder = "C:\MCon_Recon\SupportFile\Recharge"
Set FLD = RSO.GetFolder(strFolder)
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
DB.Close

Set RSO = Nothing
Set RS = Nothing
Set Line = Nothing
Set ORS = Nothing
Set DB = Nothing
Set Line = Nothing
Set arr = Nothing
Set FLD = Nothing



''''''''''''''''''''''''''''''BILLDESK''''''''''''''''''

Set RSO = CreateObject("Scripting.FileSystemObject")

Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\MCon_Recon\MCon_Recon.accdb;"
Set RS = DB.Execute("Alter Table SUPPORT_BP Alter Column SERIAL_SUP COUNTER(1,1)")
Set RS = Nothing

strFolder = "C:\MCon_Recon\SupportFile\BillPay"
Set FLD = RSO.GetFolder(strFolder)
For Each fil In FLD.Files

Set RS = RSO.OpenTextFile(fil.path, ForReading)
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
ORS("UNIQUE_REF") = arr(4)&arr(2)
ORS("BANK") = arr(5)
ORS("PAYMENT_ID") = arr(6)
ORS.Update
	
Loop
Next
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