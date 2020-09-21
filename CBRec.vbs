Option Explicit

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

Dim FSO, OTF, ORS, DB, CMD, SNL, NDate, Tran_ID, CAmount, DAmount, Tran_Type, Balance, RRN, RS, OBal, CBal, Rex, oFldr, ofile
Dim InDate

Set DB = WScript.CreateObject("ADODB.Connection")
Set ORS = CreateObject("ADODB.Recordset")
Set CMD = CreateObject("ADODB.Command")


Set FSO = CreateObject("Scripting.FileSystemObject")

Do
InDate = inputbox ("Please select Date of Reconciliation in 'DD-MM-YYYY' format")

Set Rex = CreateObject("VBScript.RegExp")
Rex.Global = True
Rex.Pattern = "(0[1-9]|[12][0-9]|3[01])[-](0[1-9]|1[012])[-](19|20)\d\d"
If Rex.Test(InDate) Then
MsgBox ("Starting Module")
Else
MsgBox ("InCorrect Date Entered")
End If
Loop While Rex.Test(InDate) = False

Set oFldr = FSO.getfolder("C:\Users\PR172959\Documents\Testing")

For Each ofile In oFldr.Files
  If lcase(FSO.GetExtensionName(ofile.Name)) = "rpt" Then
    ofile.Name = "CBS_Finacle.txt"
  End If
Next
WScript.Echo "Reading Pool Account Statement"
Set FSO = Nothing


Set FSO = CreateObject("Scripting.FileSystemObject")
Set OTF = FSO.OpenTextFile ("C:\Users\PR172959\Documents\Testing\CBS_Finacle.txt", ForReading)

DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\PR172959\Documents\Pritimay\Database.accdb;"

With CMD
.ActiveConnection = DB
.CommandText = "Delete from Recharge_CBS"
End With
CMD.Execute
WScript.Echo "DB Data Cleared"
Set CMD = Nothing


ORS.Open "Recharge_CBS", DB, adOpenStatic, adLockOptimistic 

Do Until OTF.AtEndOfStream
	SNL = OTF.ReadLine
	If Mid(SNL,39,4) = "MBK/" OR Mid(SNL,39,4) = "NEFT" OR Mid(SNL,39,4) = "MOB " Then
				NDate = Mid(SNL,2,10)
				Tran_ID = Mid(SNL,13,9)
				RRN = Mid(SNL,43,12)
				Tran_Type = Mid(SNL,39,4)
				DAmount = Replace((Mid(SNL,69,15))," ","0")
				CAmount = Replace((Mid(SNL,90,15))," ","0")
				Balance = Replace((Mid(SNL,115,15))," ","0")
	
					ORS.AddNew
					ORS("NDate") = Trim(NDate)
					ORS("Tran_ID") = Trim(Tran_ID)
					ORS("RRN") = Trim(RRN)
					ORS("Tran_Type") = Trim(Tran_Type)
					ORS("DAmount") = DAmount*1
					ORS("CAmount") = CAmount*1
					ORS("Balance") = Balance*1
					ORS.Update 

End If
Loop
ORS.Close
Set ORS = Nothing


Set RS = DB.Execute("select Top 1 Balance from Recharge_CBS where NDate = '" & CDate(InDate)-1 &"' order by Serial Desc")

DO WHILE NOT RS.EOF
OBal = RS.Fields(0)
RS.MoveNext
Loop
Set RS = Nothing

Set RS = DB.Execute("select Top 1 Balance from Recharge_CBS where NDate = '" & InDate &"' order by Serial Desc")

DO WHILE NOT RS.EOF
CBal = RS.Fields(0)
RS.MoveNext
Loop
Set RS = Nothing

WScript.Echo ("Opening Balance: " &OBal)
WScript.Echo ("Closing Balance: " &CBal)




DB.Close
OTF.Close
Set FSO = Nothing
Set OTF = Nothing
Set SNL = Nothing
Set oFldr = Nothing
Set ofile = Nothing

