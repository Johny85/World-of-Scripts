Option Explicit
On Error Resume Next

Const DeleteReadOnly = True
Dim objShl, DateF, objFSO, REX, DB, ORS, CMD
Dim ObjOutlook, SSession, Inbox, OSub, Item1, IntC, Atchm
Dim RSO, RS, WS, WSO, WSP
Dim arr, SLine
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

Do
DateF = inputbox ("Please select Date of Reconciliation in 'DD-MM-YYYY' format")

Set Rex = CreateObject("VBScript.RegExp")
Rex.Global = True
Rex.Pattern = "(0[1-9]|[12][0-9]|3[01])[-](0[1-9]|1[012])[-](19|20)\d\d"
If Rex.Test(DateF) Then
MsgBox ("Correct")

WScript.Echo ("Module Start Success")


Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists("C:\MConnect\Reports.zip") Then
objFSO.DeleteFile("C:\MConnect\Reports.zip")
objFSO.DeleteFile("C:\MConnect\Reports\*")
End If
objFSO.DeleteFile("C:\Users\PR172959\Documents\Testing\*")
objFSO.DeleteFile("C:\Users\PR172959\Documents\Testing\Recharge\*")
objFSO.DeleteFile("C:\Users\PR172959\Documents\Testing\BillPay\*")
Set objFSO = Nothing

WScript.Echo ("Previous Files Deleted")

WScript.Sleep 5000

Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")
Set Inbox = SSession.GetDefaultFolder(6).Folders("BOB Internal Mail")

For Each Item1 in Inbox.Items

OSub = Item1.Subject
OSub = Trim(Replace(OSub," ",""))

	If UCase(OSub) = "REPORTS|"&DateF Then
	
	IntC = Item1.Attachments.Count
			If IntC > 0 Then
			For Each Atchm In Item1.Attachments
			If UCase(Left(Atchm.FileName,7)) = "REPORTS" Then
			Atchm.SaveAsFile "C:\MConnect\Reports.zip"
			Set Atchm = Nothing
			End If
			Next
		End If
		End If
Next
Set OSub = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing

WScript.Echo ("Files downloaded from MailBox")

Set objShl = WScript.CreateObject ("WScript.shell")
objShl.run """C:\Program Files\WinRAR\WinRAR.exe"" X ""C:\MConnect\Reports.zip"" ""C:\MConnect"""
Set objShl = Nothing


WScript.Sleep 5000
WScript.Echo ("Files extraction completed")





Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

'Dim Tran_Date, Cust_Id, Particulars, Amount, Account, RRN

Set DB = WScript.CreateObject("ADODB.Connection")
Set ORS = CreateObject("ADODB.Recordset")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\PR172959\Documents\Pritimay\Database.accdb;"
ORS.Open "Recharge_Status", DB, adOpenStatic, adLockOptimistic

Set WS = WSO.CreateTextFile("C:\MConnect\Reports\Recharge_Ready.txt", ForWriting)
Set WSP = WSO.CreateTextFile("C:\MConnect\Reports\Recharge_DB.csv", ForWriting)
Set RS = RSO.OpenTextFile("C:\MConnect\Reports\Recharge.txt", ForReading)

Set CMD = CreateObject("ADODB.Command")
With CMD
.ActiveConnection = DB
.CommandText = "Delete from Recharge_Status where Tran_Date="&"""&DateF&"";"
End With
CMD.Execute
Set CMD = Nothing

			
			If not RS.AtEndOfStream Then RS.Skipline
			Do Until RS.AtEndOfStream
			SLine = RS.ReadLine
			SLine = Replace(SLine,chr(34),"")
			'WScript.Echo (SLine)
	   		arr = split(SLine,"|")

WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|XXXXXXXXXXXXXX|"& arr(4) &"|"& arr(5) & vbCrLf
WSP.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) & "|"& arr(4) &"|"& arr(5) &"|"& arr(6) & vbCrLf


ORS.AddNew
ORS("Tran_Date") = DateF
ORS("Cust_Id") = arr(0)
ORS("Reference") = arr(1)
ORS("Amount") = CDbl(arr(2))
ORS("Account") = arr(3)
ORS("RRN") = arr(4)
ORS.Update 
			
Loop

	RS.Close
	WS.Close
	WSP.Close
	ORS.Close
	
'Clean up
Set arr = Nothing
Set SLine = Nothing
Set RS = Nothing
Set WS = Nothing
Set WSP = Nothing
Set RSO = Nothing
Set WSO = Nothing
Set ORS = Nothing

WScript.Echo ("Recharge File processing completed")

WScript.Sleep 5000


Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")
Set ORS = CreateObject("ADODB.Recordset")
Set WS = WSO.CreateTextFile("C:\MConnect\Reports\Bill_Payment_Ready.txt", ForWriting)
Set WSP = WSO.CreateTextFile("C:\MConnect\Reports\Bill_Payment_DB.csv", ForWriting)
Set RS = RSO.OpenTextFile("C:\MConnect\Reports\Bill Payment.txt", ForReading)

Set CMD = CreateObject("ADODB.Command")
With CMD
.ActiveConnection = DB
.CommandText = "Delete from BillPay_Status where Tran_Date="&"""&DateF&"";"
End With
CMD.Execute
Set CMD = Nothing

ORS.Open "BillPay_Status", DB, adOpenStatic, adLockOptimistic

			
			If not RS.AtEndOfStream Then RS.Skipline
			Do Until RS.AtEndOfStream
			SLine = RS.ReadLine
			SLine = Replace(SLine,chr(34),"")
	   		arr = split(SLine,"|")

WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|XXXXXXXXXXXXXX|"& arr(4) &"|"& arr(5) & vbCrLf
WSP.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) & "|"& arr(4) &"|"& arr(5) & vbCrLf
			

ORS.AddNew
ORS("Tran_Date") = DateF
ORS("Cust_Id") = arr(0)
ORS("Reference") = arr(1)
ORS("Amount") = CDbl(arr(2))
ORS("Account") = arr(3)
ORS("RRN") = arr(4)
ORS.Update 

Loop

	RS.Close
	WS.Close
	WSP.Close
	ORS.Close
	DB.Close
	
Set DB = Nothing	
Set ORS = Nothing	
Set arr = Nothing
Set SLine = Nothing
Set RS = Nothing
Set WS = Nothing
Set WSP = Nothing
Set RSO = Nothing
Set WSO = Nothing

WScript.Echo ("Bill Payment File processing completed")

WScript.Sleep 5000


Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.CreateItem(0)
With SSession
'.To = "sanjay.margi@billdesk.com; narayan@billdesk.com; maheshj@billdesk.com; santosh.kudalkar@billdesk.com; ganesh.dalvi@billdesk.com; maheshgohil@billdesk.com; kaustubh@billdesk.com; krushnali.pawaskar@billdesk.com; ebpprecon@billdesk.com; ashishgupta@billdesk.com"
'.Cc = "abdul.rehman@billdesk.com; naveenujagiri@billdesk.com; hitesh@billdesk.com; yogesh.agare@billdesk.com; ashwini.chavan@billdesk.com"
.Subject = "Bill pay recon file - New Mobile Banking "&DateF
.Attachments.Add "C:\MConnect\Reports\Bill_Payment_Ready.txt"
.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "We are attaching herewith report for Bill Pay through New Mobile Banking of " & DateF & "." & vbCrLf & vbCrLf & "Arrange to share file for all successful bill pay transactions " & DateF & " and for all failed bill pay transactions. (file should be in xls format)" & vbCrLf & vbCrLf & "Raise claim for all successful bill pay transactions and this should tally with the sum of figures given in above asked successful transactions file." & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Srikanth Reddy Alluri" & vbCrLf & "Officer, Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking, Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India"
End With

SSession.Send

Set ObjOutlook = Nothing
Set SSession = Nothing



WScript.Sleep 5000


Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.CreateItem(0)
With SSession
'.To = "sanjay.margi@billdesk.com; narayan@billdesk.com; maheshj@billdesk.com; santosh.kudalkar@billdesk.com; ganesh.dalvi@billdesk.com; maheshgohil@billdesk.com; kaustubh@billdesk.com; krushnali.pawaskar@billdesk.com; ebpprecon@billdesk.com; ashishgupta@billdesk.com"
'.Cc = "abdul.rehman@billdesk.com; naveenujagiri@billdesk.com; hitesh@billdesk.com; yogesh.agare@billdesk.com; ashwini.chavan@billdesk.com"
.Subject = "Recharge - New Mobile Banking "&DateF
.Attachments.Add "C:\MConnect\Reports\Recharge_Ready.txt"
.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "We are attaching herewith report for Recharge through New Mobile Banking of " & DateF &"," & vbCrLf & vbCrLf & "Arrange to share file for all successful recharge transactions " & DateF & " and for all failed bill pay transactions. (file should be in xls format)" & vbCrLf & vbCrLf & "Raise claim for all successful recharge transactions and this should tally with the sum of figures given in above asked successful transactions file." & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Srikanth Reddy Alluri" & vbCrLf & "Officer, Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking, Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf
End With

SSession.Send

Set ObjOutlook = Nothing
Set SSession = Nothing


WScript.Echo ("Mails sent successfully to BillDesk")
WScript.Sleep 2000
WScript.Echo ("Module Execution Success")

Else
MsgBox ("InCorrect Date Entered")
End If
Loop While Rex.Test(DateF) = False
