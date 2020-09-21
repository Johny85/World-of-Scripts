Option Explicit
On Error Resume Next

Const DeleteReadOnly = True
Dim objShl, DateF, objFSO
Dim ObjOutlook, SSession, Inbox, OSub, Item1, IntC, Atchm
Dim RSO, RS, WS, WSO
Dim arr, SLine
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

WScript.Echo ("Module Start Success")

Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists("C:\MConnect\Reports.zip") Then
objFSO.DeleteFile("C:\MConnect\Reports.zip"), DeleteReadOnly
objFSO.DeleteFile("C:\MConnect\Reports\*"), DeleteReadOnly
End If
Set objFSO = Nothing

WScript.Echo ("Previous Files Deleted")

WScript.Sleep 10000

DateF = Date()-1

Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")
Set Inbox = SSession.GetDefaultFolder(6).Folders("BOB Internal Mail")

For Each Item1 in Inbox.Items

OSub = Item1.Subject

	If Left(OSub,7) = "Reports" AND Right(OSub,11) = " "&(DateF) Then
	IntC = Item1.Attachments.Count
			If IntC > 0 Then
			For Each Atchm In Item1.Attachments
			If Left(Atchm.FileName,7) = "Reports" Then
			Atchm.SaveAsFile "C:\MConnect\" & Atchm.FileName
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


WScript.Sleep 10000
WScript.Echo ("Files extraction completed")


Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

Set WS = WSO.CreateTextFile("C:\MConnect\Reports\Recharge_Ready.txt", ForWriting)
Set RS = RSO.OpenTextFile("C:\MConnect\Reports\Recharge.txt", ForReading)
			
			If not RS.AtEndOfStream Then RS.Skipline
			Do Until RS.AtEndOfStream
			SLine = RS.ReadLine
			SLine = Replace(SLine,chr(34),"")
			'WScript.Echo (SLine)
	   		arr = split(SLine,"|")

'WScript.Echo arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &vbCrLf

WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|XXXXXXXXXXXXXX|"& arr(4) &"|"& arr(5) & vbCrLf	
			
Loop

		'Close the file
		RS.Close
	WS.Close
	
'Clean up
Set arr = Nothing
Set SLine = Nothing
Set RS = Nothing
Set WS = Nothing
Set RSO = Nothing
Set WSO = Nothing

WScript.Echo ("Recharge File processing completed")

WScript.Sleep 10000


Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

Set WS = WSO.CreateTextFile("C:\MConnect\Reports\Bill_Payment_Ready.txt", ForWriting)
Set RS = RSO.OpenTextFile("C:\MConnect\Reports\Bill Payment.txt", ForReading)
			
			If not RS.AtEndOfStream Then RS.Skipline
			Do Until RS.AtEndOfStream
			SLine = RS.ReadLine
			SLine = Replace(SLine,chr(34),"")
			'WScript.Echo (SLine)
	   		arr = split(SLine,"|")

'WScript.Echo arr(0) &"|"& arr(1) &"|"& arr(2) &"|"& arr(3) &"|"& arr(4) &"|"& arr(5) &vbCrLf

WS.Write arr(0) &"|"& arr(1) &"|"& arr(2) &"|XXXXXXXXXXXXXX|"& arr(4) &"|"& arr(5) & vbCrLf	
			
Loop

		'Close the file
		RS.Close
	WS.Close
	
'Clean up
Set arr = Nothing
Set SLine = Nothing
Set RS = Nothing
Set WS = Nothing
Set RSO = Nothing
Set WSO = Nothing

WScript.Echo ("Bill Payment File processing completed")

WScript.Sleep 10000


Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.CreateItem(0)
With SSession
.To = "sanjay.margi@billdesk.com; narayan@billdesk.com; maheshj@billdesk.com; santosh.kudalkar@billdesk.com; ganesh.dalvi@billdesk.com; maheshgohil@billdesk.com; kaustubh@billdesk.com; krushnali.pawaskar@billdesk.com; ebpprecon@billdesk.com; ashishgupta@billdesk.com"
.Cc = "abdul.rehman@billdesk.com; naveenujagiri@billdesk.com; hitesh@billdesk.com; yogesh.agare@billdesk.com; ashwini.chavan@billdesk.com"
.Subject = "Bill pay recon file - New Mobile Banking "&DateF
.Attachments.Add "C:\MConnect\Reports\Bill_Payment_Ready.txt"
.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "We are attaching herewith report for Bill Pay through New Mobile Banking of " & DateF& "." & vbCrLf & vbCrLf & "Arrange to share file for all successful bill pay transactions " & DateF & " and for all failed bill pay transactions. (file should be in xls format)" & vbCrLf & vbCrLf & "Raise claim for all successful bill pay transactions and this should tally with the sum of figures given in above asked successful transactions file." & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Srikanth Reddy Alluri" & vbCrLf & "Officer, Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking, Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India"
End With

SSession.Send

Set ObjOutlook = Nothing
Set SSession = Nothing



WScript.Sleep 10000


Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.CreateItem(0)
With SSession
.To = "sanjay.margi@billdesk.com; narayan@billdesk.com; maheshj@billdesk.com; santosh.kudalkar@billdesk.com; ganesh.dalvi@billdesk.com; maheshgohil@billdesk.com; kaustubh@billdesk.com; krushnali.pawaskar@billdesk.com; ebpprecon@billdesk.com; ashishgupta@billdesk.com"
.Cc = "abdul.rehman@billdesk.com; naveenujagiri@billdesk.com; hitesh@billdesk.com; yogesh.agare@billdesk.com; ashwini.chavan@billdesk.com"
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