Dim ObjOutlook, SSession, DateF


Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.CreateItem(0)
With SSession
'.To = "sanjay.margi@billdesk.com; narayan@billdesk.com; maheshj@billdesk.com; santosh.kudalkar@billdesk.com; ganesh.dalvi@billdesk.com; maheshgohil@billdesk.com; kaustubh@billdesk.com; krushnali.pawaskar@billdesk.com; ebpprecon@billdesk.com; ashishgupta@billdesk.com"
'.Cc = "abdul.rehman@billdesk.com; naveenujagiri@billdesk.com; hitesh@billdesk.com; yogesh.agare@billdesk.com; ashwini.chavan@billdesk.com"
.To = "john.rock85@gmail.com; pritimay_83@yahoo.co.in"
.Cc = "john.rock85@gmail.com; pritimay_83@yahoo.co.in"
.Subject = "Bill pay recon file - New Mobile Banking "&DateF
.Attachments.Add "C:\MConnect\Reports\Bill_Payment_Ready.txt"
.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "We are attaching herewith report for Bill Pay through New Mobile Banking of " & DateF& "." & vbCrLf & vbCrLf & "Arrange to share file for all successful bill pay transactions " & DateF & " and for all failed bill pay transactions. (file should be in xls format)" & vbCrLf & vbCrLf & "Raise claim for all successful bill pay transactions and this should tally with the sum of figures given in above asked successful transactions file." & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Srikanth Reddy Alluri" & vbCrLf & "Officer, Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking, Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India"
End With

SSession.Send

Set ObjOutlook = Nothing
Set SSession = Nothing



Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.CreateItem(0)
With SSession
'.To = "sanjay.margi@billdesk.com; narayan@billdesk.com; maheshj@billdesk.com; santosh.kudalkar@billdesk.com; ganesh.dalvi@billdesk.com; maheshgohil@billdesk.com; kaustubh@billdesk.com; krushnali.pawaskar@billdesk.com; ebpprecon@billdesk.com; ashishgupta@billdesk.com"
'.Cc = "abdul.rehman@billdesk.com; naveenujagiri@billdesk.com; hitesh@billdesk.com; yogesh.agare@billdesk.com; ashwini.chavan@billdesk.com"
.To = "john.rock85@gmail.com; pritimay_83@yahoo.co.in"
.Cc = "john.rock85@gmail.com; pritimay_83@yahoo.co.in"
.Subject = "Recharge - New Mobile Banking "&DateF
.Attachments.Add "C:\MConnect\Reports\Recharge_Ready.txt"
.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "We are attaching herewith report for Recharge through New Mobile Banking of " & DateF &"," & vbCrLf & vbCrLf & "Arrange to share file for all successful recharge transactions " & DateF & " and for all failed bill pay transactions. (file should be in xls format)" & vbCrLf & vbCrLf & "Raise claim for all successful recharge transactions and this should tally with the sum of figures given in above asked successful transactions file." & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Srikanth Reddy Alluri" & vbCrLf & "Officer, Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking, Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf
End With

SSession.Send

Set ObjOutlook = Nothing
Set SSession = Nothing



