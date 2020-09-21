Option Explicit
'On Error Resume Next

Dim DateF, RS, WS, RSO, WSO, SLine, arr, ObjOutlook, SSession

Const ForReading = 1, ForWriting = 2, ForAppending = 8 
DateF = Date()-1

Set RSO = CreateObject("Scripting.FileSystemObject")
Set WSO = CreateObject("Scripting.FileSystemObject")

Set ObjOutlook = CreateObject("Outlook.Application")
			
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.AHMEDABAD@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Ahmedabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\AHMEDABAD.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing


			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.BARODA@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Baroda Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\BARODA.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



	
			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.BENGALURU@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Bengaluru Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\BENGALURU.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.BHOPAL@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Bhopal Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\BHOPAL.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "zdm.zochd@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Chandigarh Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\CHANDIGARH.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.CHENNAI@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for CHENNAI ZONE"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\CHENNAI.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing


			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "zdm.ekmz@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Ernakulam Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\ERNAKULAM.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing


			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.Hyderabad@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Hyderabad Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\HYDERABAD.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing


			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.JAIPUR@bankofbaroda.com;IT.RZ@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Jaipur Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\JAIPUR.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.KOLKATA@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Kolkata Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\KOLKATA.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.LUCKNOW@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Lucknow Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\LUCKNOW.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "it.zomglr@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Mangaluru Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\MANGALURU.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.BAREILLY@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\MEERUT.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.MUMBAI@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Mumbai Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\MUMBAI.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.NEWDELHI@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Delhi Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\NEWDELHI.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.PATNA@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Patna Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\PATNA.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing


			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "ZDM.PUNE@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Pune Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\PUNE.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing



			Set SSession = ObjOutlook.CreateItem(0)
			With SSession
			.To = "zdm.zorajkot@bankofbaroda.com"
			.Cc = "KIRUBANANDAN.K@bankofbaroda.com"
			.Subject = "List of Eligible Accounts from eDENA Mobile Banking for Rajkot Zone"
			.Attachments.Add "C:\Users\PR172959\Documents\Pritimay\Dena Bank Data\RAJKOT.rar"
			.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Kindly find attached list of users eligible for registration to Mobile Banking." & vbCrLf & vbCrLf & "Note: This data is the list of eligible users only and thus progress report won't be available on daily basis" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Mobile Banking" & vbCrLf & "Fintech, Partnerships & Mobile Banking Dept." & vbCrLf & "Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India" & vbCrLf & "022 67592577"
			End With
			
SSession.Send
Set SSession = Nothing


