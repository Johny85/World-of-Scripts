Option Explicit
On Error Resume Next

Dim xlApp, DateC, ObjOutlook, SSession


Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

'WScript.Sleep 12000

DateC = Date()-1

Set xlApp = CreateObject("Excel.Application")
xlApp.Workbooks.Open("E:\Mconnect Plus\MIS\"&DateC&".xlsx")

xlApp.ActiveSheet.Range("A1:O39").ExportAsFixedFormat 0, "E:\Mconnect Plus\PDF\"&DateC&".pdf", 0, 1, 0,,,0


xlApp.ActiveWorkbook.Save
xlApp.ActiveWorkbook.Close
xlApp.Application.Quit
		
		
Set xlApp = Nothing


Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.CreateItem(0)
With SSession
.To = "KIRUBANANDAN.K@bankofbaroda.com; mconnect@bankofbaroda.com; ARUN.KUMAR17@bankofbaroda.co.in"
'.To = "mobility@bankofbaroda.co.in"
.Subject = "ChannelWise Daily MIS Report for Dated "&DateC
.Attachments.Add "E:\Mconnect Plus\MIS\"&DateC&".xlsx"
.Attachments.Add "E:\Mconnect Plus\PDF\"&DateC&".pdf"
.Body = "Dear Sir/Madam," & vbCrLf & vbCrLf & "Please find attached ChannelWise Daily MIS Report for Dated " & DateC & "." & vbCrLf & vbCrLf & vbCrLf & "Thanks and Regards" & vbCrLf & "Pritimay" & vbCrLf & "Fintech, Partnerships & Mobile Banking, Bank of Baroda" & vbCrLf & "Baroda Sun Tower, 6th Floor, C-34, G Block, BKC, Bandra (E), Mumbai - 400 051, India"

End With

SSession.Send

Set ObjOutlook = Nothing
Set SSession = Nothing
Set DateC = Nothing