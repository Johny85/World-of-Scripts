Option Explicit
On Error Resume Next

Dim xlApp, xlBook, xlSheet, TS, intRow, strcom, DB, ORS
Dim DateF, DateC, ObjOutlook, SSession, Item1, Atchm, Inbox, OSub, OFrm, ODate, IntC

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 
Const DeleteReadOnly = True



DateF = Date()-1
DateC = Date()


Set ObjOutlook = CreateObject("Outlook.Application")
Set SSession = ObjOutlook.GetNameSpace("MAPI")
Set Item1 = CreateObject("Outlook.Application")
Set Atchm = CreateObject("Outlook.Application")
Set Inbox = SSession.GetDefaultFolder(6).Folders("BOB Internal Mail")

For Each Item1 in Inbox.Items

OSub = Item1.Subject
OFrm = Item1.Sender.GetExchangeUser().PrimarySmtpAddress
ODate = Left(Item1.ReceivedTime,10)

If OFrm&ODate  = "mconnect.ito@bankofbaroda.com"&DateC Then
IntC = Item1.Attachments.Count
If IntC > 0 Then
For Each Atchm In Item1.Attachments
If Right(Atchm.FileName,12) = Replace(DateF,"-","")&".xls" OR Right(Atchm.FileName,13) = Replace(DateF,"-","")&".xlsx" Then
Atchm.SaveAsFile "E:\Mconnect Plus\Dashboard\MIS.xls"
Set Atchm = Nothing
End If
Next
End If
End If
Next

Set OSub = Nothing
Set OFrm = Nothing
Set ODate = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing
Set DateF = Nothing
Set DateC = Nothing






intRow = 1
Set xlApp = CreateObject("Excel.Application")
Call xlApp.Workbooks.Open("E:\Mconnect Plus\Dashboard\MIS.xls", ForReading)
Set TS = xlApp.ActiveWorkbook.Worksheets("NOV")

DateC= Date()-1 


Select Case Mid(DateC,4,7)
Case "01-2019"
DateC = Left(DateC,2)&"-Jan-19"
Case "02-2019"
DateC = Left(DateC,2)&"-Feb-19"
Case "03-2019"
DateC = Left(DateC,2)&"-Mar-19"
Case "04-2019"
DateC = Left(DateC,2)&"-Apr-19"
Case "05-2019"
DateC = Left(DateC,2)&"-May-19"
Case "06-2019"
DateC = Left(DateC,2)&"-Jun-19"
Case "07-2019"
DateC = Left(DateC,2)&"-Jul-19"
Case "08-2019"
DateC = Left(DateC,2)&"-Aug-19"
Case "09-2019"
DateC = Left(DateC,2)&"-Sep-19"
Case "10-2019"
DateC = Left(DateC,2)&"-Oct-19"
Case "11-2019"
DateC = Left(DateC,2)&"-Nov-19"
Case "12-2019"
DateC = Left(DateC,2)&"-Dec-19"
End Select


Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\Mconnect Plus\MConnect_Plus.accdb;"
Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "DAILY_MIS", DB, adOpenStatic, adLockOptimistic

Do While intRow < TS.UsedRange.Rows.Count
set strcom = Nothing
strcom = StrComp(DateC,TS.Cells(intRow, 1).Text)
If strcom = 0 Then

ORS.AddNew
ORS("DATEF") = TS.Cells(intRow,1).Value
ORS("SBCAODSS") = TS.Cells(intRow,2).Value
ORS("SBCAODTD") = TS.Cells(intRow,3).Value
ORS("SBCAODBD") = TS.Cells(intRow,4).Value
ORS("PPFSS") = TS.Cells(intRow,5).Value
ORS("PPFTD") = TS.Cells(intRow,6).Value
ORS("PPFBD") = TS.Cells(intRow,7).Value
ORS("LOANSS") = TS.Cells(intRow,8).Value
ORS("LOANTD") = TS.Cells(intRow,9).Value
ORS("LOANBD") = TS.Cells(intRow,10).Value
ORS("SSFTSS") = TS.Cells(intRow,11).Value
ORS("SSFTTD") = TS.Cells(intRow,12).Value
ORS("SSFTBD") = TS.Cells(intRow,13).Value
ORS("TPFTSS") = TS.Cells(intRow,14).Value
ORS("TPFTTD") = TS.Cells(intRow,15).Value
ORS("TPFTBD") = TS.Cells(intRow,16).Value
ORS("IMPSSS") = TS.Cells(intRow,17).Value
ORS("IMPSTDBEN") = TS.Cells(intRow,18).Value
ORS("IMPSTDREM") = TS.Cells(intRow,19).Value
ORS("IMPSBD") = TS.Cells(intRow,20).Value
ORS("NEFTSS") = TS.Cells(intRow,21).Value
ORS("NEFTTD") = TS.Cells(intRow,22).Value
ORS("NEFTBD") = TS.Cells(intRow,23).Value
ORS("MRECHARGESS") = TS.Cells(intRow,24).Value
ORS("MRECHARGETD") = TS.Cells(intRow,25).Value
ORS("MRECHARGEBD") = TS.Cells(intRow,26).Value
ORS("BBPSDRECHARGESS") = TS.Cells(intRow,27).Value
ORS("BBPSDRECHARGETD") = TS.Cells(intRow,28).Value
ORS("BBPSDRECHARGEBD") = TS.Cells(intRow,29).Value
ORS("NBBPSDRECHARGESS") = TS.Cells(intRow,30).Value
ORS("NBBPSDRECHARGETD") = TS.Cells(intRow,31).Value
ORS("NBBPSDRECHARGEBD") = TS.Cells(intRow,32).Value
ORS("REGBILLPAYSS") = TS.Cells(intRow,33).Value
ORS("REGBILLPAYTD") = TS.Cells(intRow,34).Value
ORS("REGBILLPAYBD") = TS.Cells(intRow,35).Value
ORS("QBILLPAYSS") = TS.Cells(intRow,36).Value
ORS("QBILLPAYTD") = TS.Cells(intRow,37).Value
ORS("QBILLPAYBD") = TS.Cells(intRow,38).Value
ORS("TBBPSPAYSS") = TS.Cells(intRow,39).Value
ORS("TBBPSPAYTD") = TS.Cells(intRow,40).Value
ORS("TBBPSPAYBD") = TS.Cells(intRow,41).Value
ORS("CCPAYSS") = TS.Cells(intRow,42).Value
ORS("CCPAYTD") = TS.Cells(intRow,43).Value
ORS("CCPAYBD") = TS.Cells(intRow,44).Value
ORS("FDOPENSS") = TS.Cells(intRow,45).Value
ORS("FDOPENTD") = TS.Cells(intRow,46).Value
ORS("FDOPENBD") = TS.Cells(intRow,47).Value
ORS("RDOPENSS") = TS.Cells(intRow,48).Value
ORS("RDOPENTD") = TS.Cells(intRow,49).Value
ORS("RDOPENBD") = TS.Cells(intRow,50).Value
ORS("STPAYSS") = TS.Cells(intRow,51).Value
ORS("STPAYTD") = TS.Cells(intRow,52).Value
ORS("STPAYBD") = TS.Cells(intRow,53).Value
ORS("COMI") = TS.Cells(intRow,54).Value
ORS("COMSETTLED") = TS.Cells(intRow,55).Value
ORS("COMTD") = TS.Cells(intRow,56).Value
ORS("COMBD") = TS.Cells(intRow,57).Value
ORS("TTSS") = TS.Cells(intRow,58).Value
ORS("TTTD") = TS.Cells(intRow,59).Value
ORS("TTBD") = TS.Cells(intRow,60).Value
ORS("PMJJBYSS") = TS.Cells(intRow,61).Value
ORS("PMJJBYTD") = TS.Cells(intRow,62).Value
ORS("PMJJBYBD") = TS.Cells(intRow,63).Value
ORS("PMSBYSS") = TS.Cells(intRow,64).Value
ORS("PMSBYTD") = TS.Cells(intRow,65).Value
ORS("PMSBYBD") = TS.Cells(intRow,66).Value
ORS("Break") = TS.Cells(intRow,67).Value
ORS("MYACCSS") = TS.Cells(intRow,68).Value
ORS("MYACCTD") = TS.Cells(intRow,69).Value
ORS("MYACCBD") = TS.Cells(intRow,70).Value
ORS("ACCDETSS") = TS.Cells(intRow,71).Value
ORS("ACCDETTD") = TS.Cells(intRow,72).Value
ORS("ACCDETBD") = TS.Cells(intRow,73).Value
ORS("MINISSS") = TS.Cells(intRow,74).Value
ORS("MINISTD") = TS.Cells(intRow,75).Value
ORS("MINISBD") = TS.Cells(intRow,76).Value
ORS("BINVEST") = TS.Cells(intRow,77).Value
ORS("CBREQSS") = TS.Cells(intRow,78).Value
ORS("CBREQTD") = TS.Cells(intRow,79).Value
ORS("CBREQBD") = TS.Cells(intRow,80).Value
ORS("CHSTATSS") = TS.Cells(intRow,81).Value
ORS("CHSTATTD") = TS.Cells(intRow,82).Value
ORS("CHSTATBD") = TS.Cells(intRow,83).Value
ORS("STOPCHQSS") = TS.Cells(intRow,84).Value
ORS("STOPCHQTD") = TS.Cells(intRow,85).Value
ORS("STOPCHQBD") = TS.Cells(intRow,86).Value
ORS("AADHARSS") = TS.Cells(intRow,87).Value
ORS("AADHARTD") = TS.Cells(intRow,88).Value
ORS("AADHARBD") = TS.Cells(intRow,89).Value
ORS("ACCSTATSS") = TS.Cells(intRow,90).Value
ORS("ACCSTATTD") = TS.Cells(intRow,91).Value
ORS("ACCSTATBD") = TS.Cells(intRow,92).Value
ORS("INTCERTSS") = TS.Cells(intRow,93).Value
ORS("INTCERTTD") = TS.Cells(intRow,94).Value
ORS("INTCERTBD") = TS.Cells(intRow,95).Value
ORS("DCHOTSS") = TS.Cells(intRow,96).Value
ORS("DCHOTTD") = TS.Cells(intRow,97).Value
ORS("DCHOTBD") = TS.Cells(intRow,98).Value
ORS("TDSCERTSS") = TS.Cells(intRow,99).Value
ORS("TDSCERTTD") = TS.Cells(intRow,100).Value
ORS("TDSCERTBD") = TS.Cells(intRow,101).Value
ORS("FORM15SS") = TS.Cells(intRow,102).Value
ORS("FORM15TD") = TS.Cells(intRow,103).Value
ORS("FORM15BD") = TS.Cells(intRow,104).Value
ORS("NOMISS") = TS.Cells(intRow,105).Value
ORS("NOMITD") = TS.Cells(intRow,106).Value
ORS("NOMIBD") = TS.Cells(intRow,107).Value
ORS("DCREQSS") = TS.Cells(intRow,108).Value
ORS("DCREQTD") = TS.Cells(intRow,109).Value
ORS("DCREQBD") = TS.Cells(intRow,110).Value
ORS("ACCTFRSS") = TS.Cells(intRow,111).Value
ORS("ACCTFRTD") = TS.Cells(intRow,112).Value
ORS("ACCTFRBD") = TS.Cells(intRow,113).Value
ORS("FDRDCLSS") = TS.Cells(intRow,114).Value
ORS("FDRDCLTD") = TS.Cells(intRow,115).Value
ORS("FDRDCLBD") = TS.Cells(intRow,116).Value
ORS("SSASS") = TS.Cells(intRow,117).Value
ORS("SSATD") = TS.Cells(intRow,118).Value
ORS("SSABD") = TS.Cells(intRow,119).Value
ORS("DGPSS") = TS.Cells(intRow,120).Value
ORS("DGPTD") = TS.Cells(intRow,121).Value
ORS("DGPBD") = TS.Cells(intRow,122).Value
ORS("IBREGSS") = TS.Cells(intRow,123).Value
ORS("IBREGTD") = TS.Cells(intRow,124).Value
ORS("IBREGBD") = TS.Cells(intRow,125).Value
ORS("IBPWDSS") = TS.Cells(intRow,126).Value
ORS("IBPWDTD") = TS.Cells(intRow,127).Value
ORS("IBPWDRD") = TS.Cells(intRow,128).Value
ORS("COMAILSS") = TS.Cells(intRow,129).Value
ORS("COMAILTD") = TS.Cells(intRow,130).Value
ORS("COMAILBD") = TS.Cells(intRow,131).Value
ORS.Update


End If

intRow = intRow +1
Loop

ORS.Close
DB.Close
Set ORS = Nothing
Set DB = Nothing
xlApp.DisplayAlerts = False
xlApp.Workbook.Close False
xlApp.Quit
		
		
Set xlApp = Nothing
Set xlBook = Nothing
Set xlSheet = Nothing
Set TS = Nothing
Set intRow = Nothing
Set DateC = Nothing
