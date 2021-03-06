Option Explicit
On Error Resume Next

Const DeleteReadOnly = True

Dim FSO, xlApp, filsys, TS, RegC, DB, ORS, intRow, RSO, RS, SLine, arr, strc, objFSO
Dim DateF, DateC, ObjOutlook, SSession, Item1, Atchm, Inbox, OSub, OFrm, ODate, IntC

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

DateC = Date()-1

Set filsys = CreateObject("Scripting.FileSystemObject")
filsys.CopyFile "E:\Mconnect Plus\November.xlsx", "E:\Mconnect Plus\MIS\"&DateC&".xlsx"
Set filsys = Nothing
WScript.Echo ("SAMPLE FILE COPIED")


Set xlApp = CreateObject("Excel.Application")
xlApp.Workbooks.Open("C:\MConnect\Reports\User Registration.xlsx")
xlApp.Workbooks(1).SaveAs "E:\Mconnect Plus\USER_REG_"&DateC&".csv",6

WScript.Echo ("USER REGISTRATION FILE COPIED")


Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\Mconnect Plus\MConnect_Plus.accdb;"
Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "USER_REG", DB, adOpenStatic, adLockOptimistic

Set RSO = CreateObject("Scripting.FileSystemObject")
Set RS = RSO.OpenTextFile("E:\Mconnect Plus\USER_REG_"&DateC&".csv")
	
	If not RS.AtEndOfStream Then RS.Skipline
	Do Until RS.AtEndOfStream
	
	SLine = RS.ReadLine
	SLine = Replace(SLine,chr(34),"")
	arr = split(SLine,",")
		
	ORS.AddNew
	ORS("USER_ID") = arr(0)
	ORS("USER_NAME") = arr(1)
	ORS("STATUS") = arr(2)
	ORS("REG_DATE") = Left(arr(3),10)
	ORS("REG_TIME") = Right(arr(3),8)
	ORS("MOBILE_NUMBER") = arr(4)
	ORS("BRANCH") = arr(5)
	ORS("CHANNEL") = arr(6)
	ORS.Update
	
	Loop

RS.Close
ORS.Close
WScript.Echo ("USER REGISTRATION FILE UPLOADED TO DB")
xlApp.DisplayAlerts = False
xlApp.Workbook.Close False
xlApp.Quit

Set RS = Nothing
Set xlApp = Nothing
Set filsys = Nothing
Set ORS = Nothing


Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists("E:\Mconnect Plus\Dashboard\MIS.xls") OR objFSO.FileExists("E:\Mconnect Plus\Dashboard\MIS.xlsx") Then
objFSO.DeleteFile("E:\Mconnect Plus\Dashboard\*")
End If
Set objFSO = Nothing


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
If UCase(Right(Atchm.FileName,11)) = "_REPORT.XLS" Then
Atchm.SaveAsFile "E:\Mconnect Plus\Dashboard\MIS.xls"
ElseIf UCase(Right(Atchm.FileName,12)) = "_REPORT.XLSX" Then
Atchm.SaveAsFile "E:\Mconnect Plus\Dashboard\MIS.xlsx"
End If
Next
End If
End If
Next

WScript.Echo ("MIS FILE EXTRACTED FROM MAILBOX")

Set Atchm = Nothing
Set OSub = Nothing
Set OFrm = Nothing
Set ODate = Nothing
Set ObjOutlook = Nothing
Set SSession = Nothing
Set Item1 = Nothing
Set Inbox = Nothing
Set DateF = Nothing
Set DateC = Nothing


'INSERT TO DB


DateC = Date()-1

Select Case Mid(DateC,4,7)
Case "01-2020"
DateC = (Left(DateC,2))*1&"-Jan-20"
Case "02-2020"
DateC = (Left(DateC,2))*1&"-Feb-20"
Case "03-2020"
DateC = (Left(DateC,2))*1&"-Mar-20"
Case "04-2020"
DateC = (Left(DateC,2))*1&"-Apr-20"
Case "05-2020"
DateC = (Left(DateC,2))*1&"-May-20"
Case "06-2020"
DateC = (Left(DateC,2))*1&"-Jun-20"
Case "07-2020"
DateC = (Left(DateC,2))*1&"-Jul-20"
Case "08-2020"
DateC = (Left(DateC,2))*1&"-Aug-20"
Case "09-2020"
DateC = (Left(DateC,2))*1&"-Sep-20"
Case "10-2020"
DateC = (Left(DateC,2))*1&"-Oct-20"
Case "11-2020"
DateC = (Left(DateC,2))*1&"-Nov-20"
Case "12-2020"
DateC = (Left(DateC,2))*1&"-Dec-20"
End Select

WScript.Echo ("DATE CONVERSION")

Set DB = CreateObject("ADODB.Connection")
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\Mconnect Plus\MConnect_Plus.accdb;"
Set ORS = CreateObject("ADODB.Recordset")
ORS.Open "DAILY_MIS", DB, adOpenStatic, adLockOptimistic

WScript.Echo ("DB CONNECTION SUCCESS")

Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists("E:\Mconnect Plus\Dashboard\MIS.xlsx") Then
intRow = 1
Set xlApp = CreateObject("Excel.Application")
xlApp.Workbooks.Open("E:\Mconnect Plus\Dashboard\MIS.xlsx")
Set TS = xlApp.ActiveWorkbook.Worksheets("APR - 2020")

Do While intRow < TS.UsedRange.Rows.Count
strc = StrComp(DateC,TS.Cells(intRow, 1).Text)
If strc = 0 Then

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
ORS("FASTAGSS") = TS.Cells(intRow,67).Value
ORS("FASTAGTD") = TS.Cells(intRow,68).Value
ORS("FASTAGBD") = TS.Cells(intRow,69).Value
ORS("APYSS") = TS.Cells(intRow,70).Value
ORS("APYTD") = TS.Cells(intRow,71).Value
ORS("APYBD") = TS.Cells(intRow,72).Value
ORS("Break") = TS.Cells(intRow,73).Value
ORS("MYACCSS") = TS.Cells(intRow,74).Value
ORS("MYACCTD") = TS.Cells(intRow,75).Value
ORS("MYACCBD") = TS.Cells(intRow,76).Value
ORS("ACCDETSS") = TS.Cells(intRow,77).Value
ORS("ACCDETTD") = TS.Cells(intRow,78).Value
ORS("ACCDETBD") = TS.Cells(intRow,79).Value
ORS("MINISSS") = TS.Cells(intRow,80).Value
ORS("MINISTD") = TS.Cells(intRow,81).Value
ORS("MINISBD") = TS.Cells(intRow,82).Value
ORS("BINVEST") = TS.Cells(intRow,83).Value
ORS("CBREQSS") = TS.Cells(intRow,84).Value
ORS("CBREQTD") = TS.Cells(intRow,85).Value
ORS("CBREQBD") = TS.Cells(intRow,86).Value
ORS("CHSTATSS") = TS.Cells(intRow,87).Value
ORS("CHSTATTD") = TS.Cells(intRow,88).Value
ORS("CHSTATBD") = TS.Cells(intRow,89).Value
ORS("STOPCHQSS") = TS.Cells(intRow,90).Value
ORS("STOPCHQTD") = TS.Cells(intRow,91).Value
ORS("STOPCHQBD") = TS.Cells(intRow,92).Value
ORS("AADHARSS") = TS.Cells(intRow,93).Value
ORS("AADHARTD") = TS.Cells(intRow,94).Value
ORS("AADHARBD") = TS.Cells(intRow,95).Value
ORS("ACCSTATSS") = TS.Cells(intRow,96).Value
ORS("ACCSTATTD") = TS.Cells(intRow,97).Value
ORS("ACCSTATBD") = TS.Cells(intRow,98).Value
ORS("INTCERTSS") = TS.Cells(intRow,99).Value
ORS("INTCERTTD") = TS.Cells(intRow,100).Value
ORS("INTCERTBD") = TS.Cells(intRow,101).Value
ORS("DCHOTSS") = TS.Cells(intRow,102).Value
ORS("DCHOTTD") = TS.Cells(intRow,103).Value
ORS("DCHOTBD") = TS.Cells(intRow,104).Value
ORS("TDSCERTSS") = TS.Cells(intRow,105).Value
ORS("TDSCERTTD") = TS.Cells(intRow,106).Value
ORS("TDSCERTBD") = TS.Cells(intRow,107).Value
ORS("FORM15SS") = TS.Cells(intRow,108).Value
ORS("FORM15TD") = TS.Cells(intRow,109).Value
ORS("FORM15BD") = TS.Cells(intRow,110).Value
ORS("NOMISS") = TS.Cells(intRow,111).Value
ORS("NOMITD") = TS.Cells(intRow,112).Value
ORS("NOMIBD") = TS.Cells(intRow,113).Value
ORS("DCREQSS") = TS.Cells(intRow,114).Value
ORS("DCREQTD") = TS.Cells(intRow,115).Value
ORS("DCREQBD") = TS.Cells(intRow,116).Value
ORS("ACCTFRSS") = TS.Cells(intRow,117).Value
ORS("ACCTFRTD") = TS.Cells(intRow,118).Value
ORS("ACCTFRBD") = TS.Cells(intRow,119).Value
ORS("FDRDCLSSS") = TS.Cells(intRow,120).Value
ORS("FDRDCLSTD") = TS.Cells(intRow,121).Value
ORS("FDRDCLSBD") = TS.Cells(intRow,122).Value
ORS("SSASS") = TS.Cells(intRow,123).Value
ORS("SSATD") = TS.Cells(intRow,124).Value
ORS("SSABD") = TS.Cells(intRow,125).Value
ORS("DGPSS") = TS.Cells(intRow,126).Value
ORS("DGPTD") = TS.Cells(intRow,127).Value
ORS("DGPBD") = TS.Cells(intRow,128).Value
ORS("IBREGSS") = TS.Cells(intRow,129).Value
ORS("IBREGTD") = TS.Cells(intRow,130).Value
ORS("IBREGBD") = TS.Cells(intRow,131).Value
ORS("IBPWDSS") = TS.Cells(intRow,132).Value
ORS("IBPWDTD") = TS.Cells(intRow,133).Value
ORS("IBPWDRD") = TS.Cells(intRow,134).Value
ORS("COMAILSS") = TS.Cells(intRow,135).Value
ORS("COMAILTD") = TS.Cells(intRow,136).Value
ORS("COMAILBD") = TS.Cells(intRow,137).Value
ORS("DCLIMITSS") = TS.Cells(intRow,138).Value
ORS("DCLIMITTD") = TS.Cells(intRow,139).Value
ORS("DCLIMITBD") = TS.Cells(intRow,140).Value
ORS.Update

End If
intRow = intRow +1
Loop

ElseIf FSO.FileExists("E:\Mconnect Plus\Dashboard\MIS.xls") Then
intRow = 1
Set xlApp = CreateObject("Excel.Application")
xlApp.Workbooks.Open("E:\Mconnect Plus\Dashboard\MIS.xls")
Set TS = xlApp.ActiveWorkbook.Worksheets("APR - 2020")

Do While intRow < TS.UsedRange.Rows.Count
strc = StrComp(DateC,TS.Cells(intRow, 1).Text)
If strc = 0 Then

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
ORS("FASTAGSS") = TS.Cells(intRow,67).Value
ORS("FASTAGTD") = TS.Cells(intRow,68).Value
ORS("FASTAGBD") = TS.Cells(intRow,69).Value
ORS("APYSS") = TS.Cells(intRow,70).Value
ORS("APYTD") = TS.Cells(intRow,71).Value
ORS("APYBD") = TS.Cells(intRow,72).Value
ORS("Break") = TS.Cells(intRow,73).Value
ORS("MYACCSS") = TS.Cells(intRow,74).Value
ORS("MYACCTD") = TS.Cells(intRow,75).Value
ORS("MYACCBD") = TS.Cells(intRow,76).Value
ORS("ACCDETSS") = TS.Cells(intRow,77).Value
ORS("ACCDETTD") = TS.Cells(intRow,78).Value
ORS("ACCDETBD") = TS.Cells(intRow,79).Value
ORS("MINISSS") = TS.Cells(intRow,80).Value
ORS("MINISTD") = TS.Cells(intRow,81).Value
ORS("MINISBD") = TS.Cells(intRow,82).Value
ORS("BINVEST") = TS.Cells(intRow,83).Value
ORS("CBREQSS") = TS.Cells(intRow,84).Value
ORS("CBREQTD") = TS.Cells(intRow,85).Value
ORS("CBREQBD") = TS.Cells(intRow,86).Value
ORS("CHSTATSS") = TS.Cells(intRow,87).Value
ORS("CHSTATTD") = TS.Cells(intRow,88).Value
ORS("CHSTATBD") = TS.Cells(intRow,89).Value
ORS("STOPCHQSS") = TS.Cells(intRow,90).Value
ORS("STOPCHQTD") = TS.Cells(intRow,91).Value
ORS("STOPCHQBD") = TS.Cells(intRow,92).Value
ORS("AADHARSS") = TS.Cells(intRow,93).Value
ORS("AADHARTD") = TS.Cells(intRow,94).Value
ORS("AADHARBD") = TS.Cells(intRow,95).Value
ORS("ACCSTATSS") = TS.Cells(intRow,96).Value
ORS("ACCSTATTD") = TS.Cells(intRow,97).Value
ORS("ACCSTATBD") = TS.Cells(intRow,98).Value
ORS("INTCERTSS") = TS.Cells(intRow,99).Value
ORS("INTCERTTD") = TS.Cells(intRow,100).Value
ORS("INTCERTBD") = TS.Cells(intRow,101).Value
ORS("DCHOTSS") = TS.Cells(intRow,102).Value
ORS("DCHOTTD") = TS.Cells(intRow,103).Value
ORS("DCHOTBD") = TS.Cells(intRow,104).Value
ORS("TDSCERTSS") = TS.Cells(intRow,105).Value
ORS("TDSCERTTD") = TS.Cells(intRow,106).Value
ORS("TDSCERTBD") = TS.Cells(intRow,107).Value
ORS("FORM15SS") = TS.Cells(intRow,108).Value
ORS("FORM15TD") = TS.Cells(intRow,109).Value
ORS("FORM15BD") = TS.Cells(intRow,110).Value
ORS("NOMISS") = TS.Cells(intRow,111).Value
ORS("NOMITD") = TS.Cells(intRow,112).Value
ORS("NOMIBD") = TS.Cells(intRow,113).Value
ORS("DCREQSS") = TS.Cells(intRow,114).Value
ORS("DCREQTD") = TS.Cells(intRow,115).Value
ORS("DCREQBD") = TS.Cells(intRow,116).Value
ORS("ACCTFRSS") = TS.Cells(intRow,117).Value
ORS("ACCTFRTD") = TS.Cells(intRow,118).Value
ORS("ACCTFRBD") = TS.Cells(intRow,119).Value
ORS("FDRDCLSSS") = TS.Cells(intRow,120).Value
ORS("FDRDCLSTD") = TS.Cells(intRow,121).Value
ORS("FDRDCLSBD") = TS.Cells(intRow,122).Value
ORS("SSASS") = TS.Cells(intRow,123).Value
ORS("SSATD") = TS.Cells(intRow,124).Value
ORS("SSABD") = TS.Cells(intRow,125).Value
ORS("DGPSS") = TS.Cells(intRow,126).Value
ORS("DGPTD") = TS.Cells(intRow,127).Value
ORS("DGPBD") = TS.Cells(intRow,128).Value
ORS("IBREGSS") = TS.Cells(intRow,129).Value
ORS("IBREGTD") = TS.Cells(intRow,130).Value
ORS("IBREGBD") = TS.Cells(intRow,131).Value
ORS("IBPWDSS") = TS.Cells(intRow,132).Value
ORS("IBPWDTD") = TS.Cells(intRow,133).Value
ORS("IBPWDRD") = TS.Cells(intRow,134).Value
ORS("COMAILSS") = TS.Cells(intRow,135).Value
ORS("COMAILTD") = TS.Cells(intRow,136).Value
ORS("COMAILBD") = TS.Cells(intRow,137).Value
ORS("DCLIMITSS") = TS.Cells(intRow,138).Value
ORS("DCLIMITTD") = TS.Cells(intRow,139).Value
ORS("DCLIMITBD") = TS.Cells(intRow,140).Value
ORS.Update

End If

intRow = intRow +1
Loop
End If

ORS.Close
DB.Close

WScript.Echo ("MIS FILE DATA UPLOADED TO DB")

Set ORS = Nothing
Set DB = Nothing
Set FSO = Nothing
xlApp.DisplayAlerts = False
xlApp.Workbook.Close False
xlApp.Quit
		
		
Set xlApp = Nothing
Set xlBook = Nothing
Set xlSheet = Nothing
Set TS = Nothing
Set intRow = Nothing
Set DateC = Nothing
Set strc = Nothing