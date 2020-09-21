Option Explicit
On Error Resume Next

Dim DateC, DB, WSO, WS, RS, strvar, TS, i, j, FieldN, FieldV, filsys
Dim xlApp


Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3

'WScript.Sleep 240000
 
DateC = Date()-1

Set filsys = CreateObject("Scripting.FileSystemObject")
filsys.CopyFile "E:\Mconnect Plus\November.xlsx", "E:\Mconnect Plus\MIS\"&DateC&".xlsx"
Set filsys = Nothing
WScript.Echo ("SAMPLE FILE COPIED")

Set DB = CreateObject("ADODB.Connection")
		
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\Mconnect Plus\MConnect_Plus.accdb;"

WScript.Echo ("Connection Success")

Set RS = DB.Execute("SELECT SBCAODSS as SUCCESS, SBCAODTD as TECHNICALD, SBCAODBD as BUSINESSD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT PPFSS, PPFTD, PPFBD FROM DAILY_MIS where DATEF = '"&DateC&"'UNION ALL SELECT LOANSS, LOANTD, LOANBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT SSFTSS, SSFTTD, SSFTBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT TPFTSS, TPFTTD, TPFTBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT IMPSSS, IMPSTDREM,IMPSBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT NEFTSS, NEFTTD, NEFTBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT MRECHARGESS, MRECHARGETD, MRECHARGEBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT BBPSDRECHARGESS, BBPSDRECHARGETD, BBPSDRECHARGEBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT NBBPSDRECHARGESS, NBBPSDRECHARGETD, NBBPSDRECHARGEBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT REGBILLPAYSS, REGBILLPAYTD, REGBILLPAYBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT QBILLPAYSS, QBILLPAYTD, QBILLPAYBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT TBBPSPAYSS, TBBPSPAYTD, TBBPSPAYBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT CCPAYSS, CCPAYTD, CCPAYBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT FDOPENSS, FDOPENTD, FDOPENBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT RDOPENSS, RDOPENTD, RDOPENBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT STPAYSS, STPAYTD, STPAYBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT COMSETTLED, COMTD, COMBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT TTSS, TTTD, TTBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT PMJJBYSS, PMJJBYTD, PMJJBYBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT PMSBYSS, PMSBYTD, PMSBYBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT FASTAGSS, FASTAGTD, FASTAGBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT APYSS, APYTD, APYBD FROM DAILY_MIS where DATEF = '"&DateC&"'")

WScript.Echo ("Data Extracted")

Set xlApp = CreateObject("Excel.Application")
xlApp.Workbooks.Open("E:\Mconnect Plus\MIS\"&DateC&".xlsx")

Set TS = xlApp.ActiveWorkbook.Worksheets("Sample")

j = 13
DO WHILE i<23
TS.Cells(j,3).Value = ""&RS.Fields(0)&""
TS.Cells(j,4).Value = ""&RS.Fields(1)&""
TS.Cells(j,5).Value = ""&RS.Fields(2)&""
RS.MoveNext
i = i+1
j = j+1
Loop

RS.Close
Set RS = Nothing
Set i = Nothing
Set j = Nothing

WScript.Echo ("FINANCIAL TRANSACTION DATA UPDATED")


Set RS = DB.Execute("SELECT MYACCSS, MYACCTD, MYACCBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT ACCDETSS, ACCDETTD, ACCDETBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT MINISSS, MINISTD, MINISBD FROM DAILY_MIS where  DATEF = '"&DateC&"' UNION ALL SELECT BINVEST,'','' FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT CBREQSS, CBREQTD, CBREQBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT CHSTATSS, CHSTATTD, CHSTATBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT STOPCHQSS, STOPCHQTD, STOPCHQBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT AADHARSS, AADHARTD, AADHARBD FROM DAILY_MIS  where DATEF = '"&DateC&"' UNION ALL SELECT ACCSTATSS, ACCSTATTD, ACCSTATBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT INTCERTSS, INTCERTTD, INTCERTBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT DCHOTSS, DCHOTTD, DCHOTBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT TDSCERTSS, TDSCERTTD, TDSCERTBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT FORM15SS, FORM15TD, FORM15BD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT NOMISS, NOMITD, NOMIBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT DCREQSS, DCREQTD, DCREQBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT ACCTFRSS, ACCTFRTD, ACCTFRBD FROM DAILY_MIS where  DATEF = '"&DateC&"' UNION ALL SELECT FDRDCLSSS, FDRDCLSTD, FDRDCLSBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT SSASS, SSATD, SSABD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT DGPSS, DGPTD, DGPBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT IBREGSS, IBREGTD, IBREGBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT IBPWDSS, IBPWDTD, IBPWDBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT COMAILSS, COMAILTD, COMAILBD FROM DAILY_MIS where DATEF = '"&DateC&"' UNION ALL SELECT DCLIMITSS, DCLIMITTD, DCLIMITBD FROM DAILY_MIS where DATEF = '"&DateC&"'")


i = 1
j = 13
DO WHILE i<=23
TS.Cells(j,11).Value = ""&RS.Fields(0)&""
TS.Cells(j,12).Value = ""&RS.Fields(1)&""
TS.Cells(j,13).Value = ""&RS.Fields(2)&""
RS.MoveNext
i = i+1
j = j+1
Loop

WScript.Echo ("NON-FINANCIAL TRANSACTION DATA UPDATED")
RS.Close
Set RS = Nothing
Set i = Nothing
Set j = Nothing



Set RS = DB.Execute("SELECT IMPSTDBEN FROM DAILY_MIS where DATEF = '"&DateC&"'")

DO WHILE NOT RS.EOF
If TS.Cells(37,01).Value = "" Then
TS.Cells(37,01).Value = "Decline in IMPS at Beneficiary End ("&RS.Fields(0)&")"
End If
RS.MoveNext
Loop


RS.Close
Set RS = Nothing




TS.Cells(1,2).Value = "MCONNECT PLUS DAILY REGISTRATION CHANNEL-WISE for "&DateC&""
TS.Cells(5,2).Value = "MCONNECT PLUS DAILY TRANSACTION MIS for "&DateC&""

Set RS = DB.Execute("SELECT CHANNEL, COUNT(*) from USER_REG where REG_DATE = '"&DateC&"' group by CHANNEL")

DO WHILE NOT RS.EOF
FieldN = RS.Fields(0)
FieldV = RS.Fields(1)


If FieldN = "ATM" Then
TS.Cells(3,3).Value = ""&FieldV&""
ElseIf FieldN = "BRANCH" Then
TS.Cells(3,5).Value = ""&FieldV&""
ElseIf FieldN = "DEVICE" Then
TS.Cells(3,7).Value = ""&FieldV&""
ElseIf FieldN = "INTERNET BANKING" Then
TS.Cells(3,10).Value = ""&FieldV&""
ElseIf FieldN = "KIOSK" Then
TS.Cells(3,9).Value = ""&FieldV&""
Else TS.Cells(3,11).Value = ""&FieldV&""
End If
RS.MoveNext
Loop

WScript.Echo ("USER REGISTRATION DATA INSERTED TO SAMPLE FILE")

RS.CLose
DB.Close






xlApp.ActiveWorkbook.Save
xlApp.ActiveWorkbook.Close
xlApp.Application.Quit
		
		
Set xlApp = Nothing
Set xlBook = Nothing
Set xlSheet = Nothing
Set TS = Nothing
Set FieldN = Nothing
Set FieldV = Nothing

Set RS = Nothing
Set DB = Nothing
