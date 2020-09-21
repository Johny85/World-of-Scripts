Option Explicit
On Error Resume Next

Const DeleteReadOnly = True

Dim DateC, DB, RS, xlApp, TS, j, strc

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 

DateC = Date()-1

Set DB = CreateObject("ADODB.Connection")
		
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\Mconnect Plus\MConnect_Plus.accdb;"

WScript.Echo ("Connection Success")

Set RS = DB.Execute("SELECT SBCAODSS, PPFSS, LOANSS, SSFTSS, TPFTSS, IMPSSS, NEFTSS, MRECHARGESS, BBPSDRECHARGESS, NBBPSDRECHARGESS, REGBILLPAYSS, QBILLPAYSS, TBBPSPAYSS, CCPAYSS, FDOPENSS, RDOPENSS, STPAYSS, COMSETTLED, TTSS, PMJJBYSS, PMSBYSS, FASTAGSS, APYSS, MYACCSS, ACCDETSS, MINISSS, BINVEST, CBREQSS, CHSTATSS, STOPCHQSS, AADHARSS, ACCSTATSS, INTCERTSS, DCHOTSS, TDSCERTSS, FORM15SS, NOMISS, DCREQSS, ACCTFRSS, FDRDCLSSS, SSASS, DGPSS, IBREGSS, IBPWDSS, COMAILSS, DCLIMITSS FROM DAILY_MIS where DATEF = '"&DateC&"'")

WScript.Echo ("Data Extracted")

Set xlApp = CreateObject("Excel.Application")
xlApp.Workbooks.Open("E:\Mconnect Plus\MISM\ChannelWise_MIS.xlsx")

Set TS = xlApp.ActiveWorkbook.Worksheets("APR-2020")

For j = 4 to 80
	strc = StrComp(DateC,TS.Cells(j,1).Value)
	If strc = 0 Then
	TS.Cells(j,2).Value = ""&RS.Fields(0)&""
	TS.Cells(j,3).Value = ""&RS.Fields(1)&""
	TS.Cells(j,4).Value = ""&RS.Fields(2)&""
	TS.Cells(j,5).Value = ""&RS.Fields(3)&""
	TS.Cells(j,6).Value = ""&RS.Fields(4)&""
	TS.Cells(j,7).Value = ""&RS.Fields(5)&""
	TS.Cells(j,8).Value = ""&RS.Fields(6)&""
	TS.Cells(j,9).Value = ""&RS.Fields(7)&""
	TS.Cells(j,10).Value = ""&RS.Fields(8)&""
	TS.Cells(j,11).Value = ""&RS.Fields(9)&""
	TS.Cells(j,12).Value = ""&RS.Fields(10)&""
	TS.Cells(j,13).Value = ""&RS.Fields(11)&""
	TS.Cells(j,14).Value = ""&RS.Fields(12)&""
	TS.Cells(j,15).Value = ""&RS.Fields(13)&""
	TS.Cells(j,16).Value = ""&RS.Fields(14)&""
	TS.Cells(j,17).Value = ""&RS.Fields(15)&""
	TS.Cells(j,18).Value = ""&RS.Fields(16)&""
	TS.Cells(j,19).Value = ""&RS.Fields(17)&""
	TS.Cells(j,20).Value = ""&RS.Fields(18)&""
	TS.Cells(j,21).Value = ""&RS.Fields(19)&""
	TS.Cells(j,22).Value = ""&RS.Fields(20)&""
	TS.Cells(j,23).Value = ""&RS.Fields(21)&""
	TS.Cells(j,24).Value = ""&RS.Fields(22)&""
	TS.Cells(j,25).Value = ""&RS.Fields(23)&""
	TS.Cells(j,26).Value = ""&RS.Fields(24)&""
	TS.Cells(j,27).Value = ""&RS.Fields(25)&""
	TS.Cells(j,28).Value = ""&RS.Fields(26)&""
	TS.Cells(j,29).Value = ""&RS.Fields(27)&""
	TS.Cells(j,30).Value = ""&RS.Fields(28)&""
	TS.Cells(j,31).Value = ""&RS.Fields(29)&""
	TS.Cells(j,32).Value = ""&RS.Fields(30)&""
	TS.Cells(j,33).Value = ""&RS.Fields(31)&""
	TS.Cells(j,34).Value = ""&RS.Fields(32)&""
	TS.Cells(j,35).Value = ""&RS.Fields(33)&""
	TS.Cells(j,36).Value = ""&RS.Fields(34)&""
	TS.Cells(j,37).Value = ""&RS.Fields(35)&""
	TS.Cells(j,38).Value = ""&RS.Fields(36)&""
	TS.Cells(j,39).Value = ""&RS.Fields(37)&""
	TS.Cells(j,40).Value = ""&RS.Fields(38)&""
	TS.Cells(j,41).Value = ""&RS.Fields(39)&""
	TS.Cells(j,42).Value = ""&RS.Fields(40)&""
	TS.Cells(j,43).Value = ""&RS.Fields(41)&""
	TS.Cells(j,44).Value = ""&RS.Fields(42)&""
	TS.Cells(j,45).Value = ""&RS.Fields(43)&""
	TS.Cells(j,46).Value = ""&RS.Fields(44)&""
	TS.Cells(j,47).Value = ""&RS.Fields(45)&""
	Else
		j = j+1
	End If
Next

RS.Close
DB.Close
Set RS = Nothing
Set j = Nothing

WScript.Echo ("TRANSACTION DATA UPDATED")


xlApp.ActiveWorkbook.Save
xlApp.ActiveWorkbook.Close
xlApp.Application.Quit
		
		
Set xlApp = Nothing
Set TS = Nothing


Set RS = Nothing
Set DB = Nothing