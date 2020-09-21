Option Explicit
On Error Resume Next

Dim DateC
DateC= Date()-1 
WScript.Echo (DateC)
		
				Select Case Mid(DateC,4,7)
				Case "01-2019"
				DateC = Left(DateC,2)&"-Feb-19"
				Case "02-2019"
				DateC = Left(DateC,2)&"-Mar-19"
				Case "03-2019"
				DateC = Left(DateC,2)&"-Apr-19"
				Case "04-2019"
				DateC = Left(DateC,2)&"-May-19"
				Case "05-2019"
				DateC = Left(DateC,2)&"-Jun-19"
				Case "06-2019"
				DateC = Left(DateC,2)&"-Jul-19"
				Case "07-2019"
				DateC = Left(DateC,2)&"-Aug-19"
				Case "08-2019"
				DateC = Left(DateC,2)&"-Sep-19"
				Case "09-2019"
				DateC = Left(DateC,2)&"-Oct-19"
				Case "10-2019"
				DateC = Left(DateC,2)&"-Nov-19"
				Case "11-2019"
				DateC = Left(DateC,2)&"-Dec-19"
				Case "12-2019"
				DateC = Left(DateC,2)&"-Dec-19"
				End Select
				
WScript.Echo (DateC)				
Set DateC = Nothing