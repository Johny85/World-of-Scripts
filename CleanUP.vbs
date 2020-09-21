Option Explicit
On Error Resume Next

Dim DB, CMD

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const adOpenStatic = 3
Const adLockOptimistic = 3 


Set DB = CreateObject("ADODB.Connection")
Set CMD = CreateObject("ADODB.Command")
		
DB.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\Mconnect Plus\MConnect_Plus.accdb;"
	
With CMD
.ActiveConnection = DB
.CommandText = "Delete from DAILY_MIS"
End With
CMD.Execute
Set CMD = Nothing


Set CMD = CreateObject("ADODB.Command")
		
With CMD
.ActiveConnection = DB
.CommandText = "Delete from USER_REG"
End With
CMD.Execute
Set CMD = Nothing

DB.Close
Set DB = Nothing