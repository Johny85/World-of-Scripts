Dim strDSN, DateF
strDSN = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=C:\Users\PR172959\Documents\Pritimay\Database.accdb"

Const adVarWChar = 202
Const adSingle = 4
Const adLockOptimistic = 3
Const adOpenDynamic = 2
Const adCmdTable = &H0002

'On Error Resume Next
Dim objCatalog
Set objCatalog = CreateObject("ADOX.Catalog")
Set objTable = CreateObject("ADOX.Table")
objCatalog.ActiveConnection = strDSN

DateF = Date()-1
With objTable
    .Name = "Recharge_Status_"&DateF
    .Columns.Append "Serial", adVarWChar, 12
    .Columns.Append "fname", adVarWChar, 50
    .Columns.Append "lname", adVarWChar, 50
    .Columns.Append "dept", adVarWChar, 50
    .Columns.Append "phone", adVarWChar, 50
    .Columns.Append "email", adVarWChar, 50
    .Columns.Append "jobID", adSingle
    .Columns.Append "birthDay", adVarWChar, 20
End With

objCatalog.Tables.Append objTable

If Err.Number <> 0 Then
    wscript.echo "fail: " & err.Number & ": " & err.Description
    Set objTable = Nothing
    Set objCatalog = Nothing
Else
    wscript.echo "info: table created successfully"
End If

Set objTable = Nothing
Set objCatalog = Nothing
Set strDSN = Nothing