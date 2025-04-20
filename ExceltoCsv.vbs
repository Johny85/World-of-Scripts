Option Explicit
On Error Resume Next

Const ForReading = 1
Const ForWriting = 2

Dim strPath, oFSO, oFile, oExcel, oBook, wcnt, i


Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = False
oExcel.DisplayAlerts = False

strPath = SelectFolder("")
If strPath = vbNull Then
WScript.Echo "Error: No Folder Selected"
Else
WScript.Echo strPath
End If


For Each oFile In oFSO.GetFolder(strPath).Files
If UCase(oFSO.GetExtensionName(oFile.Name)) = "XLSX" OR UCase(oFSO.GetExtensionName(oFile.Name)) = "XLS" Then

'Wscript.Echo strPath&"\"&(oFile.Name)
Set oBook = oExcel.Workbooks.Open(strPath&"\"&(oFile.Name))
'Wscript.Echo (oBook.Name)
'oBook.SaveAs strPath & "\" & oBook.Name & ".csv", 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, true

wcnt = oBook.Worksheets.Count
'Wscript.Echo wcnt
For i = 1 to wcnt
'Wscript.Echo oBook.Worksheets(i).Name
oBook.Worksheets(i).SaveAs strPath & "\" & oBook.Name & "_" & oBook.Worksheets(i).Name & ".csv", 6
Next

Else
Wscript.Echo ("Error: No Excel File found on the selected location")
End if
oBook.Close False
Set oBook = Nothing
Next
oExcel.Quit

Function SelectFolder(myStartFolder)
Dim objFolder, objItem, objShell

On Error Resume Next

SelectFolder = vbNull

Set objShell  = CreateObject( "Shell.Application" )
Set objFolder = objShell.BrowseForFolder( 0, "Select Folder", 0, myStartFolder )

If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path

Set objFolder = Nothing
Set objshell  = Nothing
Set objItem = Nothing

End Function


Set strPath = Nothing
Set oFSO = Nothing
Set oFile = Nothing
Set oExcel = Nothing
Set oBook = Nothing
Set wcnt = Nothing
