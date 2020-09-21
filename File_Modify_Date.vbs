Option Explicit
On Error Resume Next

Dim objFS, objFile, DateC

DateC = Date()-1

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.GetFile("E:\Mconnect Plus\MIS\"&DateC&".xlsx")

WScript.Echo left(objFile.DateLastModified,10)

Set DateC = Nothing
Set objFS = Nothing
Set objFile = Nothing
