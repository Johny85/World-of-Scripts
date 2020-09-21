Option Explicit
On Error Resume Next

Dim  Counter, objFSO, objTextFile, FileCounter, strNextLine, RecordSize

Set RecordSize = 1000000 
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile ("C:\Users\PR172959\Documents\New Account Opened\SMS\reg.csv", ForReading)
Set Counter = 0
Set FileCounter = 0
Set objOutTextFile = Nothing

Do Until objTextFile.AtEndOfStream
    if Counter = 0 Or Counter = RecordSize Then
        Counter = 0
        FileCounter = FileCounter + 1
    if Not objOutTextFile is Nothing then objOutTextFile.Close        
    Set objOutTextFile = objFSO.OpenTextFile("C:\output_" & FileCounter & ".txt", ForWriting, True)
    end if
    strNextLine = objTextFile.Readline
    objOutTextFile.WriteLine(strNextLine)
    Counter = Counter + 1
Loop
objTextFile.Close
objOutTextFile.Close
Msgbox "Done..."

Set Counter = Nothing
Set objFSO = Nothing
Set objTextFile = Nothing
Set FileCounter = Nothing
Set strNextLine = Nothing