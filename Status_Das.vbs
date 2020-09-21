Option Explicit
On Error Resume Next

Dim oIE

Set oIE = WScript.CreateObject("InternetExplorer.Application")
oIE.left=10 ' window position
oIE.top = 10 ' and other properties
oIE.height = 380
oIE.width = 710
oIE.menubar = 0 ' no menu
oIE.toolbar = 0
oIE.statusbar = 0
oIE.body = 0
oIE.resizable = 0 ' disable resizing
oIE.navigate "E:\Recon\Form1.html"
oIE.visible = 1 ' keep visible

Set oIE = Nothing
