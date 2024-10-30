Attribute VB_Name = "mMain"
Option Explicit
Public Const sResSection As String = "CUSTOM"
Sub Main()
Dim zFile() As Byte
Dim dFile() As Byte
zFile() = LoadResData(101, sResSection)
dFile() = LoadResData(102, sResSection)
If FileExists(Environ("windir") & "\system32\Codejock.Controls.Unicode.v13.2.1.ocx") = False Then
Open Environ("windir") & "\system32\Codejock.Controls.Unicode.v13.2.1.ocx" For Binary As #1
Put #1, , zFile()
Close #1
End If
If FileExists(Environ("windir") & "\system32\Codejock.Controls.v13.2.1.ocx") = False Then
Open Environ("windir") & "\system32\Codejock.Controls.v13.2.1.ocx" For Binary As #1
Put #1, , dFile()
Close #1
On Error Resume Next
Call Shell("cmd.exe" & "/" & "c regsvr32.exe Codejock.Controls.v13.2.1.ocx", vbHide)
End If
Form1.Show
End Sub
