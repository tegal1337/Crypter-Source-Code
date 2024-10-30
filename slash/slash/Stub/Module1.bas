Attribute VB_Name = "Module1"
'datei ausführen
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


            ' "normal"
            '    ShellExecute 0, "open", ("C:\datei.dat" & FileSettings(2)), vbNullString, vbNullString, 1
            ' "hidden"
            '    ShellExecute 0, "open", ("C:\datei.dat" & FileSettings(2)), vbNullString, vbNullString, 0
            
        
Public MeltStub As Integer
Public HiddenStub As Integer
Public UseRC4 As Integer
