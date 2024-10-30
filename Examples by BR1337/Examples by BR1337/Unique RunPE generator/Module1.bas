Attribute VB_Name = "Module1"

Public Function RandomNumber() As Integer
Randomize
var1 = Int(9 * Rnd)
RandomNumber = var1
End Function

Public Function strings()
For i = 1 To 23
If i = 2 Or i = 4 Or i = 6 Then
strings = strings & RandomNumber
Else
strings = strings & RandomLetter
End If
Next i

End Function


Public Function RandomLetter() As String
Anfang:
Keyset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuyvwz"
Randomize
var1 = Int(Len(Keyset) * Rnd)
If var1 = 0 Then GoTo Anfang
RandomLetter = Mid(Keyset, var1, 1)
End Function
