Attribute VB_Name = "Module1"
Public Function Rc5(sText As String, sKey As String) As String
Dim i, x, y As Integer, b() As Byte, k() As Byte

Rc5 = vbNullString
x = 0
b() = StrConv(sText, vbFromUnicode)
k() = StrConv(sKey, vbFromUnicode)
For i = 0 To Len(sText) - 1
    If x = Len(sKey) - 1 Then
        x = 0
    Else
        x = x + 1
    End If
   
    For y = 1 To 255
        b(i) = b(i) Xor k(x) Mod (y + 5)
    Next y
Next i
Rc5 = StrConv(b, vbUnicode)
End Function
