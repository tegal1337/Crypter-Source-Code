Attribute VB_Name = "mRandom"
Public Function lRan(chs As String)
  Dim num_characters As Integer
  Dim i As Integer
  Dim txt As String
  Dim ch As Integer
  Randomize
  num_characters = CInt(chs)
  For i = 1 To num_characters
  ch = Int((26 + 26 + 10) * Rnd)
  If ch < 26 Then
  txt = txt & Chr$(ch + Asc("A"))
  ElseIf ch < 2 * 26 Then
  ch = ch - 26
  txt = txt & Chr$(ch + Asc("a"))
  Else
  ch = ch - 26 - 26
  End If
  Next i
  lRan = txt
End Function

Function StringtoChar(Text As String)
Dim tmp As String, Data As String
tmp = Len(Text)
For i = 1 To tmp
Data = Data & "chr(" & Asc(Mid(Text, i, 1)) & ")" & " &  "
Next i
StringtoChar = Left(Data, Len(Data) - 3)
End Function

Public Function RandomLetter() As String
  RandomLetter = ""
  Dim Keyset As String
  Keyset = "ABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGHIJKLMNOPQRSTUVWXYZ"
Anfang:
  Randomize
  var1 = Int(26 * Rnd)
  If var1 = 0 Then GoTo Anfang
  RandomLetter = Mid(Keyset, var1, 1)
End Function

Public Function RandomNumber() As String
  RandomNumber = ""
als:
  Randomize
  var1 = Int(9 * Rnd)
  RandomNumber = var1
If RandomNumber = "0" Then GoTo als
End Function

Public Function ByteArrayToString(bytArray() As Byte) As String
    Dim sAns As String
    Dim iPos As String
    
    sAns = StrConv(bytArray, vbUnicode)
    iPos = InStr(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)
    
    ByteArrayToString = sAns
 
 End Function


Public Function Xorex(ByVal sStr As String, ByVal sKey As String) As String
Dim i As Long
    For i = 1 To Len(sStr)
        Xorex = Xorex & Chr(Asc(Mid(sKey, IIf(i Mod Len(sKey) <> 0, i Mod Len(sKey), Len(sKey)), 1)) Xor Asc(Mid(sStr, i, 1)))
    Next i
End Function


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

