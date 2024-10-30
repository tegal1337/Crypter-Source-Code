Attribute VB_Name = "MakeCrypt"
Public Function MakeXor()

Dim a As String
Dim XorVal As String
XorVal = GenNumKey(RandomNumber(15, 8), 4)

a = a + "Public Function " + XorName + " (ByVal sData As String, Optional xKey as string) As String" + vbCrLf
a = a + "Dim " + XorVal + "      " + " as long " + vbCrLf
a = a + "for " + XorVal + " = 1 to len(sdata) " + vbCrLf
a = a + "   " + XorName + " = " + XorName + " & Chr$(Asc(Mid$(sData, " + XorVal + ", 1)) Xor xKey)" + vbCrLf
a = a + "Next " + XorVal + vbCrLf
a = a + "end function " + vbCrLf + vbCrLf

Randomize

MakeXor = a

End Function

Public Function MakeRotX()
    Dim i     As Long
    Dim a     As String
    
        a = a + "Public Function " + RotName + "(ByVal sData As String) As String" & vbCrLf
        a = a + "    Dim i       As Long" & vbCrLf
        a = a + "" & vbCrLf
        a = a + "    For i = 1 To Len(sData)" & vbCrLf
        a = a + RotName + " = " + RotName + " & Chr$(Asc(Mid$(sData, i, 1)) + " + RotNumber + ")" & vbCrLf
        a = a + "    Next i" & vbCrLf
        a = a + "End Function" & vbCrLf
        
MakeRotX = a

   End Function

Public Function MakeStrHex()

Dim a As String
Dim StrHex1 As String, StrHex2 As String

StrHex1 = GenNumKey(RandomNumber(15, 8), 4)
StrHex2 = GenNumKey(RandomNumber(15, 8), 4)

a = a + "Public Function " + HexName + " (Byval StrData as string) " + vbCrLf
a = a + "Dim " + StrHex1 + " as long, " + StrHex2 + " as string " + vbCrLf
a = a + "On Local Error Resume Next" + vbCrLf
a = a + "For " + StrHex1 + " = 1 to Len(StrData) Step 2 " + vbCrLf
a = a + StrHex2 + " = " + StrHex2 + " & Chr$(Val(""&H"" & Mid$(StrData, " + StrHex1 + ", 2))) " + vbCrLf
a = a + "Next " + StrHex1 + vbCrLf
a = a + HexName + " = " + StrHex2 + vbCrLf
a = a + "end function" + vbCrLf + vbCrLf

MakeStrHex = a

End Function


