Attribute VB_Name = "mFunctions"
Option Explicit
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Const N1 As String = "abcdefghijklmnopqrstuvwxyz"
Private Const N2 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const N3 As String = "1234567890"
Private Const N4 As String = "ß´+#-.,;:_'*?`=)(/&%$§!°^<>|"
Public Function RndNames() As String
Dim nb As String
Dim i As Integer
Randomize
nb = N1 + N2 + N3 + N4
For i = 1 To 28
RndNames = RndNames & Mid$(nb, Int((Rnd * Len(nb)) + 1), 1)
Next i
End Function
Public Function FileExists(ByVal sFilename As String) As Boolean
Dim Fl As Integer: Fl = Len(Dir$(sFilename))
    On Local Error Resume Next
    FileExists = IIf(Err Or Fl = 0, False, True)
End Function
Public Function AddJpgHeader(bArray() As Byte) As Byte()
    Dim bJpg(9) As Byte
    Dim bRes() As Byte
    
    bJpg(0) = 255
    bJpg(1) = 216
    bJpg(2) = 255
    bJpg(3) = 224
    bJpg(4) = 32
    bJpg(5) = 16
    bJpg(6) = 74
    bJpg(7) = 70
    bJpg(8) = 73
    bJpg(9) = 70
    
    bRes() = bJpg()
    Dim lPos As Long
                
    lPos = UBound(bRes)
    ReDim Preserve bRes(UBound(bRes) + UBound(bArray) + 1)
    RtlMoveMemory bRes(lPos + 1), bArray(0), UBound(bArray) + 1
    
    AddJpgHeader = bRes()
End Function
