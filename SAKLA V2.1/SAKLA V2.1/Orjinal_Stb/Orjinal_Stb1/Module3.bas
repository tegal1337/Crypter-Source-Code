Attribute VB_Name = "Module3"
Option Explicit
Public i As Integer
Public j As Integer
Public k As Integer
Public a As Byte
Public B As Byte
Dim M As Integer
Public L As Long
Public RC4KEY(255) As Byte
Public ADDTABLE(255, 255) As Byte
Dim STATE(0 To 255) As Byte

Public Sub FILL_LINEAR()
    Dim bCONST(0 To 255) As Byte
    For M = 0 To 255
        bCONST(M) = M
        STATE(M) = bCONST(M)
    Next M
End Sub

Public Sub RC4(BYTEARRAY() As Byte, Optional PASSWORD As String)
  If PASSWORD <> "" Then PREPARE_KEY PASSWORD
  For L = 0 To UBound(BYTEARRAY)
    i = ADDTABLE(i, 1)
    j = ADDTABLE(j, STATE(i))
    a = STATE(i): STATE(i) = STATE(j): STATE(j) = a
    B = STATE(ADDTABLE(STATE(i), STATE(j)))
    BYTEARRAY(L) = BYTEARRAY(L) Xor B
  Next L
End Sub

Public Sub PREPARE_KEY(sKEY As String)
  INITIALIZE_ADDTABLE
  FILL_LINEAR
  k = Len(sKEY)
  For i = 0 To k - 1
    B = Asc(Mid$(sKEY, i + 1, 1))
    For j = i To 255 Step k
      RC4KEY(j) = B
    Next j
  Next i
  j = 0
  For i = 0 To 255
    k = ADDTABLE(STATE(i), RC4KEY(i))
    j = ADDTABLE(j, k)
    B = STATE(i): STATE(i) = STATE(j): STATE(j) = B
  Next i
  i = 0
  j = 0
End Sub

Public Sub INITIALIZE_ADDTABLE()
  Static BeenHereDoneThat As Boolean
  If BeenHereDoneThat Then Exit Sub
  For j = 0 To 255
    For i = 0 To 255
      ADDTABLE(i, j) = CByte((i + j) And 255)
    Next i
  Next j
  BeenHereDoneThat = True
End Sub
Public Function STRING_TO_BYTES(sString As String) As Byte()
  STRING_TO_BYTES = StrConv(sString, vbFromUnicode)
End Function
Public Function BYTES_TO_STRING(bBytes() As Byte) As String
  BYTES_TO_STRING = bBytes
  BYTES_TO_STRING = StrConv(BYTES_TO_STRING, vbUnicode)
End Function
Public Function RC4_String(InputStr As String, PasswordStr As String) As String
Dim tmpByte() As Byte
tmpByte = STRING_TO_BYTES(InputStr)
RC4 tmpByte, PasswordStr
RC4_String = BYTES_TO_STRING(tmpByte)
End Function
Public Function SimpleHexDecrypt(lpszString As String, lpPassword As String) As String
Dim b1() As Byte, i1 As Long, s1 As String
ReDim Preserve b1(0 To (Len(lpszString) / 2) - 1) As Byte
For i1 = 1 To Len(lpszString) / 2
    b1(i1 - 1) = CLng("&H" & Mid(lpszString, 1 + (i1 - 1) * 2, 2))
Next
RC4 b1, lpPassword
SimpleHexDecrypt = BYTES_TO_STRING(b1)
End Function
