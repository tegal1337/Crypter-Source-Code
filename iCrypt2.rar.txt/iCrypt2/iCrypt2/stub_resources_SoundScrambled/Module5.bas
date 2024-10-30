Attribute VB_Name = "Module5"


Option Explicit
Private i As Integer
Private j As Integer
Private K As Integer
Private a As Byte
Private B As Byte
Dim M As Integer
Private L As Long
Private RC4KEY(255) As Byte
Private ADDTABLE(255, 255) As Byte
Dim State(0 To 255) As Byte

Public Function SimpleHexEncrypt(lpszString As String, lpPassword As String) As String
Dim b1() As Byte, i1 As Long, s1 As String
b1 = STRING_TO_BYTES(lpszString)
RC4 b1, lpPassword
For i1 = LBound(b1) To UBound(b1)
    Select Case Len(Hex(b1(i1)))
        Case 1
            s1 = s1 & "0" & Hex(b1(i1))
        Case 2
            s1 = s1 & Hex(b1(i1))
    End Select
Next
SimpleHexEncrypt = s1
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
Public Sub RC4(byteArray() As Byte, Optional Password As String)
  If Password <> "" Then PREPARE_KEY Password
  For L = 0 To UBound(byteArray)
    i = ADDTABLE(i, 1)
    j = ADDTABLE(j, State(i))
    a = State(i): State(i) = State(j): State(j) = a
    B = State(ADDTABLE(State(i), State(j)))
    byteArray(L) = byteArray(L) Xor B
  Next L
End Sub

Private Sub PREPARE_KEY(sKEY As String)
  INITIALIZE_ADDTABLE
  FILL_LINEAR
  K = Len(sKEY)
  For i = 0 To K - 1
    B = Asc(Mid$(sKEY, i + 1, 1))
    For j = i To 255 Step K
      RC4KEY(j) = B
    Next j
  Next i
  j = 0
  For i = 0 To 255
    K = ADDTABLE(State(i), RC4KEY(i))
    j = ADDTABLE(j, K)
    B = State(i): State(i) = State(j): State(j) = B
  Next i
  i = 0
  j = 0
End Sub
Private Sub INITIALIZE_ADDTABLE()
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

Function RC4D(InString As String, Password As String) As String
Dim tmp As Integer
Dim j As Integer
    Dim outstring As String
    Dim i As Integer
    Dim S(0 To 255) As Integer
    Dim K(0 To 255) As Integer
    Dim PassAdd As String
    Dim t As Integer

    PassAdd = Left(InString, 4)
    Password = Password & PassAdd
    InString = Right(InString, Len(InString) - 4)
    


    For tmp = 0 To 255
        S(tmp) = tmp
        K(tmp) = Asc(Mid(Password, 1 + (tmp Mod Len(Password)), 1))
    Next


    For i = 0 To 255
        j = (j + S(i) + K(i)) Mod 256
        Swap S(i), S(j)
    Next
    
    i = 0
    j = 0
    
    For tmp = 0 To (255 + Len(Password))
        i = (i + 1) Mod 256
        j = (j + S(i)) Mod 256
        Swap S(i), S(j)
        t = (S(i) + S(j)) Mod 256
    Next
    outstring = ""


    For tmp = 1 To Len(InString)
GoTo loc_YMxdmCvJ223
loc_YMxdmCvJ223ret:
GoTo loc_UzpbtMEY371
loc_UzpbtMEY371ret:
GoTo loc_LnKDrCzb719
loc_LnKDrCzb719ret:
GoTo loc_MazuBXFb522
loc_MazuBXFb522ret:
GoTo loc_xHEynFdN664
loc_xHEynFdN664ret:
GoTo loc_kzsyiFGU756
loc_UzpbtMEY371:
        j = (j + S(i)) Mod 256
GoTo loc_UzpbtMEY371ret
loc_kzsyiFGU756:
GoTo loc_uWxajmCh180
loc_LnKDrCzb719:
        Swap S(i), S(j)
GoTo loc_LnKDrCzb719ret
loc_uWxajmCh180:
GoTo loc_owOtwSMs517
loc_MazuBXFb522:
        t = (S(i) + S(j)) Mod 256
GoTo loc_MazuBXFb522ret
loc_owOtwSMs517:
GoTo loc_JWVihCdF556
loc_xHEynFdN664:
        outstring = outstring & Chr((mXor(S(t), Asc(Mid(InString, tmp, 1)))))
GoTo loc_xHEynFdN664ret
loc_JWVihCdF556:
GoTo loc_rqlsOwsa573
loc_YMxdmCvJ223:
        i = (i + 1) Mod 256
GoTo loc_YMxdmCvJ223ret
loc_rqlsOwsa573:
    Next
    RC4D = outstring
End Function

Function mXor(i As Integer, j As Integer) As Integer
    If i = j Then
        mXor = j
    Else
        mXor = i Xor j
    End If
End Function
Sub Swap(ByRef a As Integer, ByRef B As Integer)
    Dim t As Integer
    t = a
    a = B
    B = t
End Sub
Function RndI(lower As Integer, higher As Integer) As Integer
    RndI = Int((higher - lower + 1) * Rnd + lower)
End Function

Private Sub FILL_LINEAR()
    Dim bCONST(0 To 255) As Byte
    For M = 0 To 255
        bCONST(M) = M
        State(M) = bCONST(M)
    Next M
End Sub





