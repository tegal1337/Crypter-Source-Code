Attribute VB_Name = "Module2"
Option Explicit
Private i As Integer
Private j As Integer
Private K As Integer
Private a As Byte
Private b As Byte
Dim M As Integer
Private L As Long
Private RC4KEY(255) As Byte
Private ADDTABLE(255, 255) As Byte
Dim State(0 To 255) As Byte

Function mXor(i As Integer, j As Integer) As Integer
    
    If i = j Then
        mXor = j
    Else
        mXor = i Xor j
    End If
    
End Function
Sub Swap(ByRef a As Integer, ByRef b As Integer)
    Dim T As Integer
    T = a
    a = b
    b = T
    
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

Public Sub RC4(byteArray() As Byte, Optional Password As String)
  If Password <> "" Then PREPARE_KEY Password
  For L = 0 To UBound(byteArray)
    i = ADDTABLE(i, 1)
    j = ADDTABLE(j, State(i))
    a = State(i): State(i) = State(j): State(j) = a
    b = State(ADDTABLE(State(i), State(j)))
    byteArray(L) = byteArray(L) Xor b
  Next L
End Sub

Private Sub PREPARE_KEY(sKEY As String)
  INITIALIZE_ADDTABLE
  FILL_LINEAR
  K = Len(sKEY)
  For i = 0 To K - 1
    b = Asc(Mid$(sKEY, i + 1, 1))
    For j = i To 255 Step K
      RC4KEY(j) = b
    Next j
  Next i
  j = 0
  For i = 0 To 255
    K = ADDTABLE(State(i), RC4KEY(i))
    j = ADDTABLE(j, K)
    b = State(i): State(i) = State(j): State(j) = b
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
Public Function XORDecryption(DataIn As String, CodeKey As String) As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    

    For lonDataPtr = 1 To (Len(DataIn) / 2)
        intXOrValue1 = Val("&H" & (Mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
   XORDecryption = strDataOut
End Function


