VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RC4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Event Progress(Percent As Long)

Private m_Key As String
Private m_sBox(0 To 255) As Integer
Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long
Public Function EncryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean
    On Error GoTo errorhandler
    If FileExist(InFile) = False Then
        EncryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True And Overwrite = False Then
        EncryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    Call EncryptByte(Buffer(), Key)
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , Buffer()
    Close #FileO
    EncryptFile = True
    Exit Function

errorhandler:
    EncryptFile = False
End Function
Public Function DecryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean
    On Error GoTo errorhandler
    If FileExist(InFile) = False Then
        DecryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True Then
        DecryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    Call DecryptByte(Buffer(), Key)
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , Buffer()
    Close #FileO
    DecryptFile = True
    Exit Function

errorhandler:
    DecryptFile = False
End Function

Public Sub DecryptByte(byteArray() As Byte, Optional Key As String)
    Call EncryptByte(byteArray(), Key)
End Sub
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputInHex As Boolean) As String
    Dim byteArray() As Byte
    byteArray() = StrConv(Text, vbFromUnicode)
    Call EncryptByte(byteArray(), Key)
    EncryptString = StrConv(byteArray(), vbUnicode)
    If OutputInHex = True Then EncryptString = EnHex(EncryptString)
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextInHex As Boolean) As String
    Dim byteArray() As Byte
    If IsTextInHex = True Then Text = DeHex(Text)
    byteArray() = StrConv(Text, vbFromUnicode)
    Call DecryptByte(byteArray(), Key)
    DecryptString = StrConv(byteArray(), vbUnicode)
End Function
Public Sub EncryptByte(byteArray() As Byte, Optional Key As String)
Dim I As Long, J As Long, Temp As Byte, Offset As Long, OrigLen As Long, CipherLen As Long, CurrPercent As Long, NextPercent As Long, sBox(0 To 255) As Integer

If (Len(Key) > 0) Then Me.Key = Key
Call CopyMem(sBox(0), m_sBox(0), 512)
OrigLen = UBound(byteArray) + 1
CipherLen = OrigLen

For Offset = 0 To (OrigLen - 1)
    I = (I + 1) Mod 256
    J = (J + sBox(I)) Mod 256
    Temp = sBox(I)
    sBox(I) = sBox(J)
    sBox(J) = Temp
    byteArray(Offset) = byteArray(Offset) Xor (sBox((sBox(I) + sBox(J)) Mod 256))
    If (Offset >= NextPercent) Then
        CurrPercent = Int((Offset / CipherLen) * 100)
        NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
        RaiseEvent Progress(CurrPercent)
    End If
Next
If (CurrPercent <> 100) Then RaiseEvent Progress(100)
End Sub
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
Private Sub Append(ByRef StringData As String, Optional Length As Long)
    Dim DataLength As Long
    If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 1024
        ReDim Preserve byteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub
Private Function DeHex(Data As String) As String
    Dim iCount As Double
    Reset
    For iCount = 1 To Len(Data) Step 2
        Append Chr$(Val("&H" & Mid$(Data, iCount, 2)))
    Next
    DeHex = GData
    Reset
End Function
Private Function EnHex(Data As String) As String
    Dim iCount As Double, sTemp As String
    Reset
    For iCount = 1 To Len(Data)
        sTemp = Hex$(Asc(Mid$(Data, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        Append sTemp
    Next
    EnHex = GData
    Reset
End Function
Private Function FileExist(FileName As String) As Boolean
    On Error GoTo errorhandler
    Call FileLen(FileName)
    FileExist = True
    Exit Function

errorhandler:
    FileExist = False
End Function
Private Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Public Property Let Key(New_Value As String)
    Dim A As Long, B As Long, Temp As Byte, Key() As Byte, KeyLen As Long
    If (m_Key = New_Value) Then Exit Property
    m_Key = New_Value
    Key() = StrConv(m_Key, vbFromUnicode)
    KeyLen = Len(m_Key)
    For A = 0 To 255
        m_sBox(A) = A
    Next A
    For A = 0 To 255
        B = (B + m_sBox(A) + Key(A Mod KeyLen)) Mod 256
        Temp = m_sBox(A)
        m_sBox(A) = m_sBox(B)
        m_sBox(B) = Temp
    Next
End Property

