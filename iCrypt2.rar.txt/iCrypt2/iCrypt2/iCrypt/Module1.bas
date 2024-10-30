Attribute VB_Name = "mod_iCrypt"
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Function vbWriteByteFile(ByVal sFileName As String, lpByte() As Byte) As Boolean
    Dim fhFile As Integer
    fhFile = FreeFile
    Open sFileName For Binary As #fhFile
    Put #fhFile, , lpByte()
    Close #fhFile
End Function
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
Public Function Rand(ByVal Low As Long, _
                     ByVal High As Long) As Long
    Dim I1 As Long
    
    Randomize timeGetTime
    
  Rand = Int((High - Low + 1) * Rnd) + Low
End Function
Public Function RandString(lpszLen As Long) As String
Dim I1 As Long, i2 As Long
Dim lpCounter As Long, s1 As String
For lpCounter = 1 To lpszLen
    I1 = Rand(1, 3)
    Select Case I1
        Case 1
            s1 = s1 & Chr(Rand(65, 65 + 25))
        Case 2
            s1 = s1 & Chr(Rand(97, 97 + 25))
        Case 3
            s1 = s1 & CStr(Rand(0, 9))
    End Select
Next
RandString = s1
End Function

