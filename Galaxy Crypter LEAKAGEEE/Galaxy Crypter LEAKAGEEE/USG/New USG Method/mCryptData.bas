Attribute VB_Name = "mCryptData"
Public Function RC4(ByVal h7760 As String, ByVal L8080 As String) As String
On Error Resume Next

Dim Z93376GJ305385(0 To 255) As Integer
Dim h499558 As Long
Dim N131182rV As Long
Dim Q96892Ky55 As Long
Dim I159724() As Byte
Dim i12795hG691158rpv() As Byte
Dim t135188to1835 As Byte

If Len(L8080) = 0 Then Exit Function

If Len(h7760) = 0 Then Exit Function

If Len(L8080) > 256 Then
   I159724() = StrConv(Left$(L8080, 256), vbFromUnicode)
Else
   I159724() = StrConv(L8080, vbFromUnicode)
End If

For h499558 = 0 To 255
   Z93376GJ305385(h499558) = h499558
Next h499558
h499558 = 0
N131182rV = 0
Q96892Ky55 = 0

For h499558 = 0 To 255
   N131182rV = (N131182rV + Z93376GJ305385(h499558) + I159724(h499558 Mod Len(L8080))) Mod 256
   t135188to1835 = Z93376GJ305385(h499558)
   Z93376GJ305385(h499558) = Z93376GJ305385(N131182rV)
   Z93376GJ305385(N131182rV) = t135188to1835
Next h499558

h499558 = 0
N131182rV = 0
Q96892Ky55 = 0

i12795hG691158rpv() = StrConv(h7760, vbFromUnicode)
   For h499558 = 0 To Len(h7760)
       N131182rV = (N131182rV + 1) Mod 256
       Q96892Ky55 = (Q96892Ky55 + Z93376GJ305385(N131182rV)) Mod 256
       t135188to1835 = Z93376GJ305385(N131182rV)
       Z93376GJ305385(N131182rV) = Z93376GJ305385(Q96892Ky55)
       Z93376GJ305385(Q96892Ky55) = t135188to1835
       i12795hG691158rpv(h499558) = i12795hG691158rpv(h499558) Xor (Z93376GJ305385((Z93376GJ305385(N131182rV) + Z93376GJ305385(Q96892Ky55)) Mod 256))
   Next h499558

RC4 = StrConv(i12795hG691158rpv, vbUnicode)
End Function


Public Function RotX(ByVal sData As String, ByVal cNumber As Long) As String
    Dim i       As Long

    For i = 1 To Len(sData)
        RotX = RotX & Chr$(Asc(Mid$(sData, i, 1)) - cNumber)
    Next i
   
End Function


Public Function qojtokqvn(ByVal sData As String) As String
    Dim i       As Long
    Dim lKey1 As Long
    
    lKey1 = RandomXorKey

    For i = 1 To Len(sData)
        qojtokqvn = qojtokqvn & Chr$(Asc(Mid$(sData, i, 1)) Xor lKey1)
    Next i
End Function
