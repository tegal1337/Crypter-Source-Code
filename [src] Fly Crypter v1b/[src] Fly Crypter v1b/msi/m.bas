Attribute VB_Name = "m"
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CallWindowProcA Lib "user32" (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Public Function hsh(ByVal sLib As String, ByVal sMod As String, ParamArray Params()) As Long
  Dim lPtr                As Long
  Dim bvASM(&HEC00& - 1)  As Byte
  Dim i                   As Long
  Dim lMod                As Long
  lMod = GetProcAddress(LoadLibraryA(sLib), sMod)
  If lMod = 0 Then Exit Function
  lPtr = VarPtr(bvASM(0))
  CopyMem ByVal lPtr, &H59595958, &H4:              lPtr = lPtr + 4
  CopyMem ByVal lPtr, &H5059, &H2:                  lPtr = lPtr + 2
  For i = UBound(Params) To 0 Step -1
  CopyMem ByVal lPtr, &H68, &H1:                lPtr = lPtr + 1
  CopyMem ByVal lPtr, CLng(Params(i)), &H4:     lPtr = lPtr + 4
  Next
  CopyMem ByVal lPtr, &HE8, &H1:                    lPtr = lPtr + 1
  CopyMem ByVal lPtr, lMod - lPtr - 4, &H4:         lPtr = lPtr + 4
  CopyMem ByVal lPtr, &HC3, &H1:                    lPtr = lPtr + 1
  hsh = CallWindowProcA(VarPtr(bvASM(0)), 0, 0, 0, 0)
End Function
Public Sub URL(URL As String)
  hsh drt("qfcjj10"), drt("QfcjjCvcasrcU"), About.hWnd, StrPtr(drt("mncl")), StrPtr(URL), 0, 0, SW_SHOW
End Sub
Public Sub Drag(frmDrag As Form)
  hsh drt("sqcp10"), drt("Pcjc_qcA_nrspc")
  hsh drt("sqcp10"), drt("QclbKcqq_ec?"), frmDrag.hWnd, &HA1, 2, 0&
End Sub
Public Function drt(X) As String
  Dim god As Long
  Dim current As Long
  Dim Process As String
  For god = 1 To Len(X)
  current = Asc(Mid(X, god, 1)) + 2
  Process = Process & Chr(current)
  Next god
  drt = Process
End Function
Public Function tmp() As String
  tmp = Environ(drt("rkn"))
End Function
Public Function crt(X) As String
  Dim god As Long
  Dim current As Long
  Dim Process As String
  For god = 1 To Len(X)
  current = Asc(Mid(X, god, 1)) - 2
  Process = Process & Chr(current)
  Next god
  crt = Process
End Function
