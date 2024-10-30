Attribute VB_Name = "modFunc"
Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

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
