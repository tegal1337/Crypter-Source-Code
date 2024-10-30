Attribute VB_Name = "Module4"
'Begin Code
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private mlngParameters() As Long
Private mlngAddress As Long
Private mbytCode() As Byte
Private mlngCP As Long
Private Type xbyte
    arr() As Byte
End Type
Public Function CallAPI(libName As String, funcName As String, ParamArray FuncParams()) As Long
    Dim arr() As Variant
    Dim i As Long
    
    arr() = FuncParams()
    CallAPI = CallRemote(libName, funcName, arr())
    For i = LBound(FuncParams()) To UBound(FuncParams())
        FuncParams(i) = arr(i)
    Next i
End Function
Private Function CallRemote(libName As String, funcName As String, FuncParams() As Variant) As Long
    Dim i As Integer
    Dim wasString() As Boolean
    Dim ParamsNull As Boolean
    Dim lb As Long
    Dim X() As xbyte
    
    ReDim mlngParameters(0)
    ReDim mbytCode(0)
    mlngAddress = CLng("0")
    
    If UBound(FuncParams()) = -1 Then
        ParamsNull = True
        GoTo ParamsNull
    End If
    
    On Error GoTo 0
        
    ReDim wasString(UBound(FuncParams()))
    For i = LBound(FuncParams()) To UBound(FuncParams())
        wasString(i) = False
        If VarType(FuncParams(i)) = vbString Then
            ReDim Preserve X(i)
            X(i).arr = StrConv(FuncParams(i), vbFromUnicode) & vbNullChar
            FuncParams(i) = VarPtr(X(i).arr(CLng("0")))
            wasString(i) = True
        End If
    Next i
        
ParamsNull:
    lb = LoadLibrary(ByVal libName)
    If lb = CLng("0") Then Exit Function

    mlngAddress = GetProcAddress(lb, ByVal funcName)
    If mlngAddress = CLng("0") Then
        FreeLibrary lb
        Exit Function
    End If
    
    ReDim mlngParameters(UBound(FuncParams) + 1)
    For i = CLng("1") To UBound(mlngParameters)
        mlngParameters(i) = CLng(FuncParams(i - CLng("1")))
    Next i
    CallRemote = CallWindowProc(PrepareCode, CLng("0"), CLng("0"), CLng("0"), CLng("0"))
    FreeLibrary lb
    If ParamsNull Then Exit Function
    For i = LBound(FuncParams()) To UBound(FuncParams())
        If wasString(i) Then
            FuncParams(i) = StrConv(X(i).arr(), vbUnicode)
        End If
    Next i
End Function
Private Function PrepareCode() As Long
    Dim lngX As Long
    Dim codeStart As Long
    ReDim mbytCode(18 + 32 + 6 * UBound(mlngParameters))
    codeStart = GetAlignedCodeStart(VarPtr(mbytCode(CLng("0"))))
    mlngCP = codeStart - VarPtr(mbytCode(CLng("0")))
    For lngX = CLng("0") To mlngCP - CLng("1")
        mbytCode(lngX) = &HCC
    Next
    AddByteToCode &H58
    AddByteToCode &H59
    AddByteToCode &H59
    AddByteToCode &H59
    AddByteToCode &H59
    AddByteToCode &H50
    For lngX = UBound(mlngParameters) To CLng("1") Step -CLng("1")
        AddByteToCode &H68
        AddLongToCode mlngParameters(lngX)
    Next
    AddCallToCode mlngAddress
    AddByteToCode &HC3
    AddByteToCode &HCC
    PrepareCode = codeStart
End Function
Private Sub AddCallToCode(lngAddress As Long)
    AddByteToCode &HE8
    AddLongToCode lngAddress - VarPtr(mbytCode(mlngCP)) - CLng("4")
End Sub
Private Sub AddLongToCode(lng As Long)
    Dim intX As Integer
    Dim byt(3) As Byte
    CopyMemory byt(CLng("0")), lng, CLng("4")
    For intX = CLng("0") To CLng("3")
        AddByteToCode byt(intX)
    Next
End Sub
Private Sub AddByteToCode(byt As Byte)
    mbytCode(mlngCP) = byt
    mlngCP = mlngCP + CLng("1")
End Sub
Private Function GetAlignedCodeStart(lngAddress As Long) As Long
    GetAlignedCodeStart = lngAddress + (CLng("15") - (lngAddress - 1) Mod 16)
    If (CLng("15") - (lngAddress - 1) Mod 16) = 0 Then GetAlignedCodeStart = GetAlignedCodeStart + 16
End Function



