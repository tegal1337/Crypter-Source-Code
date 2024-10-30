Attribute VB_Name = "modFunc"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Dim cCrypt As CRijndael

Public Function lv_TimerCallBack(ByVal hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tButtons As lvButtons_H
    CopyMemory tButtons, GetProp(hwnd, "lv_ClassID"), &H4
    Call tButtons.TimerUpdate(GetProp(hwnd, "lv_TimerID"))
    CopyMemory tButtons, 0&, &H4
End Function

Public Function EncryptData(sData As String, sPass As String) As String
    Dim bData()         As Byte
    Dim bPass()         As Byte
    Set cCrypt = New CRijndael
    
    bData() = StrConv(sData, vbFromUnicode)
    bPass() = StrConv(sPass, vbFromUnicode)
    bData() = cCrypt.EncryptData(bData(), bPass())
    sData = StrConv(bData(), vbUnicode)
    sData = RC4(sData, sPass)
    
    EncryptData = sData
End Function

Private Function RC4(ByVal Expression As String, ByVal Password As String) As String
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

Public Function GenKey(Length As Integer) As String
    Dim sTemp           As String
    Dim sName           As String
    Dim iLength         As Integer
    Dim iStep           As Integer
    Dim iRnd            As Integer
    
    sTemp = "1234567890AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz"
    iLength = Len(sTemp)
    Randomize
    sName = ""
    
    For iStep = 1 To Length
        iRnd = Int((iLength * Rnd) + 1)
        sName = sName & Mid(sTemp, iRnd, 1)
    Next iStep
    
    GenKey = sName
End Function

Public Function GetEOFData(sFilePath As String) As String
    On Error GoTo ErrHandler
    Dim sFileBuffer                 As String
    Dim sEOFBuffer                  As String
    Dim lPos                        As Long
    
    Open sFilePath For Binary As #1
        sFileBuffer = Space(LOF(1))
        Get #1, , sFileBuffer
    Close #1
    
    lPos = InStr(1, StrReverse(sFileBuffer), GetNullBytes(30))
    sEOFBuffer = (Mid(StrReverse(sFileBuffer), 1, lPos - 1))
    GetEOFData = StrReverse(sEOFBuffer)
    Exit Function
    
ErrHandler:
        GetEOFData = vbNullString
End Function

Public Sub WriteEOFData(sFilePath As String, sEOFData As String)
    On Error Resume Next
    Dim sFile           As String
    Dim lFF             As Long
    
    lFF = FreeFile
    
    Open sFilePath For Binary As #lFF
        sFile = Space(LOF(lFF))
        Get #lFF, , sFile
    Close #lFF
    
    Kill sFilePath
    lFF = FreeFile
    
    Open sFilePath For Binary As #lFF
        Put #lFF, , sFile & sEOFData
    Close #lFF
End Sub

Private Function GetNullBytes(lNum) As String
    Dim sBuffer         As String
    Dim i               As Integer
    
    For i = 1 To lNum
        sBuffer = sBuffer & Chr(0)
    Next i
    
    GetNullBytes = sBuffer
End Function

Public Function AddTheData(sData As String, sSectionName As String) As Boolean
    Dim dwSettingsRVA As Long, dwSettingsRaw As Long
    
    dwSettingsRaw = AddSection(frmMain.dlgSave.FileName, sSectionName, Len(sData), &HC0000040, dwSettingsRVA, True)
    
    If dwSettingsRaw Then
        Open frmMain.dlgSave.FileName For Binary Access Write As #1
            Put #1, dwSettingsRaw + 1, sData
        Close #1
        
        AddTheData = True
    Else
        AddTheData = False
    End If
End Function

