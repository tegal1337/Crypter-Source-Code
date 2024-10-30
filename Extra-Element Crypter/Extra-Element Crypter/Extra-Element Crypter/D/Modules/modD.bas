Attribute VB_Name = "modD"
' :D
Option Explicit


Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Const IMAGE_DOS_SIGNATURE As Integer = &H5A4D
Const IMAGE_NT_SIGNATURE As Long = &H4550


Private Type IMAGE_DOS_HEADER
    e_magic                 As Integer
    e_cblp                  As Integer
    e_cp                    As Integer
    e_crlc                  As Integer
    e_cparhdr               As Integer
    e_minalloc              As Integer
    e_maxalloc              As Integer
    e_ss                    As Integer
    e_sp                    As Integer
    e_csum                  As Integer
    e_ip                    As Integer
    e_cs                    As Integer
    e_lfarlc                As Integer
    e_onvo                  As Integer
    e_res(0 To 3)           As Integer
    e_oemid                 As Integer
    e_oeminfo               As Integer
    e_res2(0 To 9)          As Integer
    e_lfanew                As Long
End Type


Private Type IMAGE_FILE_HEADER
    Machine                 As Integer
    NumberOfSections        As Integer
    TimeDataStamp           As Long
    PointerToSymbolTable    As Long
    NumberOfSymbols         As Long
    SizeOfOptionalHeader    As Integer
    Characteristics         As Integer
End Type


Private Type IMAGE_DATA_DIRECTORY
  VirtualAddress As Long
  isize As Long
End Type


Private Type IMAGE_OPTIONAL_HEADER32
    Magic                   As Integer
    MajorLinkerVersion      As Byte
    MinorLinkerVersion      As Byte
    SizeOfCode              As Long
    SizeOfInitalizedData    As Long
    SizeOfUninitalizedData  As Long
    AddressOfEntryPoint     As Long
    BaseOfCode              As Long
    BaseOfData              As Long
    ImageBase               As Long
    SectionAlignment        As Long
    FileAlignment           As Long
    MajorOperatingSystemVer As Integer
    MinorOperatingSystemVer As Integer
    MajorImageVersion       As Integer
    MinorImageVersion       As Integer
    MajorSubsystemVersion   As Integer
    MinorSubsystemVersion   As Integer
    Reserved1               As Long
    SizeOfImage             As Long
    SizeOfHeaders           As Long
    CheckSum                As Long
    Subsystem               As Integer
    DllCharacteristics      As Integer
    SizeOfStackReserve      As Long
    SizeOfStackCommit       As Long
    SizeOfHeapReserve       As Long
    SizeOfHeapCommit        As Long
    LoaerFlags              As Long
    NumberOfRvaAndSizes     As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type


Private Type IMAGE_SECTION_HEADER
    Name As String * 8
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
End Type


Private Type IMAGE_NT_HEADERS
    Signature As Long
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER32
End Type

Public Function GetSettings(ByVal szTargetSectionName As String) As String
Dim MZHeader As IMAGE_DOS_HEADER
Dim PEHeader As IMAGE_NT_HEADERS
Dim Section As IMAGE_SECTION_HEADER

Dim pMe As Long, pSection As Long
Dim i As Integer

If Len(szTargetSectionName) < 1 Then Exit Function

If Len(szTargetSectionName) > 8 Then szTargetSectionName = Left$(szTargetSectionName, 8)

pMe = GetModuleHandle(vbNullString)

If pMe Then
    CopyMemory MZHeader, ByVal pMe, Len(MZHeader)

    If MZHeader.e_magic = IMAGE_DOS_SIGNATURE Then
        CopyMemory PEHeader, ByVal pMe + MZHeader.e_lfanew, Len(PEHeader)

        If PEHeader.Signature = IMAGE_NT_SIGNATURE Then
            pSection = pMe + MZHeader.e_lfanew + 24 + PEHeader.FileHeader.SizeOfOptionalHeader

            For i = 0 To PEHeader.FileHeader.NumberOfSections - 1
                CopyMemory Section, ByVal pSection, Len(Section)

                If Left$(Section.Name, Len(szTargetSectionName)) = szTargetSectionName Then
                    GetSettings = String(Section.VirtualSize, Chr$(0))
                    CopyMemory ByVal GetSettings, ByVal pMe + Section.VirtualAddress, Section.VirtualSize
                    Exit For
                End If

                pSection = pSection + Len(Section)
            Next i
        End If
    End If
End If
End Function

