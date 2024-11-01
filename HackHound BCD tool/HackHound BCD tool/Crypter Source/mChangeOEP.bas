Attribute VB_Name = "mChangeOEP"
'---------------------------------------------------------------------------------------
' Module      : mChangeOEP
' DateTime    : 02/12/2008 19:02
' Author      : Cobein
' Mail        : cobein27@hotmail.com
' WebPage     : http://www.advancevb.com.ar
' Purpose     : Change PE Entry Point
' Usage       : At your own risk
' Requirements: None
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
'
' Reference   : http://www.governmentsecurity.org/archive/t5873.html
'
' History     : 02/12/2008 First Cut....................................................
'               02/12/2008 Two Bug Fixed on code position and entry point...............
'---------------------------------------------------------------------------------------
Option Explicit

Private Const IMAGE_DOS_SIGNATURE           As Long = &H5A4D&
Private Const IMAGE_NT_SIGNATURE            As Long = &H4550&
Private Const IMAGE_NT_OPTIONAL_HDR32_MAGIC As Long = &H10B&

Private Const SIZE_DOS_HEADER               As Long = &H40
Private Const SIZE_NT_HEADERS               As Long = &HF8
Private Const SIZE_SECTION_HEADER           As Long = &H28

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
    characteristics             As Integer
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
    characteristics             As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)

Public Function ChangeOEPFromFile( _
       ByVal sSrcFile As String) As Boolean

    Dim bvData() As Byte
    
'    On Error GoTo ChangeOEPFromFile_Error

    Open sSrcFile For Binary Access Read As #2
    ReDim bvData(LOF(2) - 1)
    Get #2, , bvData()
    Close
    
    If ChangeOEPFromBytes(bvData) Then
        Open sSrcFile For Binary Access Write As #2
        Put #2, , bvData()
        Close
    End If

    ChangeOEPFromFile = True

    On Error GoTo 0
    Exit Function

ChangeOEPFromFile_Error:
End Function

Public Function ChangeOEPFromBytes( _
       ByRef bvData() As Byte) As Boolean
    
    Dim lOEP                    As Long
    Dim lCodePos                As Long
    Dim tIMAGE_DOS_HEADER       As IMAGE_DOS_HEADER
    Dim tIMAGE_NT_HEADERS       As IMAGE_NT_HEADERS
    Dim tIMAGE_SECTION_HEADER   As IMAGE_SECTION_HEADER

    CopyMemory tIMAGE_DOS_HEADER, bvData(0), SIZE_DOS_HEADER

    If Not tIMAGE_DOS_HEADER.e_magic = IMAGE_DOS_SIGNATURE Then
        Exit Function
    End If
    
    CopyMemory tIMAGE_NT_HEADERS, bvData(tIMAGE_DOS_HEADER.e_lfanew), SIZE_NT_HEADERS

    If Not tIMAGE_NT_HEADERS.Signature = IMAGE_NT_SIGNATURE Then
        Exit Function
    End If
    
    If Not tIMAGE_NT_HEADERS.OptionalHeader.Magic = IMAGE_NT_OPTIONAL_HDR32_MAGIC Then
        Exit Function
    End If
    
    CopyMemory tIMAGE_SECTION_HEADER, bvData(tIMAGE_DOS_HEADER.e_lfanew + SIZE_NT_HEADERS), SIZE_SECTION_HEADER
    
    lOEP = tIMAGE_NT_HEADERS.OptionalHeader.AddressOfEntryPoint + tIMAGE_NT_HEADERS.OptionalHeader.ImageBase

    lCodePos = tIMAGE_SECTION_HEADER.VirtualSize + tIMAGE_SECTION_HEADER.PointerToRawData
    
    'Debug.Print Hex(lCodePos), Hex(tIMAGE_SECTION_HEADER.VirtualSize), Hex(tIMAGE_SECTION_HEADER.PointerToRawData)
    
    tIMAGE_NT_HEADERS.OptionalHeader.AddressOfEntryPoint = lCodePos + tIMAGE_NT_HEADERS.OptionalHeader.BaseOfCode - tIMAGE_SECTION_HEADER.PointerToRawData
    
    CopyMemory bvData(tIMAGE_DOS_HEADER.e_lfanew), tIMAGE_NT_HEADERS, SIZE_NT_HEADERS
    
    tIMAGE_SECTION_HEADER.VirtualSize = tIMAGE_SECTION_HEADER.SizeOfRawData
    
    CopyMemory bvData(tIMAGE_DOS_HEADER.e_lfanew + SIZE_NT_HEADERS), tIMAGE_SECTION_HEADER, SIZE_SECTION_HEADER
    
    'Debug.Print Hex(lCodePos)
    bvData(lCodePos) = &H68 'Push
    CopyMemory bvData(lCodePos + 1), NewOEP, &H4 'OEP
    'Debug.Print Hex(lOEP)
    bvData(lCodePos + 5) = &HC3 'Ret
    
    ChangeOEPFromBytes = True
End Function



