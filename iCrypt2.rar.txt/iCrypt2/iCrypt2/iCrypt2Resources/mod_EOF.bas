Attribute VB_Name = "mod_EOF"
Option Explicit

'========Main staff for any API code :)===========
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)
'==========PE staff==============
Private Enum ImageSignatureTypes
    IMAGE_DOS_SIGNATURE = &H5A4D     ''\\ MZ
    IMAGE_OS2_SIGNATURE = &H454E     ''\\ NE
    IMAGE_OS2_SIGNATURE_LE = &H454C  ''\\ LE
    IMAGE_VXD_SIGNATURE = &H454C     ''\\ LE
    IMAGE_NT_SIGNATURE = &H4550      ''\\ PE00
End Enum

Private Type IMAGE_DOS_HEADER
    e_magic As Integer        ' Magic number
    e_cblp As Integer         ' Bytes on last page of file
    e_cp As Integer           ' Pages in file
    e_crlc As Integer         ' Relocations
    e_cparhdr As Integer      ' Size of header in paragraphs
    e_minalloc As Integer     ' Minimum extra paragraphs needed
    e_maxalloc As Integer     ' Maximum extra paragraphs needed
    e_ss As Integer           ' Initial (relative) SS value
    e_sp As Integer           ' Initial SP value
    e_csum As Integer         ' Checksum
    e_ip As Integer           ' Initial IP value
    e_cs As Integer           ' Initial (relative) CS value
    e_lfarlc As Integer       ' File address of relocation table
    e_ovno As Integer         ' Overlay number
    e_res(0 To 3) As Integer  ' Reserved words
    e_oemid As Integer        ' OEM identifier (for e_oeminfo)
    e_oeminfo As Integer      ' OEM information; e_oemid specific
    e_res2(0 To 9) As Integer ' Reserved words
    e_lfanew As Long          ' File address of new exe header
End Type

' MSDOS File header
Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    characteristics As Integer
End Type

' Directory format.
Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

' Optional header format.
Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16

Private Type IMAGE_OPTIONAL_HEADER
    ' Standard fields.
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
    DataDirectory(0 To IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_NT_HEADERS
    Signature As Long
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

' Section header
Const IMAGE_SIZEOF_SHORT_NAME = 8

Private Type IMAGE_SECTION_HEADER
   SecName As String * IMAGE_SIZEOF_SHORT_NAME
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

'=============Code================
Const OFFSET_4 = 4294967296#


Public Function EndOfFile(lpszFile As String) As Long
    Dim idh As IMAGE_DOS_HEADER
    Dim inh As IMAGE_NT_HEADERS
    Dim ish As IMAGE_SECTION_HEADER
    Dim ImageBase As Long, ret As Long, i As Long
    Dim addr As Long, lOffset As Long
    Dim lpMemory() As Byte, lpHigh As Long, Best_ISH As IMAGE_SECTION_HEADER
    lpMemory = LoadFile(lpszFile)
    CopyMemory idh, lpMemory(0), Len(idh)
    If idh.e_magic <> IMAGE_DOS_SIGNATURE Then
       MsgBox "MZ signature not found!", vbCritical, "File load error"
       Exit Function
    End If
    CopyMemory inh, lpMemory(idh.e_lfanew), Len(inh)
    If inh.Signature <> IMAGE_NT_SIGNATURE Then
       MsgBox "PE signature not found!", vbCritical, "File load error"
       Exit Function
    End If
    lOffset = LenB(inh) + idh.e_lfanew
    For i = 0 To inh.FileHeader.NumberOfSections - 1
        CopyMemory ish, lpMemory(lOffset + i * Len(ish)), Len(ish)
        If ish.PointerToRawData + ish.SizeOfRawData > lpHigh Then lpHigh = ish.PointerToRawData + ish.SizeOfRawData
    Next i
    EndOfFile = lpHigh
End Function
Public Function EndOfFileByte(lpszFile As String) As Byte()
    Dim idh As IMAGE_DOS_HEADER
    Dim inh As IMAGE_NT_HEADERS
    Dim ish As IMAGE_SECTION_HEADER
    Dim ImageBase As Long, ret As Long, i As Long
    Dim addr As Long, lOffset As Long
    Dim lpMemory() As Byte, lpHigh As Long, Best_ISH As IMAGE_SECTION_HEADER
    lpMemory = LoadFile(lpszFile)
    CopyMemory idh, lpMemory(0), Len(idh)
    If idh.e_magic <> IMAGE_DOS_SIGNATURE Then
       MsgBox "MZ signature not found!", vbCritical, "File load error"
       Exit Function
    End If
    CopyMemory inh, lpMemory(idh.e_lfanew), Len(inh)
    If inh.Signature <> IMAGE_NT_SIGNATURE Then
       MsgBox "PE signature not found!", vbCritical, "File load error"
       Exit Function
    End If
    lOffset = LenB(inh) + idh.e_lfanew
    For i = 0 To inh.FileHeader.NumberOfSections - 1
        CopyMemory ish, lpMemory(lOffset + i * Len(ish)), Len(ish)
        If ish.PointerToRawData + ish.SizeOfRawData > lpHigh Then lpHigh = ish.PointerToRawData + ish.SizeOfRawData
    Next i
    
    Dim b1() As Byte
    i = UBound(lpMemory) - lpHigh
    ReDim b1(0 To i)
    CopyMemory b1(0), lpMemory(lpHigh), i + 1
    EndOfFileByte = b1
End Function
Public Function vbWriteEOF(lpszPath As String, lpData() As Byte, lpEof() As Byte) As Long
    Dim i1 As Long
    i1 = FreeFile
    Open lpszPath For Binary As #i1
    Put #i1, , lpData()
    Put #i1, , lpEof()
    Close #i1
End Function

