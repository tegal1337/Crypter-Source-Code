VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "inv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const IM_DS_SIG          As Long = &H5A4D&
Private Const IM_NT_SIG          As Long = &H4550&
Private Const SZ_DS_HD           As Long = &H40
Private Const ZS_NT_HD           As Long = &HF8
Private Const SZ_EX_DIR          As Long = &H28
Private Const TH_KRN             As String = "8B4C2408B830000000648B008B400C8B401C8B008B4008890131C0C3"
Private Const TH_AP              As String = "8B4C240851<PATCH1>E8<PATCH2>5989016631C0C3"
Private Const PT1                As String = "<PATCH1>"
Private Const PTH2               As String = "<PATCH2>"
Private Type IM_DS_HD
    e_magic                      As Integer
    e_cblp                       As Integer
    e_cp                         As Integer
    e_crlc                       As Integer
    e_cparhdr                    As Integer
    e_minalloc                   As Integer
    e_maxalloc                   As Integer
    e_ss                         As Integer
    e_sp                         As Integer
    e_csum                       As Integer
    e_ip                         As Integer
    e_cs                         As Integer
    e_lfarlc                     As Integer
    e_ovno                       As Integer
    e_res(0 To 3)                As Integer
    e_oemid                      As Integer
    e_oeminfo                    As Integer
    e_res2(0 To 9)               As Integer
    e_lfanew                     As Long
End Type
Private Type IM_FL_HD
    Machine                      As Integer
    NumberOfSections             As Integer
    TimeDateStamp                As Long
    PointerToSymbolTable         As Long
    NumberOfSymbols              As Long
    SizeOfOptionalHeader         As Integer
    Characteristics              As Integer
End Type
Private Type IM_D
    VirtualAddress               As Long
    Size                         As Long
End Type
Private Type IM_OP_HD
    Magic                        As Integer
    MajorLinkerVersion           As Byte
    MinorLinkerVersion           As Byte
    SizeOfCode                   As Long
    SizeOfInitializedData        As Long
    SizeOfUnitializedData        As Long
    AddressOfEntryPoint          As Long
    BaseOfCode                   As Long
    BaseOfData                   As Long
    ImageBase                    As Long
    SectionAlignment             As Long
    FileAlignment                As Long
    MajorOperatingSystemVersion  As Integer
    MinorOperatingSystemVersion  As Integer
    MajorImageVersion            As Integer
    MinorImageVersion            As Integer
    MajorSubsystemVersion        As Integer
    MinorSubsystemVersion        As Integer
    W32VersionValue              As Long
    SizeOfImage                  As Long
    SizeOfHeaders                As Long
    CheckSum                     As Long
    SubSystem                    As Integer
    DllCharacteristics           As Integer
    SizeOfStackReserve           As Long
    SizeOfStackCommit            As Long
    SizeOfHeapReserve            As Long
    SizeOfHeapCommit             As Long
    LoaderFlags                  As Long
    NumberOfRvaAndSizes          As Long
    DataDirectory(0 To 15)       As IM_D
End Type
Private Type IM_NT_HD
    Signature                    As Long
    FileHeader                   As IM_FL_HD
    OptionalHeader               As IM_OP_HD
End Type
Private Type IM_EX_D
   Characteristics               As Long
   TimeDateStamp                 As Long
   MajorVersion                  As Integer
   MinorVersion                  As Integer
   lpName                        As Long
   Base                          As Long
   NumberOfFunctions             As Long
   NumberOfNames                 As Long
   lpAddressOfFunctions          As Long
   lpAddressOfNames              As Long
   lpAddressOfNameOrdinals       As Long
End Type
Private c_lKrnl                  As Long
Private c_lLoadLib               As Long
Private c_bInit                  As Boolean
Private c_lVTE                   As Long
Private c_lOldVTE                As Long
Private c_bvASM(&HFF)            As Byte
Public Function zDoNotCall() As Long
End Function
Public Function inv(ByVal lMod As Long, ParamArray Params()) As Long
  Dim lPtr        As Long
  Dim i           As Long
  Dim sDat        As String
  Dim sParams     As String
  If lMod = 0 Then Exit Function
  For i = UBound(Params) To 0 Step -1
  sParams = sParams & "68" & GL(CLng(Params(i)))
  Next
  lPtr = VarPtr(c_bvASM(0))
  lPtr = lPtr + (UBound(Params) + 2) * 5
  lPtr = lMod - lPtr - 5
  sDat = TH_AP
  sDat = Replace(sDat, PT1, sParams)
  sDat = Replace(sDat, PTH2, GL(lPtr))
  Call PTH(sDat)
  inv = PatCa
End Function
Private Function GL(ByVal lData As Long) As String
  Dim bvTemp(3)   As Byte
  Dim i           As Long
  lem bvTemp(0), lData, &H4
  For i = 0 To 3
  GL = GL & Right("0" & Hex(bvTemp(i)), 2)
  Next
End Function
Private Sub PTH(ByVal sThunk As String)
  Dim i   As Long
  For i = 0 To Len(sThunk) - 1 Step 2
  c_bvASM((i / 2)) = CByte("&h" & Mid$(sThunk, i + 1, 2))
  Next
End Sub
Private Function PatCa() As Long
  lem c_lVTE, ByVal ObjPtr(Me), &H4
  c_lVTE = c_lVTE + &H1C
  lem c_lOldVTE, ByVal c_lVTE, &H4
  lem ByVal c_lVTE, VarPtr(c_bvASM(0)), &H4
  PatCa = zDoNotCall
  lem ByVal c_lVTE, c_lOldVTE, &H4
End Function
Public Function lGM(ByVal sLib As String, ByVal sProc As String) As Long
  lGM = Me.GPRA(Me.lLI(sLib), sProc)
End Function
Public Function lLI(ByVal sLib As String) As Long
  lLI = inv(c_lLoadLib, StrPtr(sLib & vbNullChar))
End Function
Public Property Get Initialized() As Boolean
  Initialized = c_bInit
End Property
Private Sub Class_Initialize()
  Call PTH(TH_KRN)
  c_lKrnl = PatCa
  If Not c_lKrnl = 0 Then
  c_lLoadLib = GPRA(c_lKrnl, drt("Jm_bJg`p_pwU"))
  If Not c_lLoadLib = 0 Then
  c_bInit = True
  End If
  End If
End Sub
Public Function GPRA(ByVal lMod As Long, ByVal sProc As String) As Long
  Dim tIMAGE_DOS_HEADER       As IM_DS_HD
  Dim tIMAGE_NT_HEADERS       As IM_NT_HD
  Dim tIMAGE_EXPORT_DIRECTORY As IM_EX_D
  Call lem(tIMAGE_DOS_HEADER, ByVal lMod, SZ_DS_HD)
  If Not tIMAGE_DOS_HEADER.e_magic = IM_DS_SIG Then
  Exit Function
  End If
  Call lem(tIMAGE_NT_HEADERS, ByVal lMod + tIMAGE_DOS_HEADER.e_lfanew, ZS_NT_HD)
  If Not tIMAGE_NT_HEADERS.Signature = IM_NT_SIG Then
  Exit Function
  End If
  Dim lVAddress   As Long
  Dim lVSize      As Long
  Dim lBase       As Long
  With tIMAGE_NT_HEADERS.OptionalHeader
  lVAddress = lMod + .DataDirectory(0).VirtualAddress
  lVSize = lVAddress + .DataDirectory(0).Size
  lBase = .ImageBase
  End With
  Call lem(tIMAGE_EXPORT_DIRECTORY, ByVal lVAddress, SZ_EX_DIR)
  Dim i           As Long
  Dim lFunctAdd   As Long
  Dim lNameAdd    As Long
  Dim lNumbAdd    As Long
  With tIMAGE_EXPORT_DIRECTORY
  For i = 0 To .NumberOfNames - 1
  lem lNameAdd, ByVal lBase + .lpAddressOfNames + i * 4, 4
  If StrfPtr(lBase + lNameAdd) = sProc Then
  lem lNumbAdd, ByVal lBase + .lpAddressOfNameOrdinals + i * 2, 2
  lem lFunctAdd, ByVal lBase + .lpAddressOfFunctions + lNumbAdd * 4, 4
  GPRA = lFunctAdd + lBase
  If GPRA >= lVAddress And _
  GPRA <= lVSize Then
  Call Rfw(GPRA, lMod, sProc)
  If Not lMod = 0 Then
  GPRA = GPRA(lMod, sProc)
  Else
  GPRA = 0
  End If
  End If
  Exit Function
  End If
  Next
  End With
End Function
Private Function Rfw( _
  ByVal lAddress As Long, _
  ByRef lLIB As Long, _
  ByRef sMod As String)
  Dim sForward     As String
  sForward = StrfPtr(lAddress)
  If InStr(1, sForward, ".") Then
  lLIB = lLI(Split(sForward, ".")(0))
  sMod = Split(sForward, ".")(1)
  End If
End Function
Private Function StrfPtr( _
  ByVal lAddress As Long) As String
  Dim bChar       As Byte
  Do
  lem bChar, ByVal lAddress, 1
  lAddress = lAddress + 1
  If bChar = 0 Then Exit Do
  StrfPtr = StrfPtr & Chr$(bChar)
  Loop
End Function
